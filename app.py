# =========================================================
# IMPORTS
# =========================================================
import os, re, csv, uuid, secrets, random, smtplib, shutil
from io import StringIO, BytesIO
from datetime import datetime, date, time, timedelta
from functools import wraps
from email.message import EmailMessage
from flask import Flask, render_template, request, redirect, url_for, flash, abort, Response, session, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_wtf.csrf import CSRFProtect
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from sqlalchemy import UniqueConstraint
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from services.dsa import merge_sort, binary_search_title, RegistrationHashIndex, WaitingQueue

load_dotenv()

# =========================================================
# APP CONFIGURATION
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-me-before-submission")
app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv("DATABASE_URL", "sqlite:///events.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["WTF_CSRF_TIME_LIMIT"] = 3600
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024  # Allow admin-uploaded campus videos up to 250 MB
app.config["SESSION_COOKIE_SECURE"] = os.getenv("COOKIE_SECURE", "0") == "1"

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USE_SSL'] = True
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
app.config['MAIL_DEFAULT_SENDER'] = os.getenv('MAIL_USERNAME')

OTP_MINUTES = 10
OTP_MAX_ATTEMPTS = 5
EXPORT_STATUSES = {"all", "absent", "present", "onsite", "completed"}
EXPORT_SORTS = {"name_asc", "name_desc", "section", "time_in"}

csrf = CSRFProtect(app)
db = SQLAlchemy(app)
login = LoginManager(app)
login.login_view = "landing"

# =========================================================
# DATABASE MODELS
# =========================================================
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(160), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False)
    
    student_number = db.Column(db.String(40), unique=True)
    year_level = db.Column(db.String(30))
    section = db.Column(db.String(20), index=True)
    program = db.Column(db.String(20))
    
    faculty_id = db.Column(db.String(40), unique=True)
    handled_sections = db.Column(db.String(300))
    must_change_password = db.Column(db.Boolean, default=False, nullable=False)
    
    otp_code = db.Column(db.String(6))
    otp_expires_at = db.Column(db.DateTime)
    otp_attempts = db.Column(db.Integer, nullable=False, default=0)
    
    preferences = db.Column(db.Text, default="{}")
    is_active = db.Column(db.Boolean, default=True)

    def set_password(self, raw): self.password_hash = generate_password_hash(raw)
    def check_password(self, raw): return check_password_hash(self.password_hash, raw)
    
    def get_prefs(self):
        try: 
            import json
            return json.loads(self.preferences)
        except: 
            return {}

class Event(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(160), nullable=False)
    description = db.Column(db.Text, default="")
    event_date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    venue = db.Column(db.String(160), nullable=False)
    program = db.Column(db.String(40), nullable=False, default="BSCS")
    capacity = db.Column(db.Integer, nullable=False, default=100)
    lunch_start = db.Column(db.Time)
    lunch_end = db.Column(db.Time)
    archived = db.Column(db.Boolean, default=False, nullable=False)
    featured = db.Column(db.Boolean, default=False, nullable=False)
    image_filename = db.Column(db.String(255))
    manual_status = db.Column(db.String(30), default="Auto", nullable=False)
    
    @property
    def confirmed_count(self): return Registration.query.filter_by(event_id=self.id, status="registered").count()
    @property
    def waiting_count(self): return WaitingList.query.filter_by(event_id=self.id).count()
    @property
    def total_expenses(self): return sum(x.amount for x in self.expenses)
    
    @property
    def status(self):
        if self.archived: return "Archived"
        
        today = date.today()
        if today > self.event_date: return "Ended"
        if today == self.event_date:
            now = datetime.now().time()
            if now >= self.end_time: return "Ended"
            if self.lunch_start and self.lunch_end and self.lunch_start <= now < self.lunch_end: return "Lunch Break"
            if now >= self.start_time: return "Ongoing"
            return "Not started yet"
        if today < self.event_date:
            if self.manual_status and self.manual_status != "Auto": return self.manual_status
            return "Not started yet"
        return "Not started yet"

class Registration(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    receipt_code = db.Column(db.String(50), unique=True, nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    status = db.Column(db.String(20), default="registered", nullable=False)
    time_in = db.Column(db.DateTime)
    time_out = db.Column(db.DateTime)
    student = db.relationship("User", backref="registrations")
    event = db.relationship("Event", backref=db.backref("registrations", lazy=True))
    __table_args__ = (UniqueConstraint("student_id", "event_id", name="uq_student_event"),)

class WaitingList(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    student = db.relationship("User")
    __table_args__ = (UniqueConstraint("student_id", "event_id", name="uq_waiting_student_event"),)

class Expense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    label = db.Column(db.String(160), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    notes = db.Column(db.String(400))
    event = db.relationship("Event", backref=db.backref("expenses", lazy=True))

class Announcement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(180), nullable=False)
    summary = db.Column(db.String(300), default="")
    body = db.Column(db.Text, nullable=False)
    image_filename = db.Column(db.String(255))
    video_filename = db.Column(db.String(255))
    published = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    photos = db.relationship("AnnouncementPhoto", backref="announcement", cascade="all, delete-orphan", order_by="AnnouncementPhoto.id", lazy=True)
    videos = db.relationship("AnnouncementVideo", backref="announcement", cascade="all, delete-orphan", order_by="AnnouncementVideo.id", lazy=True)

class AnnouncementPhoto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    announcement_id = db.Column(db.Integer, db.ForeignKey("announcement.id", ondelete="CASCADE"), nullable=False, index=True)
    image_filename = db.Column(db.String(255), nullable=False)
    caption = db.Column(db.String(240), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class AnnouncementVideo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    announcement_id = db.Column(db.Integer, db.ForeignKey("announcement.id", ondelete="CASCADE"), nullable=False, index=True)
    video_filename = db.Column(db.String(255), nullable=False)
    caption = db.Column(db.String(240), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class CampusHighlight(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(140), nullable=False)
    caption = db.Column(db.String(300), default="")
    image_filename = db.Column(db.String(255), nullable=False)
    sort_order = db.Column(db.Integer, default=0, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

class ActivityLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    actor = db.Column(db.String(160), default="anonymous")
    action = db.Column(db.String(100), nullable=False)
    details = db.Column(db.Text)
    level = db.Column(db.String(20), default="INFO")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

# =========================================================
# HELPER FUNCTIONS
# =========================================================
@login.user_loader
def load_user(uid): return db.session.get(User, int(uid))

def log(action, details="", level="INFO"):
    db.session.add(ActivityLog(actor=current_user.email if current_user.is_authenticated else "anonymous", action=action, details=details, level=level))
    db.session.commit()

def role_required(*roles):
    def deco(fn):
        @wraps(fn)
        @login_required
        def wrapper(*a, **kw):
            if current_user.role not in roles: abort(403)
            return fn(*a, **kw)
        return wrapper
    return deco

def parse_date(s): return datetime.strptime(s, "%Y-%m-%d").date()
def parse_time(s): return datetime.strptime(s, "%H:%M").time()

def handled_sections_of(user):
    return [x.strip() for x in (user.handled_sections or "").split(",") if x.strip()]

def csv_safe(value):
    text = "" if value is None else str(value)
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text

def send_otp_notification(user, otp):
    email_user = os.getenv("MAIL_USERNAME")
    email_pass = os.getenv("MAIL_PASSWORD")
    if not email_user or not email_pass:
        print("SYSTEM ERROR: MAIL_USERNAME or MAIL_PASSWORD is not set.")
        flash("The email service is not configured.", "error")
        return False

    msg = EmailMessage()
    msg.set_content(f"Your LSPU Verification Code is: {otp}\n\nThis code expires in {OTP_MINUTES} minutes.")
    msg["Subject"] = "LSPU Verification Code"
    msg["From"] = f"LSPU Event Management <{email_user}>"
    msg["To"] = user.email
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=15) as smtp:
            smtp.login(email_user, email_pass)
            smtp.send_message(msg)
        flash(f"A verification code has been sent to {user.email}.", "success")
        return True
    except Exception as e:
        print(f"SYSTEM ERROR: Email failed to send. Reason: {e}")
    flash("We could not send the verification email.", "error")
    return False

# =========================================================
# PUBLIC ROUTES
# =========================================================
@app.route("/")
def landing():
    featured_announcement = Announcement.query.filter_by(published=True).order_by(Announcement.created_at.desc()).first()
    recent_query = Announcement.query.filter_by(published=True)
    if featured_announcement:
        recent_query = recent_query.filter(Announcement.id != featured_announcement.id)
    announcements = recent_query.order_by(Announcement.created_at.desc()).paginate(
        page=max(request.args.get("page", 1, type=int), 1), per_page=4, error_out=False
    )
    return render_template("landing.html", featured_announcement=featured_announcement, announcements=announcements)

@app.route("/announcements/<int:item_id>")
def announcement_detail(item_id):
    item = db.get_or_404(Announcement, item_id)
    if not item.published:
        abort(404)
    media = ([{"kind": "photo", "filename": photo.image_filename, "caption": photo.caption, "id": photo.id, "created_at": photo.created_at} for photo in item.photos]
             + [{"kind": "video", "filename": video.video_filename, "caption": video.caption, "id": video.id, "created_at": video.created_at} for video in item.videos])
    media.sort(key=lambda entry: (entry["created_at"], entry["kind"], entry["id"]))
    media_pagination = max(request.args.get("media_page", 1, type=int), 1)
    per_page = 6
    media_total = len(media)
    media_pages = max((media_total + per_page - 1) // per_page, 1)
    media_pagination = min(media_pagination, media_pages)
    start = (media_pagination - 1) * per_page
    visible_media = media[start:start + per_page]
    return render_template("announcement_detail.html", item=item, gallery_media=visible_media,
                           media_page=media_pagination, media_pages=media_pages, media_total=media_total)

@app.route("/register/<role>", methods=["GET", "POST"])
def sign_up(role):
    if role != "student": abort(404)
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        
        if User.query.filter_by(email=email).first():
            flash("Email is already registered.", "error")
            return redirect(url_for("sign_up", role=role))
        if len(password) < 12 or password != confirm_password:
            flash("Passwords must match and be at least 12 characters.", "error")
            return redirect(url_for("sign_up", role=role))

        student_number = request.form.get("student_number", "").strip()
        if not re.fullmatch(r"\d{4}-\d{4}", student_number):
            flash("Student number must be in the format 0000-0000.", "error")
            return redirect(url_for("sign_up", role=role))
        if User.query.filter_by(student_number=student_number).first():
            flash("Student number already exists.", "error")
            return redirect(url_for("sign_up", role=role))

        otp = str(secrets.randbelow(900000) + 100000)
        session['pending_registration'] = {
            'name': name, 'email': email, 'password_hash': generate_password_hash(password),
            'role': role, 'otp': otp, 'otp_expires_at': (datetime.utcnow() + timedelta(minutes=OTP_MINUTES)).isoformat(),
            'attempts': 0, 'student_number': student_number,
            'section': request.form.get("section", "").strip().upper(),
            'program': request.form.get("program", "").strip().upper(),
            'year_level': request.form.get("year_level", "")
        }

        dummy_user = User(email=email)
        if not send_otp_notification(dummy_user, otp):
            session.pop('pending_registration', None)
            return redirect(url_for("sign_up", role=role))
        return redirect(url_for("verify_registration", role=role, email=email))
    return render_template("register.html", role=role)

@app.route("/verify-registration/<role>", methods=["GET", "POST"])
def verify_registration(role):
    email = request.args.get("email", "").strip().lower()
    pending = session.get('pending_registration')
    if not pending or pending['email'] != email:
        flash("Session expired or invalid. Please register again.", "error")
        return redirect(url_for("sign_up", role=role))

    if request.method == "POST":
        otp = request.form.get("otp", "").strip()
        if datetime.utcnow() > datetime.fromisoformat(pending['otp_expires_at']):
            session.pop('pending_registration', None)
            flash("Code expired. Please register again.", "error")
            return redirect(url_for("sign_up", role=role))

        pending['attempts'] += 1
        session['pending_registration'] = pending
        if pending['attempts'] > OTP_MAX_ATTEMPTS:
            session.pop('pending_registration', None)
            flash("Too many incorrect attempts. Please register again.", "error")
            return redirect(url_for("sign_up", role=role))

        if not secrets.compare_digest(otp, pending['otp']):
            flash("Invalid OTP.", "error")
        else:
            u = User(name=pending['name'], email=pending['email'], role="student")
            u.password_hash = pending['password_hash']
            u.student_number = pending['student_number']
            u.section = pending['section']
            u.program = pending['program']
            u.year_level = pending['year_level']
            db.session.add(u)
            db.session.commit()
            session.pop('pending_registration', None)
            log("ACCOUNT_CREATED", f"student: {pending['email']}")
            flash("Account verified and created successfully!", "success")
            return redirect(url_for("sign_in", role="student"))
    return render_template("verify_otp.html", role=role, email=email, is_registration=True)

# =========================================================
# AUTHENTICATION ROUTES
# =========================================================
@app.route("/login/<role>", methods=["GET", "POST"])
def sign_in(role):
    if role not in ("admin", "professor", "student"): abort(404)
    if current_user.is_authenticated:
        return redirect(url_for({"admin":"admin_home","professor":"professor_home","student":"student_home"}[role]))

    if request.method == "POST":
        email, password = request.form["email"].strip(), request.form["password"]
        if role == "professor":
            user = User.query.filter_by(faculty_id=email, role=role).first()
        else:
            user = User.query.filter_by(email=email.lower(), role=role).first()
            
        if user and user.check_password(password):
            login_user(user); log("LOGIN_SUCCESS", role)
            if role == "professor" and user.must_change_password:
                flash("Please change your temporary password immediately.", "warning")
                return redirect(url_for("change_password"))
            return redirect(url_for({"admin":"admin_home","professor":"professor_home","student":"student_home"}[role]))
        
        log("LOGIN_FAILED", f"{role} login failed for {email}", "WARNING")
        flash("Invalid credentials.", "error")
    return render_template("login.html", role=role)

@app.route("/forgot-password/<role>", methods=["GET", "POST"])
def forgot_password(role):
    if role != "student": abort(404)
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        user = User.query.filter_by(email=email, role=role).first()
        if user:
            otp = str(secrets.randbelow(900000) + 100000)
            user.otp_code = otp
            user.otp_expires_at = datetime.utcnow() + timedelta(minutes=OTP_MINUTES)
            user.otp_attempts = 0
            db.session.commit()
            if not send_otp_notification(user, otp):
                return redirect(url_for("forgot_password", role=role))
            return redirect(url_for("verify_otp", role=role, email=email))
        flash("If an account matches, a reset code can be generated.", "success")
        return redirect(url_for("forgot_password", role=role))
    return render_template("forgot_password.html", role=role)

@app.route("/verify-otp/<role>", methods=["GET", "POST"])
def verify_otp(role):
    email = request.args.get("email", "").strip().lower()
    if not email: return redirect(url_for("forgot_password", role=role))
    if request.method == "POST":
        otp = request.form.get("otp", "").strip()
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        user = User.query.filter_by(email=email, role=role).first()

        if not user or not user.otp_code or not user.otp_expires_at or user.otp_expires_at < datetime.utcnow():
            flash("Code is invalid or expired.", "error")
        elif user.otp_attempts >= OTP_MAX_ATTEMPTS:
            flash("Too many attempts.", "error")
        elif not secrets.compare_digest(otp, user.otp_code):
            user.otp_attempts += 1; db.session.commit(); flash("Invalid OTP.", "error")
        elif len(new_password) < 12 or new_password != confirm_password:
            flash("Passwords must match and be 12+ characters.", "error")
        else:
            user.set_password(new_password); user.otp_code = None; user.otp_expires_at = None; user.otp_attempts = 0
            db.session.commit(); log("PASSWORD_RESET", f"{role} account")
            flash("Password reset successful.", "success")
            return redirect(url_for("sign_in", role=role))
    return render_template("verify_otp.html", role=role, email=email)

@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if not current_user.check_password(current_password):
            flash("Current password is incorrect.", "error")
        elif len(new_password) < 12 or new_password != confirm_password:
            flash("New passwords must match and be 12+ characters.", "error")
        elif current_user.check_password(new_password):
            flash("Choose a different password.", "error")
        else:
            current_user.set_password(new_password)
            current_user.must_change_password = False
            db.session.commit(); log("PASSWORD_CHANGED", "User changed password")
            flash("Password changed successfully.", "success")
            return redirect(url_for({"admin": "admin_home", "professor": "professor_home", "student": "student_home"}[current_user.role]))
    return render_template("change_password.html")

# =========================================================
# USER SETTINGS ROUTES
# =========================================================
@app.route("/settings")
@login_required
def settings():
    return render_template("settings.html", prefs=current_user.get_prefs())

@app.route("/settings/profile", methods=["POST"])
@login_required
def settings_profile():
    email = request.form.get("email", "").strip().lower()
    if email and email != current_user.email:
        if User.query.filter_by(email=email).first():
            flash("Email is already in use by another account.", "error")
            return redirect(url_for("settings"))
        current_user.email = email
        db.session.commit()
        log("PROFILE_UPDATED", f"email updated to {email}")
        flash("Profile updated successfully.", "success")
    else:
        flash("No changes made to email.", "info")
    return redirect(url_for("settings"))

@app.route("/settings/preferences", methods=["POST"])
@login_required
def settings_preferences():
    import json
    prefs = current_user.get_prefs()
    prefs["notif_register"] = request.form.get("notif_register") == "on"
    prefs["notif_waitlist"] = request.form.get("notif_waitlist") == "on"
    prefs["notif_reminders"] = request.form.get("notif_reminders") == "on"
    current_user.preferences = json.dumps(prefs)
    db.session.commit()
    flash("Preferences saved successfully.", "success")
    return redirect(url_for("settings"))

@app.route("/settings/delete", methods=["POST"])
@login_required
def settings_delete():
    if current_user.role != "student": abort(403)
    active_regs = Registration.query.filter_by(student_id=current_user.id, status="registered").count()
    if active_regs > 0:
        flash("You cannot delete your account while you have active event registrations.", "error")
        return redirect(url_for("settings"))

    user = current_user
    logout_user()
    db.session.delete(user)
    db.session.commit()
    log("ACCOUNT_DELETED", f"student account deleted", "WARNING")
    flash("Your account has been deleted.", "info")
    return redirect(url_for("landing"))

@app.route("/logout")
@login_required
def sign_out():
    logout_user()
    return redirect(url_for("landing"))

# =========================================================
# ADMIN ROUTES
# =========================================================
@app.route("/admin")
@role_required("admin")
def admin_home():
    events = Event.query.filter_by(archived=False).order_by(Event.event_date).all()
    past = Event.query.filter_by(archived=True).order_by(Event.event_date.desc()).all()
    logs = ActivityLog.query.order_by(ActivityLog.created_at.desc()).limit(30).all()
    professors = User.query.filter_by(role="professor").order_by(User.name).all()
    students = User.query.filter_by(role="student").order_by(User.section, User.name).all()
    return render_template("admin.html", events=events, past=past, logs=logs, professors=professors, students=students,
        student_count=User.query.filter_by(role="student").count(),
        registration_count=Registration.query.filter_by(status="registered").count())

@app.route("/admin/campus-content", methods=["GET", "POST"])
@role_required("admin")
def admin_campus_content():
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        body = request.form.get("body", "").strip()
        summary = request.form.get("summary", "").strip()
        cover_image = request.files.get("image")
        cover_video = request.files.get("video")
        if not title or not body:
            flash("Announcement title and article content are required.", "error")
            return redirect(url_for("admin_campus_content"))
        if not ((cover_image and cover_image.filename) or (cover_video and cover_video.filename)):
            flash("Please choose a cover photo or featured video.", "error")
            return redirect(url_for("admin_campus_content"))

        upload_dir = os.path.join(app.root_path, "static", "uploads")
        os.makedirs(upload_dir, exist_ok=True)
        image_exts = {".jpg", ".jpeg", ".png", ".webp"}
        video_exts = {".mp4", ".webm", ".mov"}
        saved_files = []

        def save_upload(file_obj, allowed, prefix):
            original = secure_filename(file_obj.filename or "")
            ext = os.path.splitext(original)[1].lower()
            if ext not in allowed:
                raise ValueError("Images must be JPG, PNG, or WEBP. Videos must be MP4, WEBM, or MOV.")
            filename = f"{prefix}_{uuid.uuid4().hex}{ext}"
            file_obj.save(os.path.join(upload_dir, filename))
            saved_files.append(filename)
            return filename

        try:
            image_filename = save_upload(cover_image, image_exts, "campus") if cover_image and cover_image.filename else None
            video_filename = save_upload(cover_video, video_exts, "campus_video") if cover_video and cover_video.filename else None
            item = Announcement(title=title, summary=summary, body=body, image_filename=image_filename,
                                video_filename=video_filename, published=True)
            db.session.add(item)
            db.session.flush()
            caption = request.form.get("media_caption", "").strip()[:240]
            for photo in request.files.getlist("photos"):
                if photo and photo.filename:
                    filename = save_upload(photo, image_exts, "campus")
                    db.session.add(AnnouncementPhoto(announcement_id=item.id, image_filename=filename, caption=caption))
            for video in request.files.getlist("videos"):
                if video and video.filename:
                    filename = save_upload(video, video_exts, "campus_video")
                    db.session.add(AnnouncementVideo(announcement_id=item.id, video_filename=filename, caption=caption))
            db.session.commit()
            log("ANNOUNCEMENT_PUBLISHED", title)
            flash("Campus story published.", "success")
        except ValueError as exc:
            db.session.rollback()
            for filename in saved_files:
                path = os.path.join(upload_dir, filename)
                if os.path.isfile(path): os.remove(path)
            flash(str(exc), "error")
        except Exception:
            db.session.rollback()
            for filename in saved_files:
                path = os.path.join(upload_dir, filename)
                if os.path.isfile(path): os.remove(path)
            raise
        return redirect(url_for("admin_campus_content"))

    announcements = Announcement.query.order_by(Announcement.created_at.desc()).paginate(
        page=max(request.args.get("page", 1, type=int), 1), per_page=10, error_out=False
    )
    return render_template("admin_campus_content.html", announcements=announcements)

@app.route("/admin/announcements/<int:item_id>/edit", methods=["GET", "POST"])
@role_required("admin")
def edit_announcement(item_id):
    item = db.get_or_404(Announcement, item_id)
    if request.method == "POST":
        item.title = request.form.get("title", "").strip()
        item.summary = request.form.get("summary", "").strip()
        item.body = request.form.get("body", "").strip()
        if not item.title or not item.body:
            flash("Title and article content are required.", "error")
            return redirect(url_for("edit_announcement", item_id=item.id))
        upload_dir = os.path.join(app.root_path, "static", "uploads")
        os.makedirs(upload_dir, exist_ok=True)
        image_exts = {".jpg", ".jpeg", ".png", ".webp"}
        video_exts = {".mp4", ".webm", ".mov"}
        saved_files = []
        old_files = []
        try:
            cover = request.files.get("image")
            if cover and cover.filename:
                ext = os.path.splitext(secure_filename(cover.filename))[1].lower()
                if ext not in image_exts: raise ValueError("Cover images must be JPG, PNG, or WEBP.")
                filename = f"campus_{uuid.uuid4().hex}{ext}"
                cover.save(os.path.join(upload_dir, filename)); saved_files.append(filename)
                old = item.image_filename; item.image_filename = filename
                if old: old_files.append(old)
            cover_video = request.files.get("video")
            if cover_video and cover_video.filename:
                ext = os.path.splitext(secure_filename(cover_video.filename))[1].lower()
                if ext not in video_exts: raise ValueError("Featured videos must be MP4, WEBM, or MOV.")
                filename = f"campus_video_{uuid.uuid4().hex}{ext}"
                cover_video.save(os.path.join(upload_dir, filename)); saved_files.append(filename)
                old = item.video_filename; item.video_filename = filename
                if old: old_files.append(old)
            caption = request.form.get("media_caption", "").strip()[:240]
            for photo in request.files.getlist("photos"):
                if photo and photo.filename:
                    ext = os.path.splitext(secure_filename(photo.filename))[1].lower()
                    if ext not in image_exts: raise ValueError("Additional photos must be JPG, PNG, or WEBP.")
                    filename = f"campus_{uuid.uuid4().hex}{ext}"
                    photo.save(os.path.join(upload_dir, filename)); saved_files.append(filename)
                    db.session.add(AnnouncementPhoto(announcement_id=item.id, image_filename=filename, caption=caption))
            for video in request.files.getlist("videos"):
                if video and video.filename:
                    ext = os.path.splitext(secure_filename(video.filename))[1].lower()
                    if ext not in video_exts: raise ValueError("Additional videos must be MP4, WEBM, or MOV.")
                    filename = f"campus_video_{uuid.uuid4().hex}{ext}"
                    video.save(os.path.join(upload_dir, filename)); saved_files.append(filename)
                    db.session.add(AnnouncementVideo(announcement_id=item.id, video_filename=filename, caption=caption))
            db.session.commit()
            for filename in old_files:
                old_path = os.path.join(upload_dir, filename)
                if os.path.isfile(old_path): os.remove(old_path)
        except ValueError as exc:
            db.session.rollback()
            for filename in saved_files:
                path = os.path.join(upload_dir, filename)
                if os.path.isfile(path): os.remove(path)
            flash(str(exc), "error")
            return redirect(url_for("edit_announcement", item_id=item.id))
        except Exception:
            db.session.rollback()
            for filename in saved_files:
                path = os.path.join(upload_dir, filename)
                if os.path.isfile(path): os.remove(path)
            raise
        log("ANNOUNCEMENT_UPDATED", item.title)
        flash("Campus story updated.", "success")
        return redirect(url_for("admin_campus_content"))
    return render_template("announcement_edit.html", item=item)

@app.route("/admin/announcements/<int:item_id>/toggle", methods=["POST"])
@role_required("admin")
def toggle_announcement(item_id):
    item = db.get_or_404(Announcement, item_id)
    item.published = not item.published
    db.session.commit()
    log("ANNOUNCEMENT_PUBLISHED" if item.published else "ANNOUNCEMENT_UNPUBLISHED", item.title)
    flash("Campus story published." if item.published else "Campus story moved to drafts.", "success")
    return redirect(url_for("admin_campus_content"))

@app.route("/admin/announcements/<int:item_id>/photos/<int:photo_id>/delete", methods=["POST"])
@role_required("admin")
def delete_announcement_photo(item_id, photo_id):
    photo = AnnouncementPhoto.query.filter_by(id=photo_id, announcement_id=item_id).first_or_404()
    filename = photo.image_filename
    db.session.delete(photo)
    db.session.commit()
    path = os.path.join(app.root_path, "static", "uploads", filename)
    if os.path.isfile(path): os.remove(path)
    flash("Additional photo removed.", "success")
    return redirect(url_for("edit_announcement", item_id=item_id))

@app.route("/admin/announcements/<int:item_id>/videos/<int:video_id>/delete", methods=["POST"])
@role_required("admin")
def delete_announcement_video(item_id, video_id):
    video = AnnouncementVideo.query.filter_by(id=video_id, announcement_id=item_id).first_or_404()
    filename = video.video_filename
    db.session.delete(video)
    db.session.commit()
    path = os.path.join(app.root_path, "static", "uploads", filename)
    if os.path.isfile(path): os.remove(path)
    flash("Additional video removed.", "success")
    return redirect(url_for("edit_announcement", item_id=item_id))

@app.route("/admin/announcements/<int:item_id>/delete", methods=["POST"])
@role_required("admin")
def delete_announcement(item_id):
    item = db.get_or_404(Announcement, item_id)
    filenames = [item.image_filename, item.video_filename] + [photo.image_filename for photo in item.photos] + [video.video_filename for video in item.videos]
    db.session.delete(item)
    db.session.commit()
    for filename in filenames:
        if filename:
            path = os.path.join(app.root_path, "static", "uploads", filename)
            if os.path.isfile(path): os.remove(path)
    log("ANNOUNCEMENT_DELETED", str(item_id))
    flash("Campus story deleted.", "success")
    return redirect(url_for("admin_campus_content"))

@app.route("/admin/directory")
@role_required("admin")
def admin_directory():
    students = User.query.filter_by(role="student").order_by(User.section, User.name).all()
    professors = User.query.filter_by(role="professor").order_by(User.name).all()
    return render_template("admin_directory.html", students=students, professors=professors)

@app.route("/admin/logs")
@role_required("admin")
def system_logs():
    logs = ActivityLog.query.order_by(ActivityLog.created_at.desc()).limit(100).all()
    alerts = ActivityLog.query.filter(ActivityLog.level.in_(["WARNING", "ERROR"])).order_by(ActivityLog.created_at.desc()).limit(50).all()
    return render_template("system_logs.html", logs=logs, alerts=alerts)

@app.route("/admin/create-professor", methods=["POST"])
@role_required("admin")
def create_professor():
    name = request.form.get("name", "").strip()
    faculty_id = request.form.get("faculty_id", "").strip().upper()
    if not re.fullmatch(r"FAC-\d{4}-\d{4}", faculty_id):
        flash("Faculty ID must be in the format FAC-0000-0000 (e.g., FAC-2026-0001).", "error")
        return redirect(url_for("admin_home"))

    selected_sections = request.form.getlist("sections")
    sections = ", ".join(selected_sections)
    temp_password = request.form.get("temp_password", "")

    if not name or not faculty_id or len(temp_password) < 8:
        flash("Provide a name, Faculty ID, and an 8+ character temporary password.", "error")
    elif User.query.filter_by(faculty_id=faculty_id).first():
        flash("That Faculty ID already exists.", "error")
    else:
        u = User(name=name, email=f"{faculty_id.lower()}@lspu.edu.ph", role="professor", faculty_id=faculty_id, handled_sections=sections, must_change_password=True)
        u.set_password(temp_password)
        db.session.add(u); db.session.commit()
        log("PROFESSOR_CREATED", faculty_id)
        flash(f"Professor account created. Give them ID: {faculty_id} and the temporary password.", "success")
    return redirect(url_for("admin_home"))

# =========================================================
# EVENT MANAGEMENT ROUTES
# =========================================================
@app.route("/admin/events/new", methods=["GET","POST"])
@role_required("admin")
def event_new():
    if request.method=="POST":
        try:
            e=Event(title=request.form["title"].strip(),description=request.form.get("description","").strip(),event_date=parse_date(request.form["event_date"]),start_time=parse_time(request.form["start_time"]),end_time=parse_time(request.form["end_time"]),venue=request.form["venue"].strip(),program=request.form["program"],capacity=int(request.form["capacity"]))
            if not e.title or not e.venue or e.capacity<1 or e.end_time<=e.start_time: raise ValueError()
            db.session.add(e); db.session.commit(); flash("Event created.","success")
            return redirect(url_for("admin_home"))
        except (ValueError,KeyError): db.session.rollback(); flash("Check event fields.","error")
    return render_template("event_form.html",event=None)

@app.route("/admin/events/<int:eid>/edit", methods=["GET","POST"])
@role_required("admin")
def event_edit(eid):
    e=db.session.get(Event,eid) or abort(404)
    if request.method=="POST":
        try:
            e.title=request.form["title"].strip()
            e.description=request.form.get("description","").strip()
            e.event_date=parse_date(request.form["event_date"])
            e.start_time=parse_time(request.form["start_time"])
            e.end_time=parse_time(request.form["end_time"])
            e.venue=request.form["venue"].strip()
            e.program=request.form["program"]
            e.capacity=int(request.form["capacity"])
            if not e.title or not e.venue or e.capacity<max(1,e.confirmed_count) or e.end_time<=e.start_time: raise ValueError()
            image = request.files.get("event_image")
            if image and image.filename:
                ext = os.path.splitext(secure_filename(image.filename))[1].lower()
                if ext not in (".jpg", ".jpeg", ".png", ".webp"): raise ValueError("image")
                os.makedirs(os.path.join(app.static_folder, "uploads"), exist_ok=True)
                e.image_filename = uuid.uuid4().hex + ext
                image.save(os.path.join(app.static_folder, "uploads", e.image_filename))
            db.session.commit(); log("EVENT_EDITED",e.title); flash("Event updated.","success")
            return redirect(url_for("admin_home"))
        except (ValueError,KeyError):
            db.session.rollback(); flash("Invalid values.","error")
    return render_template("event_form.html",event=e)

@app.route("/admin/events/<int:eid>/feature", methods=["POST"])
@role_required("admin")
def event_feature(eid):
    event = db.session.get(Event, eid) or abort(404)
    if event.archived:
        flash("Archived events cannot be featured.", "error")
    else:
        Event.query.update({Event.featured: False})
        event.featured = True
        db.session.commit()
        log("EVENT_FEATURED", event.title)
        flash("Main-page featured event updated.", "success")
    return redirect(url_for("admin_home"))

@app.route("/admin/events/<int:eid>/unfeature", methods=["POST"])
@role_required("admin")
def event_unfeature(eid):
    event = db.session.get(Event, eid) or abort(404)
    event.featured = False
    db.session.commit()
    log("EVENT_UNFEATURED", event.title)
    flash("Featured event removed from the main page.", "success")
    return redirect(url_for("admin_home"))

@app.route("/admin/events/<int:eid>/status", methods=["POST"])
@role_required("admin")
def event_status(eid):
    event = db.session.get(Event, eid) or abort(404)
    new_status = request.form.get("status")
    allowed = ["Auto", "Not started yet", "Ongoing", "Lunch Break", "Ended"]
    if new_status in allowed:
        event.manual_status = new_status
        db.session.commit()
        log("EVENT_STATUS_CHANGED", f"{event.title} -> {new_status}")
        flash(f"Event status updated to {new_status}.", "success")
    else:
        flash("Invalid status.", "error")
    return redirect(url_for("admin_home"))

@app.route("/admin/events/<int:eid>/delete",methods=["POST"])
@role_required("admin")
def event_delete(eid):
    e=db.session.get(Event,eid) or abort(404)
    if (Registration.query.filter_by(event_id=e.id).count() or
            WaitingList.query.filter_by(event_id=e.id).count() or
            Expense.query.filter_by(event_id=e.id).count()):
        flash("Events with records cannot be deleted.", "error")
    else:
        title=e.title; db.session.delete(e); db.session.commit(); log("EVENT_DELETED",title,"WARNING"); flash("Event deleted.","success")
    return redirect(url_for("admin_home"))

@app.route("/admin/events/<int:eid>/archive",methods=["POST"])
@role_required("admin")
def event_archive(eid):
    e=db.session.get(Event,eid) or abort(404)
    if e.status!="Ended": flash("Only ended events can be archived.","error")
    else: e.archived=True; db.session.commit(); log("EVENT_ARCHIVED",e.title); flash("Event archived.","success")
    return redirect(url_for("admin_home"))

# =========================================================
# REGISTRATION & EXPENSE ROUTES
# =========================================================
@app.route("/admin/events/<int:eid>/registrations")
@role_required("admin")
def event_registrations(eid):
    event = db.session.get(Event, eid) or abort(404)
    term = request.args.get("q", "").strip().casefold()
    records = Registration.query.filter_by(event_id=eid, status="registered").all()
    if term:
        records = [r for r in records if term in r.student.name.casefold() or term in (r.student.student_number or "").casefold()]
    return render_template("registrations.html", event=event, registrations=records, term=term)

@app.route("/admin/events/<int:eid>/registrations.csv")
@role_required("admin")
def event_registrations_csv(eid):
    event = db.session.get(Event, eid) or abort(404)
    records = Registration.query.filter_by(event_id=eid, status="registered").all()
    out = StringIO(); writer = csv.writer(out)
    writer.writerow(["Receipt Code", "Student Number", "Name", "Email", "Time In", "Time Out"])
    for r in records:
        writer.writerow([r.receipt_code, r.student.student_number, r.student.name, r.student.email, r.time_in, r.time_out])
    return Response(out.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=event-{eid}-registrations.csv"})

@app.route("/admin/events/<int:eid>/expenses",methods=["GET","POST"])
@role_required("admin")
def expenses(eid):
    e=db.session.get(Event,eid) or abort(404)
    if request.method=="POST":
        try:
            label=request.form["label"].strip(); amount=float(request.form["amount"])
            if not label or amount<0: raise ValueError()
            db.session.add(Expense(event_id=e.id,label=label,amount=amount,notes=request.form.get("notes","")))
            db.session.commit(); flash("Expense added.","success")
            return redirect(url_for("expenses",eid=e.id))
        except ValueError: db.session.rollback(); flash("Invalid amount.","error")
    return render_template("expenses.html",event=e)

@app.route("/admin/expenses/<int:xid>/edit", methods=["GET", "POST"])
@role_required("admin")
def expense_edit(xid):
    expense = db.session.get(Expense, xid) or abort(404)
    if request.method == "POST":
        label = request.form.get("label", "").strip()
        notes = request.form.get("notes", "").strip()
        try:
            amount = float(request.form.get("amount", ""))
            if not label or amount < 0: raise ValueError()
            expense.label, expense.amount, expense.notes = label, amount, notes
            db.session.commit()
            flash("Expense updated.", "success")
            return redirect(url_for("expenses", eid=expense.event_id))
        except ValueError:
            db.session.rollback()
            flash("Invalid values.", "error")
    return render_template("expense_form.html", expense=expense)

@app.route("/admin/expenses/<int:xid>/delete", methods=["POST"])
@role_required("admin")
def expense_delete(xid):
    expense = db.session.get(Expense, xid) or abort(404)
    eid, label = expense.event_id, expense.label
    db.session.delete(expense); db.session.commit()
    flash("Expense deleted.", "success")
    return redirect(url_for("expenses", eid=eid))

# =========================================================
# DATABASE BACKUP & RESTORE
# =========================================================
@app.route("/admin/backup/download")
@role_required("admin")
def download_backup():
    db_path = os.path.join(app.instance_path, 'events.db')
    if not os.path.exists(db_path):
        flash("Database file not found.", "error")
        return redirect(url_for("admin_home"))
    log("DB_BACKUP_DOWNLOADED", "Admin downloaded a database backup.")
    return send_file(db_path, as_attachment=True, download_name="event_registration_backup.db")

@app.route("/admin/backup/restore", methods=["POST"])
@role_required("admin")
def restore_backup():
    file = request.files.get("backup_file")
    if not file or file.filename == '':
        flash("Please select a backup file to upload.", "error")
        return redirect(url_for("admin_home"))
    if not file.filename.endswith('.db'):
        flash("Invalid file type. Please upload a .db file.", "error")
        return redirect(url_for("admin_home"))
    db_path = os.path.join(app.instance_path, 'events.db')
    try:
        file.save(db_path)
        log("DB_BACKUP_RESTORED", "Admin restored the database from a backup.", "WARNING")
        flash("Database successfully restored! Please restart the application to apply changes.", "success")
    except Exception as e:
        flash(f"Error restoring database: {e}", "error")
    return redirect(url_for("admin_home"))

# =========================================================
# ATTENDANCE ROUTES
# =========================================================
@app.route("/attendance")
@role_required("admin", "professor")
def attendance_dashboard():
    term = request.args.get("q", "").strip().casefold()
    event_id = request.args.get("event_id", type=int)
    query = Registration.query.filter_by(status="registered")
    if event_id: query = query.filter_by(event_id=event_id)
    records = query.all()
    if current_user.role == "professor":
        sections = handled_sections_of(current_user)
        records = [r for r in records if r.student.section in sections]
    visible_event_ids = {r.event_id for r in records}
    if term:
        records = [r for r in records if term in r.receipt_code.casefold() or term in r.student.name.casefold() or term in (r.student.student_number or "").casefold()]
    records = merge_sort(records, key=lambda r: (r.student.section or "ZZZ", r.event.event_date, r.student.name.casefold()))
    events = Event.query.filter_by(archived=False).order_by(Event.event_date).all()
    if current_user.role == "professor":
        events = [e for e in events if e.id in visible_event_ids]
    return render_template("attendance.html", registrations=records, events=events, term=term, selected_event=event_id, is_admin=(current_user.role == "admin"))

@app.route("/attendance/<int:rid>/<action>", methods=["POST"])
@role_required("admin", "professor")
def attendance(rid, action):
    r = db.session.get(Registration, rid) or abort(404)
    if current_user.role == "professor":
        sections = handled_sections_of(current_user)
        if r.student.section not in sections: abort(403)
        
    manual_time = request.form.get("manual_time")
    now = datetime.now()
    if manual_time:
        try:
            parsed_time = datetime.strptime(manual_time, "%H:%M").time()
            now = datetime.combine(date.today(), parsed_time)
        except ValueError:
            flash("Invalid time format.", "error")
            return redirect(request.referrer or url_for("admin_home"))
    if action == "time-in" and not r.time_in:
        r.time_in = now; db.session.commit(); flash("Time-in recorded.", "success")
    elif action == "time-out" and r.time_in and not r.time_out:
        r.time_out = now; db.session.commit(); flash("Time-out recorded.", "success")
    else:
        flash("Invalid or duplicate action.", "error")
    return redirect(request.referrer or url_for("admin_home"))

# =========================================================
# PROFESSOR ROUTES
# =========================================================
@app.route("/professor")
@role_required("professor")
def professor_home():
    sections = handled_sections_of(current_user)
    students = User.query.filter_by(role="student").filter(User.section.in_(sections)).order_by(User.section, User.name).all() if sections else []
    events = Event.query.order_by(Event.event_date.desc(), Event.title).all()
    ongoing_events = sum(1 for e in events if e.status in ["Ongoing", "Lunch Break"])
    return render_template("professor.html", students=students, sections=sections, events=events, ongoing_events=ongoing_events)

@app.route("/professor/export.csv")
@role_required("professor")
def professor_export_csv():
    sections = handled_sections_of(current_user)
    event_id = request.args.get("event_id", type=int)
    section = request.args.get("section", "").strip().upper()
    status = request.args.get("status", "all")
    sort = request.args.get("sort", "name_asc")

    if section and section not in sections:
        flash("You can only export sections that you handle.", "error")
        return redirect(url_for("professor_home"))
        
    event = None
    if event_id:
        event = db.session.get(Event, event_id)
        if not event:
            flash("The selected event was not found.", "error")
            return redirect(url_for("professor_home"))

    query = Registration.query.filter_by(status="registered")
    if event_id: query = query.filter_by(event_id=event_id)
    allowed = [section] if section else sections
    records = [r for r in query.all() if r.student.section in allowed]

    if status == "absent":      records = [r for r in records if not r.time_in]
    elif status == "present":   records = [r for r in records if r.time_in]
    elif status == "onsite":    records = [r for r in records if r.time_in and not r.time_out]
    elif status == "completed": records = [r for r in records if r.time_in and r.time_out]

    if not records:
        flash("No records match those export options.", "error")
        return redirect(url_for("professor_home"))

    def state(r):
        if not r.time_in: return "Absent"
        return "Completed" if r.time_out else "Timed in"

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["Student Number", "Full Name", "Section", "Year Level", "Program", "Event", "Event Date", "Receipt Code", "Time In", "Time Out", "Attendance"])
    for r in records:
        ws.append([r.student.student_number, r.student.name, r.student.section,
                   r.student.year_level, r.student.program, r.event.title,
                   r.event.event_date, r.receipt_code, r.time_in, r.time_out, state(r)])
    for cell in ws[1]: cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        row[6].number_format = "yyyy-mm-dd"
        row[8].number_format = "yyyy-mm-dd hh:mm"
        row[9].number_format = "yyyy-mm-dd hh:mm"
    for col in ws.columns:
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = max(12, min(longest + 3, 45))
    buffer = BytesIO()
    wb.save(buffer)
    filename = f"attendance-{'event-' + str(event_id) if event_id else 'all-events'}-{date.today().isoformat()}.xlsx"
    return Response(buffer.getvalue(),
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename={filename}"})

@app.route("/professor/events")
@role_required("professor")
def professor_events():
    events = Event.query.filter_by(archived=False).all()
    term = request.args.get("q", "").strip()
    if term:
        t = term.casefold()
        events = [e for e in events if t in e.title.casefold()]
    return render_template("professor_events.html", events=events, term=term)

@app.route("/professor/events/<int:eid>/registrations")
@role_required("professor")
def professor_event_registrations(eid):
    event = db.session.get(Event, eid) or abort(404)
    sections = handled_sections_of(current_user)
    term = request.args.get("q", "").strip().casefold()
    records = Registration.query.filter_by(event_id=eid, status="registered").all()
    records = [r for r in records if r.student.section in sections]
    if term:
        records = [r for r in records if term in r.student.name.casefold()]
    return render_template("professor_event.html", event=event, registrations=records, term=term)

# =========================================================
# STUDENT ROUTES
# =========================================================
@app.route("/student")
@role_required("student")
def student_home():
    events=Event.query.filter_by(archived=False).all()
    term=request.args.get("q","").strip()
    status_order = {"Ongoing": 1, "Lunch Break": 1, "Not started yet": 2, "Ended": 3, "Archived": 4}
    events=merge_sort(events,key=lambda e:(status_order.get(e.status, 5), e.title.casefold()))
    if term:
        exact=binary_search_title(events,term)
        events=[e for e in events if (e.title.casefold()==term.casefold())] if exact else [e for e in events if term.casefold() in e.title.casefold()]
    return render_template("student.html",events=events,term=term)

@app.route("/student/events/<int:eid>")
@role_required("student")
def event_detail(eid):
    e=db.session.get(Event,eid) or abort(404)
    registration=Registration.query.filter_by(student_id=current_user.id,event_id=e.id,status="registered").first()
    waiting=WaitingList.query.filter_by(student_id=current_user.id,event_id=e.id).first()
    return render_template("event_detail.html",event=e,registration=registration,waiting=waiting)

@app.route("/student/events/<int:eid>/register", methods=["POST"])
@role_required("student")
def register(eid):
    e = db.session.get(Event, eid) or abort(404)
    if e.archived or e.status in ["Ended", "Lunch Break"]:
        flash("Registration is currently closed for this event.", "error")
        return redirect(url_for("event_detail", eid=e.id))
    if e.confirmed_count >= e.capacity:
        db.session.add(WaitingList(student_id=current_user.id, event_id=e.id)); db.session.commit(); flash("Event full; waitlisted.", "success")
    else:
        code = f"EVT-{e.id}-{current_user.id}-{uuid.uuid4().hex[:6]}"
        r = Registration(receipt_code=code, student_id=current_user.id, event_id=e.id)
        db.session.add(r); db.session.commit(); return redirect(url_for("receipt", rid=r.id))
    return redirect(url_for("event_detail", eid=e.id))

@app.route("/student/waitlist/<int:wid>/cancel", methods=["POST"])
@role_required("student")
def cancel_waitlist(wid):
    item = db.session.get(WaitingList, wid) or abort(404)
    if item.student_id != current_user.id: abort(403)
    db.session.delete(item); db.session.commit(); flash("Left waitlist.", "success")
    return redirect(url_for("event_detail", eid=item.event_id))

@app.route("/student/registrations/<int:rid>/cancel", methods=["POST"])
@role_required("student")
def cancel_registration(rid):
    registration = db.session.get(Registration, rid) or abort(404)
    if registration.student_id != current_user.id: abort(403)
    if registration.time_in or registration.time_out:
        flash("A registration with recorded attendance cannot be cancelled.", "error")
        return redirect(url_for("event_detail", eid=registration.event_id))
    event = registration.event
    if event.archived or event.status == "Ended":
        flash("Registration cannot be cancelled after the event has ended.", "error")
        return redirect(url_for("event_detail", eid=event.id))
    code = registration.receipt_code
    db.session.delete(registration)
    db.session.flush()
    waiting_rows = WaitingList.query.filter_by(event_id=event.id).order_by(WaitingList.created_at, WaitingList.id).all()
    queue = WaitingQueue()
    for row in waiting_rows: queue.enqueue(row.id)
    next_waiting_id = queue.dequeue()
    if next_waiting_id is not None:
        next_waiting = db.session.get(WaitingList, next_waiting_id)
        if next_waiting:
            promoted_student_id = next_waiting.student_id
            promoted_code = f"EVT-{event.id}-{promoted_student_id}-{datetime.now().strftime('%Y%m%d%H%M%S%f')[-8:]}"
            promoted = Registration(receipt_code=promoted_code, student_id=promoted_student_id, event_id=event.id)
            db.session.add(promoted)
            db.session.delete(next_waiting)
            db.session.commit()
            log("WAITLIST_PROMOTED", f"{promoted_code}; promoted after cancellation of {code}")
            flash("Registration cancelled. The next student on the waiting list has been promoted.", "success")
        else:
            db.session.commit()
            flash("Registration cancelled.", "success")
    else:
        db.session.commit()
        flash("Registration cancelled.", "success")
    log("REGISTRATION_CANCELLED", code, "WARNING")
    return redirect(url_for("student_home"))

@app.route("/student/receipt/<int:rid>")
@role_required("student")
def receipt(rid):
    r=db.session.get(Registration,rid) or abort(404)
    if r.student_id!=current_user.id: abort(403)
    return render_template("receipt.html",registration=r)

# =========================================================
# ERROR HANDLERS
# =========================================================
@app.errorhandler(403)
def forbidden(_): return render_template("error.html",code=403,message="You do not have permission to view this page."),403
@app.errorhandler(404)
def not_found(_): return render_template("error.html",code=404,message="The requested page or record was not found."),404

# =========================================================
# DATABASE INITIALIZATION & SEEDING
# =========================================================
with app.app_context():
    db.create_all()
    from sqlalchemy import inspect
    
    if "user" in inspect(db.engine).get_table_names():
        user_columns = {column["name"] for column in inspect(db.engine).get_columns("user")}
        with db.engine.begin() as connection:
            if "faculty_id" not in user_columns: connection.exec_driver_sql("ALTER TABLE user ADD COLUMN faculty_id VARCHAR(40)")
            if "must_change_password" not in user_columns: connection.exec_driver_sql("ALTER TABLE user ADD COLUMN must_change_password BOOLEAN NOT NULL DEFAULT 0")
            if "preferences" not in user_columns: connection.exec_driver_sql("ALTER TABLE user ADD COLUMN preferences TEXT DEFAULT '{}'")
            if "is_active" not in user_columns: connection.exec_driver_sql("ALTER TABLE user ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT 1")
                
    if "announcement" in inspect(db.engine).get_table_names():
        announcement_columns = {column["name"] for column in inspect(db.engine).get_columns("announcement")}
        if "video_filename" not in announcement_columns:
            with db.engine.begin() as connection:
                connection.exec_driver_sql("ALTER TABLE announcement ADD COLUMN video_filename VARCHAR(255)")

    if "event" in inspect(db.engine).get_table_names():
        event_columns = {column["name"] for column in inspect(db.engine).get_columns("event")}
        with db.engine.begin() as connection:
            if "manual_status" not in event_columns: connection.exec_driver_sql("ALTER TABLE event ADD COLUMN manual_status VARCHAR(30) NOT NULL DEFAULT 'Auto'")

    admin_email=os.getenv("ADMIN_EMAIL","admin@example.com").lower()
    admin_password=os.getenv("ADMIN_PASSWORD","ChangeThisAdminPassword123!")
    if not User.query.filter_by(email=admin_email).first():
        u=User(name="System Administrator",email=admin_email,role="admin")
        u.set_password(admin_password); db.session.add(u); db.session.commit()
        
    if Event.query.count()==0:
        today=date.today()
        db.session.add_all([
            Event(title="CCS Acquaintance Party",description="Meet students and faculty from the College of Computer Studies.",event_date=today,start_time=time(8),end_time=time(17),venue="Campus Gymnasium",program="BSCS",capacity=100,lunch_start=time(12),lunch_end=time(13)),
            Event(title="Programming Workshop",description="A hands-on introductory programming workshop.",event_date=today,start_time=time(9),end_time=time(15),venue="Computer Laboratory",program="BSIT",capacity=40,lunch_start=time(12),lunch_end=time(13))
        ]); db.session.commit()

if __name__=="__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "0") == "1")