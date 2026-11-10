# LSPU Event Management System

A comprehensive, role-based campus event management platform built for Laguna State Polytechnic University (LSPU). This system streamlines event creation, student registration, waitlist management, attendance tracking, and financial reporting.

## Features

### Student Portal
- Discover Events: Browse upcoming and ongoing campus events.
- Registration: Reserve a seat with real-time capacity tracking.
- Waitlist (FIFO): Automatically joins a waitlist when an event is full; promotes the next student if someone cancels.
- Receipts: Generate and print official registration receipts with unique reference codes.
- Profile and Settings: Update email preferences and enable dark mode.

### Professor Portal
- Section Roster: View all students assigned to handled sections.
- Event Tracking: See which students from their sections registered for events.
- Attendance Export: Download comprehensive Excel (.xlsx) reports filtered by event, section, and attendance status.

### Admin Workspace
- Dashboard: High-level metrics on active events, total registrations, and system activity.
- Event Management: Create, edit, feature (spotlight), archive, and delete events.
- Manual Status Override: Force an event to "Ongoing," "Lunch Break," or "Ended" regardless of the system clock.
- Account Management: Create faculty accounts with temporary passwords, view directories, and edit user details.
- Attendance Desk: Record manual time-in and time-out for students.
- Financials: Track event expenses and view total recorded costs.
- System Logs: Terminal-style real-time activity feed with warning/error alerts.
- Database Maintenance: Download and restore .db backup files.

### Security and Auth
- Role-Based Access Control (RBAC) ensuring users only see their designated portals.
- Secure password hashing (Werkzeug).
- OTP (One-Time Password) email verification for student registration and password resets.
- CSRF protection on all forms.
- Session management with secure cookies.

---

## Technology Stack
- Backend: Python 3, Flask
- Database: SQLite (SQLAlchemy ORM)
- Authentication: Flask-Login
- Security: Flask-WTF (CSRF), Werkzeug Security
- Email: smtplib (Gmail SMTP)
- Excel Export: openpyxl
- Frontend: HTML5, Jinja2, Vanilla CSS, Vanilla JavaScript

---

## Data Structures and Algorithms (DSA) Showcase
This project intentionally implements custom algorithms to demonstrate core computer science concepts:
- Merge Sort (O(n log n)): Used to efficiently sort event lists by status/date and attendance records by section/name.
- Binary Search (O(log n)): Used for exact-title matching in the student event search.
- Hash Index (O(1)): Used to quickly check registration pairs (student_id, event_id) to prevent duplicates.
- FIFO Queue: Implemented using collections.deque to manage the waiting list and automatically promote the next student when a slot opens.

---

## Getting Started

### 1. Prerequisites
- Python 3.8+
- pip

### 2. Installation
Clone the repository and install the required dependencies:
```bash
git clone <your-repo-url>
cd LSPU-Event-Management-System
pip install -r requirements.txt

## Campus homepage content (added)

Administrators can open **Campus content** from the admin navigation or dashboard to:
- publish campus announcements/articles with a cover image, summary, and full text;
- add homepage campus highlight images for organizations, departments, facilities, and activities; and
- remove announcements or highlight images when they are no longer needed.

Uploaded images are stored in `static/uploads/`. Supported formats are JPG, JPEG, PNG, and WEBP. The homepage displays published announcements and campus highlights automatically. The new `announcement` and `campus_highlight` database tables are created automatically by `db.create_all()` when the app starts; existing event and account records are not intentionally changed.

## Campus Bulletin media

Administrators can publish a campus story with a cover photo, a featured video, or both. Additional photos and videos can be attached to the same story, with an optional shared caption for each upload batch. Supported image formats are JPG, PNG, and WEBP; supported video extensions are MP4, WEBM, and MOV. MP4 (H.264 video/AAC audio) is recommended for the broadest browser compatibility. The application accepts uploads up to 250 MB per request; the deployment server may have its own lower upload limit.

The public homepage shows the newest published story as the featured story and paginates the remaining announcements four at a time. Story galleries show six media items per page. Existing announcement records and photos are retained; startup creates the additional video table and adds the optional featured-video column to existing databases.
