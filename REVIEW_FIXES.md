=========================================================
LSPU EVENT REGISTRATION SYSTEM - REVIEW FIXES & CHANGELOG
=========================================================

1. DATABASE ARCHITECTURE & INTEGRITY
---------------------------------------------------------
- (Nov 3) Added an Auto-Migration script in app.py to automatically generate missing columns if the database is reset.
- (Nov 3) Added UniqueConstraint to the Registration table to prevent double-booking.
- (Nov 3) Established bridge tables (registration, waiting_list) to resolve Many-to-Many relationships.

2. AUTHENTICATION & USER ROLES
---------------------------------------------------------
- (Nov 3) Removed public professor registration. Admin now creates professor accounts.
- (Nov 3) Professors now log in using their Faculty ID instead of an email.
- (Nov 3) Implemented real Gmail SMTP OTP sending (10-minute expiration).
- (Nov 3) Added "Confirm Password" fields to Registration and Password Reset.
- (Nov 3) Forced new Professors to change their temporary password on first login.
- (Nov 8) Enforced strict ID formatting using Regex: Student Number must be "0000-0000" (numbers only), Faculty ID must be "FAC-0000-0000" (letters, numbers, dashes).

3. EVENT & REGISTRATION LOGIC
---------------------------------------------------------
- (Nov 3) Added manual Event Status override for Admin (Auto, Ongoing, Lunch Break, Ended).
- (Nov 3) Blocked student registration if the event is "Ended" or "Lunch Break".
- (Nov 3) Re-enabled FIFO Queue logic: Cancelling a registration automatically promotes the next student in the waiting list.
- (Nov 3) Ended events are now automatically sorted to the bottom of the Student Event Board using Merge Sort.
- (Nov 8) Fixed real-time status logic: Past events automatically force "Ended" status.

4. ADMIN & PROFESSOR FEATURES
---------------------------------------------------------
- (Nov 3) Created a "Logs" tab in the Admin navigation with a terminal-style UI.
- (Nov 3) Professors can now see which of their students registered for an event.
- (Nov 3) Admins can manually input time-in/time-out using a time picker.
- (Nov 3) Removed Professor permissions to modify attendance (View-only).
- (Nov 3) Fixed the Excel Export bug (now generates an actual filtered .xlsx file using openpyxl).
- (Nov 4) Added a Database Backup & Restore feature (Download/Upload .db file) to the Admin Dashboard.

5. FRONTEND & UI/UX FIXES
---------------------------------------------------------
- (Nov 3) Replaced section text inputs with checkboxes for Admin and dropdowns for Students.
- (Nov 3) Removed the "01 — OVERVIEW" text from Event Boards.
- (Nov 3) Fixed the Student Registration layout (centered the left-side blue panel).
- (Nov 4) Created a Role-Based Settings page (Profile, Security, Notifications, Preferences).
- (Nov 9) Added a Dark Mode toggle in Settings (saves to browser localStorage).
- (Nov 9) Fixed Dark Mode CSS contrast issues on tables, hover states, and portal cards.
- (Nov 9) Fixed the Landing Page featured card (centered text, removed overlapping).

6. DSA (DATA STRUCTURES & ALGORITHMS) IMPLEMENTED
---------------------------------------------------------
- (Nov 3) Hash Map (RegistrationHashIndex): O(1) instant duplicate registration checks.
- (Nov 3) Queue (WaitingQueue): FIFO waitlist promotion.
- (Nov 3) Merge Sort: O(N log N) sorting for events and dashboards.
- (Nov 3) Binary Search: O(log N) exact-title event search.