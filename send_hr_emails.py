"""
send_hr_emails.py
-----------------
Reads HR/recruiter contacts from an Excel file and sends
personalized emails via Gmail SMTP with intelligent cold-email
optimizations.

Features:
- CLI Flags (e.g. `py send_hr_emails.py --dry-run 10`, `--send 50`, `--test`)
- Interactive Search & Filter Web UI for `report.html`
- Anti-Spam Human Jitter Delay & Rest Breaks
- First-Name Greeting Extractor
- Multi-Role & Company Name Cleaner
- Location Smart Matching
- Zero-Bounce Email Pre-Validator
- Incremental Crash-Safe Logging & Auto-Reconnect

SETUP
-----
1. Enable 2-Step Verification on your Google account.
2. Create an App Password: https://myaccount.google.com/apppasswords
3. Set environment variable:
   setx HR_MAIL_APP_PASSWORD "abcdefghijklmnop"
"""

import sys
import time
import smtplib
import ssl
import csv
import os
import re
import random
import argparse

from datetime import datetime, date

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import openpyxl


# =====================================================================
# CONFIG
# =====================================================================

EXCEL_FILE = "contacts.xlsx"
SHEET_NAME = None

SENDER_EMAIL = "gajendradhanoliya01@gmail.com"
APP_PASSWORD = os.getenv("HR_MAIL_APP_PASSWORD", "")

SENDER_NAME = "Gajendra Dhanoliya"
SENDER_PHONE = "+91-9109485566"

LINKEDIN_URL = "https://www.linkedin.com/in/gajendra-dhanoliya-813345359/"
GITHUB_URL = "https://github.com/dhanoliya-ji"

# Subject line rotation
SUBJECT_TEMPLATES = [
    "Application for {role} Opportunity at {company}",
    "Interested in {role} — {company}",
    "{role} Opportunity — Gajendra Dhanoliya, IIT Delhi",
]

# Role categories
ROLE_CATEGORIES = [
    {
        "name": "ai_ml",
        "keywords": [
            "ai", "ml", "machine learning", "deep learning",
            "data science", "computer vision", "nlp",
            "natural language", "artificial intelligence",
        ],
        "skills": (
            "Embeddings, Vector Databases, Semantic Search,"
            " OCR (Tesseract, PyMuPDF),"
            " Computer Vision (OpenCV),"
            " Face Matching & Liveness Detection,"
            " and Combinatorial Optimization (OR-Tools)"
        ),
        "experience_highlight": (
            "During my internship at ZeTheta Algorithms,"
            " I developed a Video KYC pipeline using"
            " OpenCV for face matching and liveness"
            " detection, achieving 93% verification"
            " accuracy."
        ),
        "project_highlight": (
            "I also built a Document Intelligence"
            " Assistant with vector embeddings in"
            " pgvector and pluggable OpenAI/HuggingFace"
            " semantic search capabilities."
        ),
    },
    {
        "name": "data",
        "keywords": [
            "data engineer", "data analyst", "data pipeline",
            "etl", "analytics", "data processing", "big data",
        ],
        "skills": (
            "PostgreSQL, Redis, Celery, Docker, pgvector,"
            " async data pipelines, SQLAlchemy 2,"
            " Alembic migrations,"
            " and document ingestion (12+ formats)"
        ),
        "experience_highlight": (
            "During my internship at ZeTheta Algorithms,"
            " I built backend systems with"
            " PostgreSQL-backed persistence and modular"
            " API design for compliance-ready workflows."
        ),
        "project_highlight": (
            "I also built a Document Intelligence"
            " Assistant with an asynchronous ingestion"
            " pipeline handling 12+ formats"
            " (PDF, DOCX, PPTX, XLSX, HTML, EML,"
            " images, ZIP) with automatic OCR fallback"
            " and token-aware chunking."
        ),
    },
    {
        "name": "frontend",
        "keywords": [
            "frontend", "front-end", "front end",
            "react", "ui developer", "ux", "web developer",
        ],
        "skills": (
            "React, TypeScript, Tailwind CSS,"
            " TanStack Query, Leaflet, Recharts,"
            " and WebSocket streaming"
        ),
        "experience_highlight": (
            "During my internship at ZeTheta Algorithms,"
            " I worked on building full-stack platforms"
            " with modern frontend and backend"
            " technologies."
        ),
        "project_highlight": (
            "I built RouteOS, a full-stack route"
            " optimization platform with React,"
            " real-time vehicle simulation streamed"
            " over WebSockets, and a Redis-cached"
            " analytics dashboard with Recharts"
            " visualizations."
        ),
    },
    {
        "name": "devops",
        "keywords": [
            "devops", "cloud", "infrastructure",
            "sre", "platform engineer",
            "site reliability", "systems engineer",
        ],
        "skills": (
            "Docker, Docker Compose, Prometheus,"
            " Structured Logging, Linux, Git/GitHub,"
            " Redis, and CI/CD pipelines"
        ),
        "experience_highlight": (
            "During my internship at ZeTheta Algorithms,"
            " I architected containerized backend"
            " services with Docker and implemented"
            " structured logging and monitoring."
        ),
        "project_highlight": (
            "I designed Docker Compose orchestration"
            " for multi-service deployments including"
            " FastAPI backends, Celery/Redis workers,"
            " PostgreSQL databases, and Prometheus"
            " monitoring."
        ),
    },
    {
        "name": "backend",
        "keywords": [
            "backend", "back-end", "back end",
            "software engineer", "software developer",
            "sde", "fullstack", "full stack",
            "full-stack", "developer", "engineer",
            "intern", "programmer",
        ],
        "skills": (
            "FastAPI, REST API Design, WebSockets,"
            " PostgreSQL, Redis, Docker,"
            " JWT Authentication, SQLAlchemy,"
            " Celery, and Async I/O"
        ),
        "experience_highlight": (
            "During my internship at ZeTheta Algorithms,"
            " I engineered an automated eKYC platform"
            " with modular REST APIs, JWT-based"
            " authentication with RBAC, and"
            " PostgreSQL-backed persistence, cutting"
            " manual onboarding time by 60%."
        ),
        "project_highlight": (
            "I also built an Online Coding Judge with"
            " Docker-sandboxed code execution, automated"
            " evaluation pipelines, and a multi-resource"
            " REST API across 7 domains."
        ),
    },
]

DEFAULT_ROLE_CONFIG = {
    "name": "default",
    "skills": (
        "C++ (C++20), Python, TypeScript, SQL,"
        " FastAPI, REST APIs, PostgreSQL,"
        " Docker, and Git/GitHub"
    ),
    "experience_highlight": (
        "I recently worked as a Software Engineer"
        " Intern at ZeTheta Algorithms, where I"
        " engineered an automated eKYC platform and"
        " developed backend APIs, cutting manual"
        " onboarding time by 60%."
    ),
    "project_highlight": (
        "I have also built projects like RouteOS"
        " (a logistics optimization platform) and an"
        " Online Coding Judge with Docker-sandboxed"
        " code execution."
    ),
}

RESUME_PATH = "Resume_Gajendra_Dhanoliya.pdf"

# Anti-spam delay settings
MIN_DELAY_SECONDS = 12
MAX_DELAY_SECONDS = 25
BATCH_BREAK_EVERY = 30
BATCH_BREAK_SECONDS = 90

# Default behavior when no CLI args are given
DRY_RUN = True
MAX_EMAILS_PER_RUN = 5

ENFORCE_BUSINESS_HOURS = False
BUSINESS_START_HOUR = 9
BUSINESS_END_HOUR = 19

SEND_LOG = "send_log.csv"
LOG_FIELDS = [
    "row_number",
    "hr_name",
    "hr_email",
    "company",
    "role",
    "experience",
    "location",
    "category",
    "status",
    "error",
    "sent_date",
    "email_type",
]

ENABLE_FOLLOW_UP = False
FOLLOW_UP_AFTER_DAYS = 7
FOLLOW_UP_SUBJECT = "Following Up — {role} at {company}"

DAILY_SEND_LIMIT = 450
REPORT_FILE = "report.html"


# =====================================================================
# HELPERS & TEXT CLEANERS
# =====================================================================

def normalize_header(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def clean_text(value, default=""):
    if value is None:
        return default
    value = str(value).strip()
    return value if value else default


def normalize_email(email):
    return email.strip().lower()


DISPOSABLE_DOMAINS = {
    "example.com", "test.com", "tempmail.com", "mailinator.com",
    "10minutemail.com", "throwaway.com", "guerrillamail.com"
}

def is_valid_email(email):
    if not email or len(email) > 254:
        return False
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    if not re.match(pattern, email):
        return False
    domain = email.split("@")[1].lower()
    if domain in DISPOSABLE_DOMAINS or "." not in domain:
        return False
    if email.startswith(("no-reply", "noreply", "donotreply")):
        return False
    return True


def clean_hr_name(name):
    name = clean_text(name, "there")
    invalid_names = {
        "sir/ma'am", "sir/ma\u2019am", "sir", "ma'am",
        "ma\u2019am", "hr", "recruiter", "there",
        "n/a", "na", "-", "admin", "team", "hiring manager",
        "none", "null", "unknown"
    }
    if name.lower() in invalid_names:
        return "there"

    cleaned = re.sub(r'^(mr\.|ms\.|mrs\.|dr\.)\s+', '', name, flags=re.IGNORECASE).strip()
    parts = cleaned.split()
    if parts:
        first = parts[0].capitalize()
        if len(first) > 1 and first.isalpha():
            return first
    return "there"


def clean_company_name(name):
    if not name or name == "the company":
        return name
    name = re.sub(r'\s+', ' ', name.strip())
    words = name.split()
    n = len(words)
    for split_at in range(1, n):
        first = " ".join(words[:split_at]).lower()
        rest = " ".join(words[split_at:]).lower()
        if first == rest:
            return " ".join(words[:split_at])
    return name


def clean_role_title(role):
    if not role:
        return "Software Engineer"
    role = clean_text(role, "Software Engineer")
    segments = re.split(r'[,;/|]', role)
    cleaned_segment = segments[0].strip()
    cleaned_segment = re.sub(r'\(.*?\)', '', cleaned_segment).strip()
    cleaned_segment = re.sub(
        r'\b(immediate joiner|urgently hiring|internship|fresher welcome)\b',
        '', cleaned_segment, flags=re.IGNORECASE
    ).strip()
    cleaned_segment = re.sub(r'\s+', ' ', cleaned_segment)
    return cleaned_segment if len(cleaned_segment) > 2 else role


def format_location_line(location):
    if not location:
        return ""
    loc = location.strip()
    if re.search(r'\bremote\b', loc, re.I):
        return "I am comfortable working in a remote setup or on-site as required."
    primary_city = re.split(r'[,/]', loc)[0].strip()
    if len(primary_city) > 2:
        return f"I am actively open to working on-site in {primary_city} or relocating as needed."
    return ""


# =====================================================================
# ROLE-BASED EMAIL BODY
# =====================================================================

def detect_role_category(role):
    role_lower = role.lower()
    for category in ROLE_CATEGORIES:
        for keyword in category["keywords"]:
            if len(keyword) <= 3:
                pattern = r'\b' + re.escape(keyword) + r'\b'
                if re.search(pattern, role_lower):
                    return category
            else:
                if keyword in role_lower:
                    return category
    return DEFAULT_ROLE_CONFIG


def build_email_body(contact):
    config = detect_role_category(contact["role"])
    loc_sentence = format_location_line(contact.get("location", ""))
    loc_paragraph = f"\n{loc_sentence}\n" if loc_sentence else ""

    body = (
        f"Hi {contact['hr_name']},\n"
        f"\n"
        f"I hope you're doing well.\n"
        f"\n"
        f"I'm Gajendra Dhanoliya, an IIT Delhi graduate"
        f" with hands-on experience in software"
        f" engineering. I came across the"
        f" {contact['role']} opportunity at"
        f" {contact['company']} and wanted to reach out"
        f" regarding the role.\n"
        f"\n"
        f"{config['experience_highlight']}\n"
        f"\n"
        f"My technical strengths include"
        f" {config['skills']}."
        f" {config['project_highlight']}\n"
        f"{loc_paragraph}"
        f"\n"
        f"I've attached my resume for your reference."
        f" I would be grateful if you could consider"
        f" my profile for this opportunity.\n"
        f"\n"
        f"Best regards,\n"
        f"{SENDER_NAME}\n"
    )
    return body


def build_follow_up_body(contact):
    config = detect_role_category(contact["role"])
    first_skill = config["skills"].split(",")[0].strip()

    body = (
        f"Hi {contact['hr_name']},\n"
        f"\n"
        f"I hope you're doing well. I had reached out"
        f" regarding the {contact['role']} opportunity"
        f" at {contact['company']} about a week ago.\n"
        f"\n"
        f"I wanted to follow up and reiterate my"
        f" interest in this role. I believe my"
        f" experience with {first_skill} and related"
        f" technologies would be a great fit.\n"
        f"\n"
        f"I'd be happy to discuss further at your"
        f" convenience. My resume is attached again"
        f" for your reference.\n"
        f"\n"
        f"Best regards,\n"
        f"{SENDER_NAME}\n"
    )
    return body


def build_html_body(plain_body):
    paragraphs = plain_body.strip().split("\n\n")
    html_paras = ""
    for para in paragraphs:
        para = para.strip().replace("\n", "<br>")
        if para:
            html_paras += (
                '<p style="margin:0 0 12px 0;'
                'line-height:1.6;color:#333;">'
                f'{para}</p>\n'
            )

    html = f"""<!DOCTYPE html>
<html>
<body style="font-family:Arial,Helvetica,sans-serif;
             font-size:14px;color:#333;max-width:600px;">
{html_paras}
<div style="margin-top:24px;padding-top:16px;
            border-top:1px solid #ddd;">
  <p style="margin:0;font-weight:bold;font-size:14px;
            color:#222;">
    {SENDER_NAME}
  </p>
  <p style="margin:2px 0;font-size:13px;color:#666;">
    IIT Delhi &bull; B.Tech Electrical Engineering
  </p>
  <p style="margin:2px 0;font-size:13px;color:#666;">
    &#128222; {SENDER_PHONE}
  </p>
  <p style="margin:2px 0;font-size:13px;">
    <a href="{LINKEDIN_URL}"
       style="color:#0077b5;text-decoration:none;">
      LinkedIn
    </a>
    &nbsp;&bull;&nbsp;
    <a href="{GITHUB_URL}"
       style="color:#333;text-decoration:none;">
      GitHub
    </a>
  </p>
</div>
</body>
</html>"""
    return html


# =====================================================================
# BUSINESS HOURS CHECK
# =====================================================================

def is_business_hours():
    now = datetime.now()
    if now.weekday() >= 5:
        return False
    if BUSINESS_START_HOUR <= now.hour < BUSINESS_END_HOUR:
        return True
    return False


# =====================================================================
# SEND LOG
# =====================================================================

def load_send_log(path):
    sent_emails = set()
    follow_up_done = set()
    entries = []

    if not os.path.isfile(path):
        return sent_emails, follow_up_done, entries

    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                entries.append(row)
                email = row.get("hr_email", "").strip().lower()
                status = row.get("status", "")
                email_type = row.get("email_type", "first")

                if status == "sent":
                    sent_emails.add(email)
                if email_type == "follow_up" and status == "sent":
                    follow_up_done.add(email)
    except Exception:
        pass

    return sent_emails, follow_up_done, entries


def count_today_sends(entries):
    today_str = date.today().isoformat()
    count = 0
    for entry in entries:
        sent_date = entry.get("sent_date", "")
        status = entry.get("status", "")
        if status == "sent" and sent_date.startswith(today_str):
            count += 1
    return count


def get_follow_up_candidates(entries, follow_up_done):
    candidates = []
    cutoff = date.today().toordinal() - FOLLOW_UP_AFTER_DAYS
    for entry in entries:
        email = entry.get("hr_email", "").strip().lower()
        status = entry.get("status", "")
        email_type = entry.get("email_type", "first")
        sent_date_str = entry.get("sent_date", "")

        if status != "sent" or email_type != "first":
            continue
        if email in follow_up_done:
            continue

        try:
            sent_date = date.fromisoformat(sent_date_str[:10])
            if sent_date.toordinal() <= cutoff:
                candidates.append(entry)
        except (ValueError, TypeError):
            continue
    return candidates


def migrate_send_log(path):
    if not os.path.isfile(path) or os.path.getsize(path) == 0:
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and set(reader.fieldnames) == set(LOG_FIELDS):
                return
            existing = list(reader)
    except Exception:
        return

    for entry in existing:
        for field in LOG_FIELDS:
            if field not in entry:
                entry[field] = ""

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=LOG_FIELDS, extrasaction="ignore"
        )
        writer.writeheader()
        writer.writerows(existing)

    print("  Log migrated to new format.")


def save_single_result(path, result):
    file_exists = os.path.isfile(path) and os.path.getsize(path) > 0
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=LOG_FIELDS, extrasaction="ignore"
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow(result)


# =====================================================================
# LOAD CONTACTS
# =====================================================================

def load_contacts(path, sheet_name=None):
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Excel file not found: {os.path.abspath(path)}"
        )

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        raise ValueError("Excel sheet is empty.")

    header_row = rows[0]
    headers = [normalize_header(h) for h in header_row]

    def find_col(*candidates):
        for candidate in candidates:
            candidate = candidate.lower()
            if candidate in headers:
                return headers.index(candidate)
        return None

    name_idx = find_col("hr name", "name")
    email_idx = find_col("hr email", "email", "e-mail")
    company_idx = find_col("company", "company name")
    role_idx = find_col("role", "job role", "position", "job title")
    experience_idx = find_col("experience")
    location_idx = find_col("location")

    missing = []
    if email_idx is None:
        missing.append("email")
    if company_idx is None:
        missing.append("company")

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
            + f"\nFound headers: {header_row}"
        )

    contacts = []
    seen_emails = set()
    skipped_no_email = 0
    skipped_invalid_email = 0
    skipped_duplicate = 0

    for row_number, row in enumerate(rows[1:], start=2):
        if not row or all(value is None for value in row):
            continue

        email = clean_text(row[email_idx] if email_idx < len(row) else None)
        if not email:
            skipped_no_email += 1
            continue

        email = normalize_email(email)
        if not is_valid_email(email):
            skipped_invalid_email += 1
            continue

        if email in seen_emails:
            skipped_duplicate += 1
            continue
        seen_emails.add(email)

        name = clean_hr_name(
            row[name_idx] if name_idx is not None and name_idx < len(row) else None
        )
        company = clean_company_name(
            clean_text(row[company_idx] if company_idx < len(row) else None, "the company")
        )
        raw_role = clean_text(
            row[role_idx] if role_idx is not None and role_idx < len(row) else None,
            "Software Engineering"
        )
        role = clean_role_title(raw_role)
        experience = clean_text(
            row[experience_idx] if experience_idx is not None and experience_idx < len(row) else ""
        )
        location = clean_text(
            row[location_idx] if location_idx is not None and location_idx < len(row) else ""
        )

        contacts.append({
            "hr_name": name,
            "hr_email": email,
            "company": company,
            "role": role,
            "experience": experience,
            "location": location,
            "row_number": row_number
        })

    print("\nContact processing summary:")
    print(f"  Valid contacts      : {len(contacts)}")
    print(f"  Missing email       : {skipped_no_email}")
    print(f"  Invalid email       : {skipped_invalid_email}")
    print(f"  Duplicate emails    : {skipped_duplicate}")

    return contacts


# =====================================================================
# SMTP & SEND EMAIL
# =====================================================================

def connect_smtp():
    context = ssl.create_default_context()
    try:
        smtp = smtplib.SMTP("smtp.gmail.com", 587, timeout=25)
        smtp.starttls(context=context)
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        return smtp
    except smtplib.SMTPAuthenticationError:
        print("\nERROR: Gmail authentication failed.")
        print("Check that you are using a valid Google App Password.")
        print("Do NOT use your normal Gmail password.")
        return None
    except Exception as error:
        print(f"\nERROR connecting to Gmail:\n{error}")
        return None


def send_email(
    smtp,
    to_email,
    subject,
    body,
    html_body=None,
    attachment_path=None
):
    msg = MIMEMultipart("mixed")
    msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"] = to_email
    msg["Subject"] = subject

    if html_body:
        alt_part = MIMEMultipart("alternative")
        alt_part.attach(MIMEText(body, "plain", "utf-8"))
        alt_part.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alt_part)
    else:
        msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachment_path:
        with open(attachment_path, "rb") as file:
            part = MIMEApplication(
                file.read(),
                Name=os.path.basename(attachment_path)
            )
        part["Content-Disposition"] = (
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        msg.attach(part)

    smtp.sendmail(SENDER_EMAIL, to_email, msg.as_string())


# =====================================================================
# INTERACTIVE REPORT (SEARCH & FILTER UI)
# =====================================================================

def generate_report(results, skipped_count, today_total, run_start, mode_label):
    """
    Generates an interactive HTML dashboard report with live search,
    status filtering, category metrics, and export capabilities.
    """
    sent = sum(1 for r in results if r["status"] == "sent")
    failed = sum(1 for r in results if r["status"] == "failed")
    dry_run_count = sum(1 for r in results if r["status"] == "dry_run")
    follow_ups = sum(
        1 for r in results
        if r.get("email_type") == "follow_up" and r["status"] in ("sent", "dry_run")
    )

    categories = {}
    for r in results:
        cat = r.get("category", "unknown")
        categories[cat] = categories.get(cat, 0) + 1

    category_rows = ""
    for cat, count in sorted(categories.items()):
        label = cat.replace("_", " ").title()
        category_rows += f"<tr><td><strong>{label}</strong></td><td><span class='badge'>{count}</span></td></tr>\n"

    # Build interactive table rows for all results in this run
    result_rows = ""
    for r in results:
        status_class = "status-" + r.get("status", "dry_run")
        status_badge = f"<span class='tag {status_class}'>{r.get('status', '').upper()}</span>"
        err = r.get("error", "")
        err_html = f"<small style='color:#ef4444;display:block;'>{err}</small>" if err else ""
        
        result_rows += f"""
        <tr data-status="{r.get('status', '')}" data-category="{r.get('category', '')}">
          <td>{r.get('hr_name', 'there')}</td>
          <td><strong>{r.get('hr_email', '')}</strong></td>
          <td>{r.get('company', '')}</td>
          <td>{r.get('role', '')}</td>
          <td><span class="category-tag">{r.get('category', 'default')}</span></td>
          <td>{status_badge}{err_html}</td>
        </tr>
        """

    duration = datetime.now() - run_start
    minutes = duration.seconds // 60
    seconds = duration.seconds % 60

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>HR Email Automation Dashboard</title>
  <style>
    :root {{
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --primary: #2563eb;
      --text: #0f172a;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --success: #16a34a;
      --danger: #dc2626;
      --warning: #d97706;
    }}
    * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }}
    body {{ background: var(--bg); color: var(--text); padding: 24px; }}
    .container {{ max-width: 1000px; margin: 0 auto; }}
    .header {{ margin-bottom: 24px; }}
    .header h1 {{ font-size: 24px; font-weight: 700; color: var(--text); }}
    .header p {{ font-size: 14px; color: var(--text-muted); margin-top: 4px; }}
    
    .stats-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}
    .stat-card {{
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px;
      text-align: center;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }}
    .stat-card .val {{ font-size: 28px; font-weight: 700; }}
    .stat-card .lbl {{ font-size: 13px; color: var(--text-muted); margin-top: 2px; }}
    .sent .val {{ color: var(--success); }}
    .failed .val {{ color: var(--danger); }}
    .dry .val {{ color: var(--warning); }}
    .followup .val {{ color: var(--primary); }}

    .card {{
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }}
    .card h2 {{ font-size: 17px; margin-bottom: 16px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; }}

    .controls {{
      display: flex;
      gap: 12px;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }}
    .search-box {{
      flex: 1;
      min-width: 220px;
      padding: 9px 14px;
      border: 1px solid var(--border);
      border-radius: 8px;
      font-size: 14px;
      outline: none;
    }}
    .search-box:focus {{ border-color: var(--primary); }}
    .filter-btn {{
      padding: 8px 14px;
      border: 1px solid var(--border);
      background: #fff;
      border-radius: 8px;
      font-size: 13px;
      cursor: pointer;
      font-weight: 500;
      color: var(--text-muted);
      transition: all 0.15s;
    }}
    .filter-btn.active, .filter-btn:hover {{
      background: var(--text);
      color: #fff;
      border-color: var(--text);
    }}

    table {{ width: 100%; border-collapse: collapse; font-size: 13.5px; }}
    th {{ background: #f1f5f9; text-align: left; padding: 10px 12px; color: var(--text-muted); font-weight: 600; font-size: 12px; text-transform: uppercase; }}
    td {{ padding: 12px; border-top: 1px solid var(--border); vertical-align: middle; }}
    tr:hover td {{ background: #f8fafc; }}

    .tag {{ display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; }}
    .status-sent {{ background: #dcfce7; color: #15803d; }}
    .status-failed {{ background: #fee2e2; color: #b91c1c; }}
    .status-dry_run {{ background: #fef3c7; color: #b45309; }}
    .category-tag {{ background: #f1f5f9; color: #475569; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 500; }}
    .badge {{ background: #e2e8f0; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: 600; }}

    .meta-box {{ font-size: 13px; color: var(--text-muted); line-height: 1.8; }}
    .meta-box strong {{ color: var(--text); }}
  </style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>HR Outreach Live Dashboard</h1>
    <p>{run_start.strftime('%A, %d %b %Y &bull; %I:%M %p')} &bull; Runtime: {minutes}m {seconds}s &bull; Today: {today_total} / {DAILY_SEND_LIMIT}</p>
  </div>

  <div class="stats-grid">
    <div class="stat-card sent"><div class="val">{sent}</div><div class="lbl">Sent Live</div></div>
    <div class="stat-card failed"><div class="val">{failed}</div><div class="lbl">Failed</div></div>
    <div class="stat-card dry"><div class="val">{dry_run_count}</div><div class="lbl">Dry Run</div></div>
    <div class="stat-card followup"><div class="val">{follow_ups}</div><div class="lbl">Follow-ups</div></div>
    <div class="stat-card"><div class="val">{skipped_count}</div><div class="lbl">Skipped (Sent)</div></div>
  </div>

  <div class="card">
    <h2>
      <span>Processed Contacts ({len(results)})</span>
    </h2>
    <div class="controls">
      <input type="text" id="searchInput" class="search-box" placeholder="🔍 Search by company, email, role, or name..." onkeyup="filterTable()">
      <button class="filter-btn active" onclick="setFilter('all', this)">All</button>
      <button class="filter-btn" onclick="setFilter('sent', this)">Sent</button>
      <button class="filter-btn" onclick="setFilter('failed', this)">Failed</button>
      <button class="filter-btn" onclick="setFilter('dry_run', this)">Dry Run</button>
    </div>

    <div style="overflow-x: auto;">
      <table id="contactsTable">
        <thead>
          <tr>
            <th>HR Name</th>
            <th>Email</th>
            <th>Company</th>
            <th>Role</th>
            <th>Category</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {result_rows}
        </tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <h2>Category Distribution</h2>
    <table>
      <thead>
        <tr><th>Role Category</th><th>Count</th></tr>
      </thead>
      <tbody>
        {category_rows}
      </tbody>
    </table>
  </div>

  <div class="card meta-box">
    <strong>Run Settings</strong> &bull;
    Mode: <span>{mode_label}</span> &bull;
    Daily Cap: <span>{DAILY_SEND_LIMIT}</span> &bull;
    Jitter Delay: <span>{MIN_DELAY_SECONDS}s – {MAX_DELAY_SECONDS}s</span> &bull;
    Attachment: <span>{os.path.basename(RESUME_PATH)}</span>
  </div>
</div>

<script>
  let currentFilter = 'all';

  function setFilter(status, btn) {{
    currentFilter = status;
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filterTable();
  }}

  function filterTable() {{
    const query = document.getElementById('searchInput').value.toLowerCase();
    const rows = document.querySelectorAll('#contactsTable tbody tr');

    rows.forEach(row => {{
      const text = row.innerText.toLowerCase();
      const status = row.getAttribute('data-status');
      const matchesSearch = text.includes(query);
      const matchesFilter = (currentFilter === 'all') || (status === currentFilter);

      if (matchesSearch && matchesFilter) {{
        row.style.display = '';
      }} else {{
        row.style.display = 'none';
      }}
    }});
  }}
</script>
</body>
</html>"""

    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nInteractive Report saved to: {REPORT_FILE}")


# =====================================================================
# CLI ARGUMENT PARSER
# =====================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="HR Email Sender — Cold outreach automation tool."
    )
    parser.add_argument(
        "--send", type=int, nargs="?", const=10, default=None,
        help="Run in LIVE mode and send N emails (default: 10). Example: --send 20"
    )
    parser.add_argument(
        "--dry-run", type=int, nargs="?", const=5, default=None,
        help="Run in DRY-RUN preview mode for N emails (default: 5). Example: --dry-run 10"
    )
    parser.add_argument(
        "--test", action="store_true",
        help="Send 1 sample email to yourself (SENDER_EMAIL) to verify layout in your inbox."
    )
    parser.add_argument(
        "--stats", action="store_true",
        help="Display today's sending stats from log without processing contacts."
    )
    return parser.parse_args()


# =====================================================================
# MAIN
# =====================================================================

def main():
    args = parse_args()
    run_start = datetime.now()

    print("=" * 70)
    print("HR EMAIL SENDER — PRO AUTOMATION")
    print("=" * 70)

    # ---------------------------------------------------------------
    # Stats-only mode
    # ---------------------------------------------------------------
    if args.stats:
        sent_emails, follow_up_done, log_entries = load_send_log(SEND_LOG)
        today_count = count_today_sends(log_entries)
        print(f"\n[STATS] Total unique emails contacted so far: {len(sent_emails)}")
        print(f"[STATS] Emails sent today: {today_count} / {DAILY_SEND_LIMIT}")
        print(f"[STATS] Follow-ups sent so far: {len(follow_up_done)}")
        return

    # Determine mode & batch count from CLI flags or default config
    is_dry_run = DRY_RUN
    batch_limit = MAX_EMAILS_PER_RUN
    is_test_self = args.test

    if args.send is not None:
        is_dry_run = False
        batch_limit = args.send
    elif args.dry_run is not None:
        is_dry_run = True
        batch_limit = args.dry_run
    elif is_test_self:
        is_dry_run = False
        batch_limit = 1

    mode_label = "TEST SELF" if is_test_self else ("DRY RUN" if is_dry_run else "LIVE SEND")

    # Business Hours Warning
    if not is_business_hours() and not is_dry_run and not is_test_self:
        now_str = datetime.now().strftime("%A, %I:%M %p")
        if ENFORCE_BUSINESS_HOURS:
            print(f"\n[!] NOTICE: Current time ({now_str}) is outside business hours (Mon-Fri 9AM-7PM).")
            print("Aborting because ENFORCE_BUSINESS_HOURS is True.")
            return
        else:
            print(f"\n[i] Info: Current time ({now_str}) is outside ideal business hours (Mon-Fri 9AM-7PM).")

    # Load contacts
    print(f"\nLoading contacts from: {EXCEL_FILE}")
    try:
        contacts = load_contacts(EXCEL_FILE, SHEET_NAME)
    except Exception as error:
        print(f"\nERROR loading Excel:\n{error}")
        sys.exit(1)

    if not contacts:
        print("\nNo valid contacts found.")
        return

    # Load send log
    print(f"\nLoading send log: {SEND_LOG}")
    migrate_send_log(SEND_LOG)
    sent_emails, follow_up_done, log_entries = load_send_log(SEND_LOG)

    if sent_emails:
        print(f"  Previously sent: {len(sent_emails)}")
    if follow_up_done:
        print(f"  Follow-ups done: {len(follow_up_done)}")

    today_count = count_today_sends(log_entries)
    print(f"\nEmails sent today: {today_count}")
    remaining_today = DAILY_SEND_LIMIT - today_count

    if remaining_today <= 0 and not is_dry_run and not is_test_self:
        print(f"\nDaily limit ({DAILY_SEND_LIMIT}) reached. Try again tomorrow.")
        return

    # Filter contacts
    total_before = len(contacts)
    if not is_test_self:
        contacts = [c for c in contacts if c["hr_email"] not in sent_emails]
    skipped_already_sent = total_before - len(contacts)

    effective_limit = len(contacts)
    if batch_limit is not None:
        effective_limit = min(effective_limit, batch_limit)
    if not is_dry_run and not is_test_self:
        effective_limit = min(effective_limit, remaining_today)

    contacts_to_process = contacts[:effective_limit]

    print(f"\nTotal contacts in Excel : {total_before}")
    print(f"Already sent (Skipping) : {skipped_already_sent}")
    print(f"New contacts remaining  : {len(contacts)}")
    print(f"Processing this run     : {len(contacts_to_process)}")
    print(f"Active Mode             : {mode_label}")

    # Resume check
    attachment_path = RESUME_PATH if RESUME_PATH else None
    if attachment_path:
        if not os.path.isfile(attachment_path):
            print("\nERROR: Resume not found:")
            print(os.path.abspath(attachment_path))
            sys.exit(1)
        size_kb = os.path.getsize(attachment_path) / 1024
        print(f"\nResume: {attachment_path} ({size_kb:.0f} KB)")

    # Connect Gmail SMTP (if sending live)
    smtp = None
    if not is_dry_run:
        if not APP_PASSWORD:
            print("\nERROR: HR_MAIL_APP_PASSWORD is not set.")
            print('\nSet it in CMD using: setx HR_MAIL_APP_PASSWORD "YOUR_APP_PASSWORD"')
            print("Then close and reopen CMD.")
            sys.exit(1)

        print("\nConnecting to Gmail SMTP...")
        smtp = connect_smtp()
        if not smtp:
            sys.exit(1)
        print("Gmail authentication successful.")

    # ---------------------------------------------------------------
    # Send loop
    # ---------------------------------------------------------------
    results = []

    for index, contact in enumerate(contacts_to_process, start=1):
        target_email = SENDER_EMAIL if is_test_self else contact["hr_email"]
        category = detect_role_category(contact["role"])
        cat_name = category.get("name", "default")
        subject = random.choice(SUBJECT_TEMPLATES).format(**contact)
        if is_test_self:
            subject = f"[TEST SAMPLE] {subject}"

        body = build_email_body(contact)
        html_body = build_html_body(body)

        print(f"\n[{index}/{len(contacts_to_process)}]")
        print(f"  Name     : {contact['hr_name']}")
        print(f"  Target   : {target_email}")
        print(f"  Company  : {contact['company']}")
        print(f"  Role     : {contact['role']}")
        print(f"  Category : {cat_name}")

        if is_dry_run:
            print("\n  --- SUBJECT ---")
            print(f"  {subject}")
            print("\n  --- BODY ---")
            print(body)
            if attachment_path:
                print(f"  --- ATTACHMENT --- {attachment_path}")
            print("\n  [DRY RUN - NOT SENT]")

            results.append({
                **contact,
                "category": cat_name,
                "status": "dry_run",
                "error": "",
                "sent_date": datetime.now().isoformat(timespec="seconds"),
                "email_type": "first",
            })
            continue

        # LIVE / TEST SEND
        try:
            send_email(
                smtp,
                target_email,
                subject,
                body,
                html_body,
                attachment_path
            )
            print("  SENT OK")
            result = {
                **contact,
                "category": cat_name,
                "status": "sent",
                "error": "",
                "sent_date": datetime.now().isoformat(timespec="seconds"),
                "email_type": "first",
            }
            results.append(result)
            if not is_test_self:
                save_single_result(SEND_LOG, result)

        except smtplib.SMTPServerDisconnected:
            print("  Connection lost. Reconnecting...")
            smtp = connect_smtp()
            if smtp:
                try:
                    send_email(
                        smtp,
                        target_email,
                        subject,
                        body,
                        html_body,
                        attachment_path
                    )
                    print("  SENT OK (after reconnect)")
                    result = {
                        **contact,
                        "category": cat_name,
                        "status": "sent",
                        "error": "",
                        "sent_date": datetime.now().isoformat(timespec="seconds"),
                        "email_type": "first",
                    }
                    results.append(result)
                    if not is_test_self:
                        save_single_result(SEND_LOG, result)
                except Exception as error:
                    print(f"  FAILED after reconnect: {error}")
                    result = {
                        **contact,
                        "category": cat_name,
                        "status": "failed",
                        "error": str(error),
                        "sent_date": datetime.now().isoformat(timespec="seconds"),
                        "email_type": "first",
                    }
                    results.append(result)
                    if not is_test_self:
                        save_single_result(SEND_LOG, result)
            else:
                print("  Reconnect failed. Stopping.")
                break

        except Exception as error:
            print(f"  FAILED: {error}")
            result = {
                **contact,
                "category": cat_name,
                "status": "failed",
                "error": str(error),
                "sent_date": datetime.now().isoformat(timespec="seconds"),
                "email_type": "first",
            }
            results.append(result)
            if not is_test_self:
                save_single_result(SEND_LOG, result)

        # Anti-spam delay
        if index < len(contacts_to_process) and not is_dry_run and not is_test_self:
            if index % BATCH_BREAK_EVERY == 0:
                print(f"\n  [Coffee Break] Pausing {BATCH_BREAK_SECONDS}s to keep delivery reputation clean...")
                time.sleep(BATCH_BREAK_SECONDS)
            else:
                delay = random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS)
                time.sleep(delay)

    # Close SMTP
    if smtp:
        try:
            smtp.quit()
        except Exception:
            pass

    # Generate interactive report
    new_sends = sum(1 for r in results if r["status"] == "sent")
    today_total = today_count + (new_sends if not is_test_self else 0)
    generate_report(results, skipped_already_sent, today_total, run_start, mode_label)

    sent = sum(1 for r in results if r["status"] == "sent")
    failed = sum(1 for r in results if r["status"] == "failed")
    dry_run_cnt = sum(1 for r in results if r["status"] == "dry_run")

    print("\n" + "=" * 70)
    print("DONE")
    print("=" * 70)
    print(f"Mode        : {mode_label}")
    if is_dry_run:
        print(f"Previewed   : {dry_run_cnt}")
    else:
        print(f"Sent        : {sent}")
        print(f"Failed      : {failed}")

    print(f"Skipped     : {skipped_already_sent} (already sent)")
    print(f"Today Total : {today_total} / {DAILY_SEND_LIMIT}")
    print(f"\nLog saved to : {SEND_LOG}")
    print(f"Report UI    : {REPORT_FILE}")


if __name__ == "__main__":
    main()
