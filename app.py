from flask import (
    Flask,
    render_template,
    render_template_string,
    request,
    redirect,
    url_for,
    session,
    flash,
    jsonify
)
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
from datetime import datetime
import time
import os
import openpyxl
from collections import Counter


# =========================================================
# APP CONFIGURATION
# =========================================================

app = Flask(__name__)

app.secret_key = "social-media-safety-project-secret-key"

DB_NAME = "database.db"

# REAL Google Form response export.
# The app automatically searches the project folder and common Windows
# download locations, so the Excel file does NOT have to be moved beside app.py.
SURVEY_FILENAME = "Online Safety & Cyber Awareness.xlsx"

def find_survey_file():
    """Find the real Google Form Excel export automatically."""
    project_folder = os.path.dirname(os.path.abspath(__file__))
    home = os.path.expanduser("~")
    search_roots = [
        project_folder,
        os.path.join(home, "Downloads"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Documents"),
    ]
    candidates = []
    seen = set()
    for root in search_roots:
        if not os.path.isdir(root):
            continue
        try:
            # Do NOT recursively scan the whole Windows user folders on every
            # /admin request. That can be extremely slow on large folders.
            for filename in os.listdir(root):
                lower = filename.lower()
                if lower.endswith(".xlsx") and "online safety" in lower and not filename.startswith("~$"):
                    path = os.path.abspath(os.path.join(root, filename))
                    if os.path.isfile(path) and path.lower() not in seen:
                        seen.add(path.lower())
                        candidates.append(path)
        except (PermissionError, OSError):
            continue
    if not candidates:
        return None
    exact = [p for p in candidates if os.path.basename(p).lower() == SURVEY_FILENAME.lower()]
    return max(exact, key=os.path.getmtime) if exact else max(candidates, key=os.path.getmtime)

SURVEY_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    SURVEY_FILENAME
)


# =========================================================
# DATABASE CONNECTION
# =========================================================

def get_db():
    conn = sqlite3.connect(
        DB_NAME,
        timeout=10
    )

    conn.row_factory = sqlite3.Row

    conn.execute(
        "PRAGMA busy_timeout = 10000"
    )

    return conn


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():

    conn = get_db()
    cursor = conn.cursor()

    # -----------------------------------------------------
    # USERS
    # -----------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT UNIQUE NOT NULL,
            password TEXT,
            created_at TEXT
        )
    """)

    # -----------------------------------------------------
    # MAIN QUIZ RESULTS
    # -----------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS quiz_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            score INTEGER,
            risk_level TEXT,
            habit_type TEXT,
            privacy_score INTEGER,
            security_score INTEGER,
            oversharing_score INTEGER,
            location_score INTEGER,
            reputation_score INTEGER,
            created_at TEXT
        )
    """)

    # -----------------------------------------------------
    # ADMIN
    # -----------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS admin (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password TEXT
        )
    """)

    # Create/reset the local project admin account so the Admin Panel
    # is immediately accessible after starting the application.
    # Change these credentials before deploying the project publicly.
    admin_username = "admin"
    admin_password = "admin123"
    cursor.execute(
        """
        INSERT INTO admin (username, password)
        VALUES (?, ?)
        ON CONFLICT(username) DO UPDATE SET password = excluded.password
        """,
        (admin_username, generate_password_hash(admin_password))
    )

    # -----------------------------------------------------
    # FINAL / WEAK-AREA TEST RESULTS
    # -----------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS final_test_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quiz_result_id INTEGER NOT NULL,
            weakest_area TEXT NOT NULL,
            score INTEGER NOT NULL,
            total_questions INTEGER NOT NULL,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()



# =========================================================
# GOOGLE FORM SURVEY ANALYTICS
# =========================================================

def _load_survey_data_uncached():
    """
    Read the REAL Google Form Excel export using openpyxl only.
    No pandas is required and no demo/fake survey records are generated.
    """
    empty = {
        "available": False,
        "total": 0,
        "file_name": os.path.basename(SURVEY_FILE),
        "awareness_score": 0,
        "awareness_level": "No survey data",
        "category_scores": {},
        "category_risks": {},
        "risk_areas": [],
        "risky_behaviors": [],
        "findings": [],
        "actions": [],
        "distributions": {},
        "responses": [],
        "learning_topics": {},
        "privacy_check": {},
        "confidence": {},
        "footprint": {},
        "career": {},
        "data_note": "Add the Google Form Excel export beside app.py."
    }

    # Automatically locate the user's real Google Form export.
    # It can remain in Downloads/Desktop/Documents or anywhere under the project folder.
    survey_path = find_survey_file()
    if not survey_path:
        empty["file_name"] = "Not found"
        empty["data_note"] = (
            "No Google Form Excel export was found automatically. "
            "Download the Google Form response .xlsx file anywhere under Downloads, "
            "Desktop, Documents, or the project folder and refresh the Admin Panel."
        )
        return empty

    try:
        from openpyxl import load_workbook

        workbook = load_workbook(
            survey_path,
            read_only=True,
            data_only=True
        )
        sheet = workbook.active

        rows = list(sheet.iter_rows(values_only=True))

        if not rows:
            workbook.close()
            empty["data_note"] = "The Google Form export contains no responses."
            return empty

        headers = [
            str(value).strip() if value is not None else ""
            for value in rows[0]
        ]

        records = []
        for raw_row in rows[1:]:
            record = {}
            for index, header in enumerate(headers):
                if not header:
                    continue
                value = raw_row[index] if index < len(raw_row) else None
                if value is None or str(value).strip() == "":
                    value = "Not answered"
                else:
                    value = str(value).strip()
                record[header] = value

            # Ignore completely blank spreadsheet rows.
            if any(value != "Not answered" for value in record.values()):
                records.append(record)

        workbook.close()

        if not records:
            empty["data_note"] = "The Google Form export contains no responses."
            return empty

        def values(column):
            if column not in headers:
                return []
            return [
                record.get(column, "Not answered")
                if str(record.get(column, "Not answered")).strip()
                else "Not answered"
                for record in records
            ]

        def counts(column):
            return dict(Counter(values(column)))

        def pct(count, total):
            if not total:
                return 0
            return round((count / total) * 100, 1)

        total = len(records)

        distributions = {
            "Age group": counts("What is your age group?"),
            "Most used platform": counts(
                "Which social media platform do you use most often?"
            ),
            "Social media frequency": counts(
                "How frequently do you use social media?"
            ),
            "Daily time": counts(
                "Approximately how much time do you spend on social media each day?"
            ),
            "Information shared": counts(
                "What type of information do you usually share on social media?"
            ),
            "Privacy settings": counts(
                "Do you usually check the privacy settings of your social media accounts?"
            ),
            "Confidence": counts(
                "How confident are you about protecting yourself from online threats?"
            ),
            "Learning topics": counts(
                "Which topic would you like to learn more about?"
            ),
            "Post-survey awareness": counts(
                "After completing this survey, do you feel more aware of your digital footprint?"
            )
        }

        answer_keys = {
            "Which of the following should NOT be shared publicly?":
                "Home address",

            "Do you accept friend or follow requests from people you do not know personally?":
                "Never",

            "Before posting something online, do you think about who might see it?":
                "Always",

            "Are you aware that information posted online can remain available even after you delete it?":
                "Yes, I am aware",

            "What do you understand by a digital footprint?":
                "The record of information and activities a person leaves online",

            "Do you think your social media activity can affect your future education or career opportunities?":
                "Yes",

            "Should you share your OTP (One-Time Password) with someone who asks for it?":
                "No",

            "What should you do if someone asks for your social media password?":
                "Refuse to share it",

            "Which of the following is the strongest password?":
                "A long and unique password containing different characters",

            "Should you use the same password for multiple accounts?":
                "No",

            "What should you do if you receive a suspicious link from an unknown person?":
                "Check it carefully or avoid clicking it",

            "What is the safest approach when using public Wi-Fi?":
                "Avoid sensitive transactions on an untrusted network",

            "What should you do if someone is cyberbullying or threatening you online?":
                "Save evidence and report or block the person",

            "What should you do if you receive an online message asking for your bank details?":
                "Ignore or report the message",

            "How can you identify a potentially fake social media account?":
                "It has suspicious activity or very little authentic information",

            "What should you do if you accidentally share sensitive personal information online?":
                "Delete or remove it and secure the affected account"
        }

        category_questions = {
            "Privacy": [
                "Which of the following should NOT be shared publicly?",
                "Do you accept friend or follow requests from people you do not know personally?",
                "Before posting something online, do you think about who might see it?",
                "What should you do if you accidentally share sensitive personal information online?"
            ],
            "Account Security": [
                "Should you share your OTP (One-Time Password) with someone who asks for it?",
                "What should you do if someone asks for your social media password?",
                "Which of the following is the strongest password?",
                "Should you use the same password for multiple accounts?",
                "What should you do if you receive a suspicious link from an unknown person?"
            ],
            "Oversharing": [
                "Which of the following should NOT be shared publicly?",
                "Before posting something online, do you think about who might see it?",
                "What should you do if you accidentally share sensitive personal information online?"
            ],
            "Location Safety": [
                "What is the safest approach when using public Wi-Fi?"
            ],
            "Digital Reputation": [
                "Are you aware that information posted online can remain available even after you delete it?",
                "What do you understand by a digital footprint?",
                "Do you think your social media activity can affect your future education or career opportunities?"
            ]
        }

        question_labels = {
            "Which of the following should NOT be shared publicly?":
                "Public sharing of personal information",
            "Do you accept friend or follow requests from people you do not know personally?":
                "Accepting requests from unknown people",
            "Before posting something online, do you think about who might see it?":
                "Thinking about the audience before posting",
            "Are you aware that information posted online can remain available even after you delete it?":
                "Understanding permanent online information",
            "What do you understand by a digital footprint?":
                "Understanding digital footprints",
            "Do you think your social media activity can affect your future education or career opportunities?":
                "Awareness of career/education impact",
            "Should you share your OTP (One-Time Password) with someone who asks for it?":
                "OTP protection",
            "What should you do if someone asks for your social media password?":
                "Password sharing",
            "Which of the following is the strongest password?":
                "Strong password selection",
            "Should you use the same password for multiple accounts?":
                "Password reuse",
            "What should you do if you receive a suspicious link from an unknown person?":
                "Suspicious-link handling",
            "What is the safest approach when using public Wi-Fi?":
                "Public Wi-Fi safety",
            "What should you do if someone is cyberbullying or threatening you online?":
                "Cyberbullying response",
            "What should you do if you receive an online message asking for your bank details?":
                "Bank-detail safety",
            "How can you identify a potentially fake social media account?":
                "Fake-profile awareness",
            "What should you do if you accidentally share sensitive personal information online?":
                "Accidental sensitive-data disclosure"
        }

        question_stats = []

        for question, correct in answer_keys.items():
            if question not in headers:
                continue

            response_values = values(question)
            correct_count = sum(
                1 for value in response_values
                if value == correct
            )
            answered = sum(
                1 for value in response_values
                if value != "Not answered"
            )
            wrong_count = max(0, total - correct_count)

            question_stats.append({
                "question": question,
                "label": question_labels.get(question, question),
                "correct": correct_count,
                "wrong": wrong_count,
                "answered": answered,
                "correct_pct": pct(correct_count, total),
                "risk_pct": pct(wrong_count, total),
                "risk": pct(wrong_count, total),
                "percentage": pct(wrong_count, total)
            })

        if question_stats:
            total_correct = sum(
                item["correct"] for item in question_stats
            )
            total_possible = len(question_stats) * total
            awareness_score = round(
                (total_correct / total_possible) * 100
            ) if total_possible else 0
        else:
            awareness_score = 0

        if awareness_score >= 80:
            awareness_level = "High Awareness"
        elif awareness_score >= 60:
            awareness_level = "Moderate Awareness"
        else:
            awareness_level = "Needs Improvement"

        category_scores = {}
        category_risks = {}

        for category, questions in category_questions.items():
            stats = [
                item for item in question_stats
                if item["question"] in questions
            ]

            if stats:
                score = round(
                    sum(item["correct_pct"] for item in stats) / len(stats)
                )
            else:
                score = None

            category_scores[category] = score
            category_risks[category] = (
                100 - score if score is not None else None
            )

        location_note = (
            "Based on the available public-Wi-Fi safety question; "
            "the current Google Form has no direct live-location question."
        )

        risk_areas = []

        for category, score in category_scores.items():
            if score is None:
                continue

            risk_areas.append({
                "label": category,
                "score": score,
                "risk_pct": 100 - score
            })

        risk_areas.sort(
            key=lambda item: item["risk_pct"],
            reverse=True
        )

        risky_behaviors = sorted(
            question_stats,
            key=lambda item: item["risk_pct"],
            reverse=True
        )[:7]

        findings = []

        privacy_counts = distributions["Privacy settings"]
        if privacy_counts:
            regular = privacy_counts.get("Regularly", 0)
            sometimes = privacy_counts.get("Sometimes", 0)
            findings.append({
                "title": "Privacy settings are not checked consistently",
                "text": (
                    f"{sometimes} of {total} respondents "
                    f"({pct(sometimes, total)}%) said they check privacy "
                    "settings only sometimes."
                )
            })

        confidence_counts = distributions["Confidence"]
        if confidence_counts:
            not_confident = sum(
                count
                for label, count in confidence_counts.items()
                if "not confident" in str(label).lower()
            )

            if not_confident:
                findings.append({
                    "title": "Confidence against threats needs attention",
                    "text": (
                        f"{not_confident} of {total} respondents "
                        f"({pct(not_confident, total)}%) reported low or "
                        "no confidence in protecting themselves online."
                    )
                })

        footprint_counts = distributions["Post-survey awareness"]
        if footprint_counts:
            aware = sum(
                count
                for label, count in footprint_counts.items()
                if str(label).lower().startswith("yes")
            )

            findings.append({
                "title": "Post-survey awareness",
                "text": (
                    f"{aware} of {total} respondents "
                    f"({pct(aware, total)}%) said they felt more aware "
                    "of their digital footprint after the survey."
                )
            })

        for item in risky_behaviors[:3]:
            findings.append({
                "title": item["label"],
                "text": (
                    f"{item['wrong']} of {total} responses "
                    f"({item['risk_pct']}%) did not select the "
                    "safer answer."
                )
            })

        findings = findings[:6]

        actions = []

        for item in risky_behaviors[:4]:
            if item["risk_pct"] < 20:
                continue

            actions.append({
                "title": f"Awareness session: {item['label']}",
                "text": (
                    "Run a short practical activity on this topic. "
                    f"Current risky-response rate is {item['risk_pct']}%."
                )
            })

        topics = distributions["Learning topics"]

        if topics:
            top_topic, top_topic_count = max(
                topics.items(),
                key=lambda pair: pair[1]
            )

            actions.append({
                "title": f"Prioritize requested topic: {top_topic}",
                "text": (
                    f"{top_topic_count} of {total} respondents "
                    "selected this as a topic they want to learn more about."
                )
            })

        actions.append({
            "title": "Repeat the survey after awareness activities",
            "text": (
                "Use the same Google Form after workshops or quizzes "
                "so the dashboard can compare future response patterns."
            )
        })

        response_columns = [
            "Timestamp",
            "Name*",
            "What is your age group?",
            "Which social media platform do you use most often?",
            "How frequently do you use social media?",
            "Approximately how much time do you spend on social media each day?",
            "Do you usually check the privacy settings of your social media accounts?",
            "How confident are you about protecting yourself from online threats?",
            "Which topic would you like to learn more about?",
            "After completing this survey, do you feel more aware of your digital footprint?"
        ]

        available_columns = [
            column for column in response_columns
            if column in headers
        ]

        rename_map = {
            "Timestamp": "timestamp",
            "Name*": "name",
            "What is your age group?": "age_group",
            "Which social media platform do you use most often?": "platform",
            "How frequently do you use social media?": "frequency",
            "Approximately how much time do you spend on social media each day?": "daily_time",
            "Do you usually check the privacy settings of your social media accounts?": "privacy",
            "How confident are you about protecting yourself from online threats?": "confidence",
            "Which topic would you like to learn more about?": "learning_topic",
            "After completing this survey, do you feel more aware of your digital footprint?": "footprint"
        }

        responses = []

        for record in records:
            response = {}

            for column in available_columns:
                key = rename_map.get(column, column)
                value = record.get(column, "")
                response[key] = (
                    "" if value == "Not answered" else str(value)
                )

            response["name"] = response.get("name") or "Anonymous"
            # Keep both names for compatibility with existing templates.
            response["privacy_settings"] = response.get("privacy", "")
            responses.append(response)

        chart_distributions = {
            "Platform": distributions["Most used platform"],
            "Privacy Settings": distributions["Privacy settings"],
            "Confidence": distributions["Confidence"],
            "Learning Topics": distributions["Learning topics"],
            "Post-Survey Awareness": distributions["Post-survey awareness"],
            "Age group": distributions["Age group"],
            "Social media frequency": distributions["Social media frequency"],
            "Daily time": distributions["Daily time"],
            "Information shared": distributions["Information shared"]
        }

        return {
            "available": True,
            "total": total,
            "file_name": os.path.basename(survey_path),
            "awareness_score": awareness_score,
            "awareness_level": awareness_level,
            "category_scores": category_scores,
            "category_risks": category_risks,
            "risk_areas": risk_areas,
            "risky_behaviors": risky_behaviors,
            "findings": findings,
            "actions": actions,
            "distributions": chart_distributions,
            "privacy_check": distributions["Privacy settings"],
            "confidence": distributions["Confidence"],
            "footprint": distributions["Post-survey awareness"],
            "career": counts(
                "Do you think your social media activity can affect your future education or career opportunities?"
            ),
            "learning_topics": topics,
            "responses": responses,
            "location_note": location_note,
            "data_note": (
                f"Live analysis of {total} responses from the "
                f"Google Form Excel export ({os.path.basename(survey_path)})."
            )
        }

    except Exception as error:
        print("SURVEY FILE ERROR:", error)

        empty["data_note"] = (
            "The Google Form file could not be analyzed: "
            + str(error)
        )

        return empty


# Cache the analysis so the Excel workbook is only read when it changes.
_SURVEY_CACHE = None
_SURVEY_CACHE_KEY = None

def load_survey_data():
    global _SURVEY_CACHE, _SURVEY_CACHE_KEY

    survey_path = find_survey_file()
    if not survey_path:
        if _SURVEY_CACHE is not None and _SURVEY_CACHE_KEY == (None,):
            return _SURVEY_CACHE
        data = _load_survey_data_uncached()
        _SURVEY_CACHE = data
        _SURVEY_CACHE_KEY = (None,)
        return data

    try:
        key = (survey_path, os.path.getmtime(survey_path), os.path.getsize(survey_path))
    except OSError:
        return _load_survey_data_uncached()

    if _SURVEY_CACHE is not None and _SURVEY_CACHE_KEY == key:
        return _SURVEY_CACHE

    data = _load_survey_data_uncached()
    _SURVEY_CACHE = data
    _SURVEY_CACHE_KEY = key
    return data


# =========================================================
# ADMIN PROTECTION
# =========================================================

def admin_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("is_admin"):

            flash(
                "Please log in as admin first.",
                "error"
            )

            return redirect(
                url_for("admin_login")
            )

        return function(*args, **kwargs)

    return wrapper


# =========================================================
# HABIT / RISK CALCULATION
# =========================================================

def habit_from_score(score):

    if score >= 85:

        return (
            "The Low-Trace User",
            "Low Risk"
        )

    elif score >= 65:

        return (
            "The Careful User",
            "Low Risk"
        )

    elif score >= 40:

        return (
            "The Oversharer",
            "Moderate Risk"
        )

    else:

        return (
            "The Easy Target",
            "High Risk"
        )


# =========================================================
# WEAKEST AREA
# =========================================================

def get_weakest_area(row):

    categories = {
        "Privacy": row["privacy_score"],
        "Account Security": row["security_score"],
        "Oversharing": row["oversharing_score"],
        "Location Safety": row["location_score"],
        "Digital Reputation": row["reputation_score"]
    }

    return min(
        categories,
        key=categories.get
    )


# =========================================================
# WEAK-AREA INFORMATION
# =========================================================

WEAK_AREA_INFO = {

    "Privacy": {
        "title": "Privacy Protection",
        "description": (
            "You need to improve how you protect personal "
            "information such as your college, phone number, "
            "birth date and other identifying details."
        ),
        "advice": [
            "Keep personal information visible only to trusted people.",
            "Review who can see your profile and posts.",
            "Avoid giving personal details to unknown people online.",
            "Think carefully before accepting new followers."
        ]
    },

    "Account Security": {
        "title": "Account Security",
        "description": (
            "Your answers suggest that your account-security "
            "habits need more attention."
        ),
        "advice": [
            "Use strong and unique passwords.",
            "Enable two-factor authentication.",
            "Never share passwords or private account screenshots.",
            "Verify suspicious requests before responding."
        ]
    },

    "Oversharing": {
        "title": "Oversharing",
        "description": (
            "You may be sharing more information online than "
            "is necessary or safe."
        ),
        "advice": [
            "Avoid posting your daily routine publicly.",
            "Do not announce when you are home alone.",
            "Think about what a stranger could learn from your posts.",
            "Use close-friends or private audiences when appropriate."
        ]
    },

    "Location Safety": {
        "title": "Location Safety",
        "description": (
            "Your location-sharing habits are the area that "
            "needs the most improvement."
        ),
        "advice": [
            "Avoid sharing your exact live location.",
            "Do not publicly reveal predictable routines.",
            "Consider posting location-based content after leaving.",
            "Check whether apps are allowed to access your location."
        ]
    },

    "Digital Reputation": {
        "title": "Digital Reputation",
        "description": (
            "Your answers indicate that you should think more "
            "carefully about the long-term effect of your posts."
        ),
        "advice": [
            "Think before posting photographs or comments.",
            "Avoid posting sensitive documents or IDs.",
            "Review old posts periodically.",
            "Remember that online content can remain accessible."
        ]
    }
}


# =========================================================
# FINAL TEST QUESTIONS
#
# These questions are selected according to the user's
# weakest category.
# =========================================================

FINAL_TEST_QUESTIONS = {

    "Privacy": [

        {
            "question": (
                "A stranger online asks which college you attend. "
                "What is the safest response?"
            ),
            "options": [
                "Tell them the college name",
                "Tell them the college and your class",
                "Avoid sharing personal information",
                "Send them your college ID"
            ],
            "answer": 2
        },

        {
            "question": (
                "Which information should generally NOT be made "
                "public on a social-media profile?"
            ),
            "options": [
                "A hobby",
                "A favourite movie",
                "Your phone number",
                "A general interest"
            ],
            "answer": 2
        },

        {
            "question": (
                "You receive a friend request from someone you do "
                "not recognize. What should you do?"
            ),
            "options": [
                "Accept immediately",
                "Share your personal details",
                "Verify who they are first",
                "Send them your phone number"
            ],
            "answer": 2
        },

        {
            "question": (
                "Why should you regularly review your social-media "
                "privacy settings?"
            ),
            "options": [
                "To increase followers",
                "To control who can access your information",
                "To get more notifications",
                "To make posts public"
            ],
            "answer": 1
        },

        {
            "question": (
                "Someone online asks for a screenshot of your "
                "private profile settings. What should you do?"
            ),
            "options": [
                "Send it immediately",
                "Send it if they seem friendly",
                "Avoid sharing private account information",
                "Post it publicly"
            ],
            "answer": 2
        }
    ],

    "Account Security": [

        {
            "question": (
                "Which password is the safest choice?"
            ),
            "options": [
                "password123",
                "yourname2005",
                "A strong unique password",
                "12345678"
            ],
            "answer": 2
        },

        {
            "question": (
                "What does two-factor authentication provide?"
            ),
            "options": [
                "More followers",
                "An additional layer of account protection",
                "A public profile",
                "Automatic password sharing"
            ],
            "answer": 1
        },

        {
            "question": (
                "A person claiming to be technical support asks "
                "for your password. What should you do?"
            ),
            "options": [
                "Give it to them",
                "Send a screenshot of it",
                "Do not share it",
                "Post it privately"
            ],
            "answer": 2
        },

        {
            "question": (
                "Before clicking a suspicious login link, you should:"
            ),
            "options": [
                "Click immediately",
                "Verify where the link came from",
                "Forward it to everyone",
                "Enter your password first"
            ],
            "answer": 1
        },

        {
            "question": (
                "Why should you avoid reusing the same password "
                "across multiple accounts?"
            ),
            "options": [
                "It makes accounts harder to remember",
                "One compromised password could expose multiple accounts",
                "It increases notifications",
                "It makes profiles public"
            ],
            "answer": 1
        }
    ],

    "Oversharing": [

        {
            "question": (
                "Why can posting 'Home alone tonight' publicly "
                "be risky?"
            ),
            "options": [
                "It gets fewer likes",
                "It reveals information about your situation",
                "It changes your password",
                "It deletes your account"
            ],
            "answer": 1
        },

        {
            "question": (
                "Someone asks what time you leave for college every "
                "morning. What is safest?"
            ),
            "options": [
                "Give the exact time",
                "Give your entire timetable",
                "Avoid sharing your routine",
                "Post it publicly"
            ],
            "answer": 2
        },

        {
            "question": (
                "What should you consider before posting personal "
                "information?"
            ),
            "options": [
                "Only how many likes it might receive",
                "Who could see it and how it could be used",
                "Whether strangers will share it",
                "How quickly you can post it"
            ],
            "answer": 1
        },

        {
            "question": (
                "Which is the safest approach to daily routines?"
            ),
            "options": [
                "Post exact times publicly",
                "Share the routine with strangers",
                "Keep predictable routines private",
                "Post your timetable"
            ],
            "answer": 2
        },

        {
            "question": (
                "If a post contains unnecessary personal information, "
                "what should you do?"
            ),
            "options": [
                "Post it publicly",
                "Add more personal information",
                "Remove or hide sensitive details",
                "Tag strangers"
            ],
            "answer": 2
        }
    ],

    "Location Safety": [

        {
            "question": (
                "What is safer when posting about a place you visited?"
            ),
            "options": [
                "Share your exact live location",
                "Share your home address",
                "Post later without revealing your exact location",
                "Share your routine"
            ],
            "answer": 2
        },

        {
            "question": (
                "Why can live-location sharing be risky?"
            ),
            "options": [
                "It increases battery life",
                "It can reveal where you currently are",
                "It makes accounts private",
                "It deletes your location"
            ],
            "answer": 1
        },

        {
            "question": (
                "Someone asks where you live. What is safest?"
            ),
            "options": [
                "Send your exact address",
                "Send your house number",
                "Avoid sharing your exact location",
                "Send your live location"
            ],
            "answer": 2
        },

        {
            "question": (
                "Which information can reveal a predictable routine?"
            ),
            "options": [
                "Favourite colour",
                "Exact daily travel times",
                "Favourite movie",
                "Music preference"
            ],
            "answer": 1
        },

        {
            "question": (
                "What should you check if you are concerned about "
                "location privacy?"
            ),
            "options": [
                "Location permissions",
                "Number of followers only",
                "Profile picture only",
                "Likes only"
            ],
            "answer": 0
        }
    ],

    "Digital Reputation": [

        {
            "question": (
                "Why should you think before posting something online?"
            ),
            "options": [
                "Posts can contribute to your long-term digital footprint",
                "Posts disappear immediately",
                "Only your friends can ever see them",
                "Nothing online can be copied"
            ],
            "answer": 0
        },

        {
            "question": (
                "You receive your college ID card. What is safest?"
            ),
            "options": [
                "Post the full ID publicly",
                "Show all personal details",
                "Avoid posting sensitive ID details",
                "Send it to strangers"
            ],
            "answer": 2
        },

        {
            "question": (
                "What should you periodically review?"
            ),
            "options": [
                "Only your profile picture",
                "Old posts and publicly visible information",
                "Only your followers' posts",
                "Only your likes"
            ],
            "answer": 1
        },

        {
            "question": (
                "Which action can help protect your digital reputation?"
            ),
            "options": [
                "Posting everything publicly",
                "Thinking about the long-term impact before posting",
                "Sharing private documents",
                "Posting when angry"
            ],
            "answer": 1
        },

        {
            "question": (
                "What can happen to content posted online?"
            ),
            "options": [
                "It is always immediately destroyed",
                "It may continue contributing to your digital footprint",
                "Nobody can copy it",
                "It automatically becomes private"
            ],
            "answer": 1
        }
    ]
}


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template(
        "home.html"
    )


# =========================================================
# CHALLENGE
# =========================================================

@app.route("/challenge")
def challenge():

    return render_template(
        "challenge.html"
    )


# =========================================================
# LEARN
# =========================================================

@app.route("/learn")
def learn():

    return render_template(
        "learn.html"
    )


# =========================================================
# ABOUT
# =========================================================

@app.route("/about")
def about():

    return render_template(
        "about.html"
    )


# =========================================================
# SUBMIT MAIN CHALLENGE
# =========================================================

@app.route(
    "/submit-challenge",
    methods=["POST"]
)
def submit_challenge():

    conn = None

    try:

        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "error": "No challenge data received."
            }), 400

        # -------------------------------------------------
        # NAME
        # -------------------------------------------------

        name = str(
            data.get("name", "")
        ).strip()

        if not name:
            name = "Guest"

        name = name[:30]

        # -------------------------------------------------
        # SCORES
        # -------------------------------------------------

        score = int(
            data.get("score", 0)
        )

        privacy = int(
            data.get("privacy", 0)
        )

        security = int(
            data.get("security", 0)
        )

        oversharing = int(
            data.get("oversharing", 0)
        )

        location = int(
            data.get("location", 0)
        )

        reputation = int(
            data.get("reputation", 0)
        )

        # -------------------------------------------------
        # KEEP EVERYTHING BETWEEN 0 AND 100
        # -------------------------------------------------

        score = max(
            0,
            min(100, score)
        )

        privacy = max(
            0,
            min(100, privacy)
        )

        security = max(
            0,
            min(100, security)
        )

        oversharing = max(
            0,
            min(100, oversharing)
        )

        location = max(
            0,
            min(100, location)
        )

        reputation = max(
            0,
            min(100, reputation)
        )

        # -------------------------------------------------
        # HABIT / RISK
        # -------------------------------------------------

        habit_type, risk_level = habit_from_score(
            score
        )

        # -------------------------------------------------
        # DATABASE
        # -------------------------------------------------

        conn = get_db()

        unique_number = int(
            time.time() * 1000000
        )

        internal_email = (
            f"participant_{unique_number}"
            "@socialsafety.local"
        )

        internal_password = generate_password_hash(
            "temporary-project-user"
        )

        # -------------------------------------------------
        # CREATE USER
        # -------------------------------------------------

        user_cursor = conn.execute(
            """
            INSERT INTO users
            (
                name,
                email,
                password,
                created_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                name,
                internal_email,
                internal_password,
                datetime.now().isoformat()
            )
        )

        user_id = user_cursor.lastrowid

        # -------------------------------------------------
        # SAVE MAIN RESULT
        # -------------------------------------------------

        result_cursor = conn.execute(
            """
            INSERT INTO quiz_results
            (
                user_id,
                score,
                risk_level,
                habit_type,
                privacy_score,
                security_score,
                oversharing_score,
                location_score,
                reputation_score,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id,
                score,
                risk_level,
                habit_type,
                privacy,
                security,
                oversharing,
                location,
                reputation,
                datetime.now().isoformat()
            )
        )

        result_id = result_cursor.lastrowid

        # Remember the latest real participant for the leaderboard.
        session["latest_result_id"] = result_id

        conn.commit()

        conn.close()

        conn = None

        return jsonify({
            "success": True,
            "id": result_id
        })

    except sqlite3.IntegrityError as error:

        if conn:
            conn.rollback()
            conn.close()

        print(
            "DATABASE INTEGRITY ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": str(error)
        }), 500

    except sqlite3.OperationalError as error:

        if conn:
            conn.rollback()
            conn.close()

        print(
            "DATABASE OPERATION ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": str(error)
        }), 500

    except (ValueError, TypeError) as error:

        if conn:
            conn.rollback()
            conn.close()

        print(
            "INVALID DATA ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": "Invalid challenge score data."
        }), 400

    except Exception as error:

        if conn:
            conn.rollback()
            conn.close()

        print(
            "GENERAL ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": str(error)
        }), 500


# =========================================================
# RESULTS
# =========================================================

@app.route(
    "/results/<int:result_id>"
)
def results(result_id):

    conn = get_db()

    row = conn.execute(
        """
        SELECT
            quiz_results.*,
            users.name AS user_name
        FROM quiz_results
        LEFT JOIN users
            ON quiz_results.user_id = users.id
        WHERE quiz_results.id = ?
        """,
        (result_id,)
    ).fetchone()

    conn.close()

    if row is None:

        flash(
            "Result not found. Please take the challenge again.",
            "error"
        )

        return redirect(
            url_for("challenge")
        )

    categories = {

        "Privacy":
            row["privacy_score"],

        "Account Security":
            row["security_score"],

        "Oversharing":
            row["oversharing_score"],

        "Location Safety":
            row["location_score"],

        "Digital Reputation":
            row["reputation_score"]
    }

    weakest = min(
        categories,
        key=categories.get
    )

    # Information used by the new learning stage
    weakest_info = WEAK_AREA_INFO.get(
        weakest,
        {
            "title": weakest,
            "description": "Focus on improving this area.",
            "advice": []
        }
    )

    return render_template(
        "results.html",
        result=row,
        categories=categories,
        weakest=weakest,
        weakest_info=weakest_info,
        user_name=row["user_name"]
    )


# =========================================================
# FINAL TEST PAGE
# =========================================================
#
# This route intentionally supports the new flow:
#
# Results
#    ↓
# Weakest area
#    ↓
# Focused learning/test
#
# It first tries final_test.html.
# If that template does not exist yet, a built-in page is
# returned so the backend will still work.
# =========================================================

@app.route(
    "/final-test/<int:result_id>"
)
def final_test(result_id):

    conn = get_db()

    row = conn.execute(
        """
        SELECT
            quiz_results.*,
            users.name AS user_name
        FROM quiz_results
        LEFT JOIN users
            ON quiz_results.user_id = users.id
        WHERE quiz_results.id = ?
        """,
        (result_id,)
    ).fetchone()

    conn.close()

    if row is None:

        flash(
            "Result not found. Please take the challenge again.",
            "error"
        )

        return redirect(
            url_for("challenge")
        )

    weakest = get_weakest_area(row)

    questions = FINAL_TEST_QUESTIONS.get(
        weakest,
        []
    )

    info = WEAK_AREA_INFO.get(
        weakest,
        {}
    )

    # -----------------------------------------------------
    # Store the result ID in the session as well.
    # This gives us an additional way to keep track of the
    # current learning session.
    # -----------------------------------------------------

    session["final_test_result_id"] = result_id
    session["final_test_weakest"] = weakest

    try:

        return render_template(
            "final_test.html",
            result=row,
            weakest=weakest,
            weakest_info=info,
            questions=questions
        )

    except Exception:

        # -------------------------------------------------
        # FALLBACK PAGE
        #
        # This means app.py can be installed before the
        # final_test.html template is added.
        # -------------------------------------------------

        return render_template_string(
            FINAL_TEST_FALLBACK_HTML,
            result=row,
            weakest=weakest,
            weakest_info=info,
            questions=questions
        )


# =========================================================
# FINAL TEST SUBMISSION
# =========================================================

@app.route(
    "/submit-final-test",
    methods=["POST"]
)
def submit_final_test():

    try:

        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "error": "No final-test data received."
            }), 400

        result_id = int(
            data.get("result_id", 0)
        )

        answers = data.get(
            "answers",
            []
        )

        if result_id <= 0:

            return jsonify({
                "success": False,
                "error": "Invalid result ID."
            }), 400

        # -------------------------------------------------
        # LOAD ORIGINAL RESULT
        # -------------------------------------------------

        conn = get_db()

        row = conn.execute(
            """
            SELECT *
            FROM quiz_results
            WHERE id = ?
            """,
            (result_id,)
        ).fetchone()

        if row is None:

            conn.close()

            return jsonify({
                "success": False,
                "error": "Original result not found."
            }), 404

        # -------------------------------------------------
        # DETERMINE WEAKEST AREA AGAIN ON SERVER
        #
        # We don't trust the browser to tell us the weakest
        # area.
        # -------------------------------------------------

        weakest = get_weakest_area(row)

        questions = FINAL_TEST_QUESTIONS.get(
            weakest,
            []
        )

        # -------------------------------------------------
        # SCORE
        # -------------------------------------------------

        correct = 0

        for index, question in enumerate(questions):

            if index >= len(answers):
                continue

            try:

                selected_answer = int(
                    answers[index]
                )

            except (ValueError, TypeError):

                continue

            if selected_answer == question["answer"]:

                correct += 1

        total_questions = len(
            questions
        )

        if total_questions:

            percentage = round(
                (
                    correct
                    / total_questions
                ) * 100
            )

        else:

            percentage = 0

        percentage = max(
            0,
            min(100, percentage)
        )

        # -------------------------------------------------
        # SAVE FINAL TEST RESULT
        # -------------------------------------------------

        conn.execute(
            """
            INSERT INTO final_test_results
            (
                quiz_result_id,
                weakest_area,
                score,
                total_questions,
                created_at
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                result_id,
                weakest,
                percentage,
                total_questions,
                datetime.now().isoformat()
            )
        )

        conn.commit()
        conn.close()

        # -------------------------------------------------
        # SESSION
        # -------------------------------------------------

        session["final_test_score"] = percentage
        session["final_test_correct"] = correct
        session["final_test_total"] = total_questions

        return jsonify({
            "success": True,
            "result_id": result_id,
            "weakest": weakest,
            "score": percentage,
            "correct": correct,
            "total": total_questions,
            "redirect": (
                f"/final-result/{result_id}"
            )
        })

    except sqlite3.Error as error:

        print(
            "FINAL TEST DATABASE ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": "Could not save final test."
        }), 500

    except Exception as error:

        print(
            "FINAL TEST ERROR:",
            error
        )

        return jsonify({
            "success": False,
            "error": str(error)
        }), 500


# =========================================================
# FINAL TEST RESULT
# =========================================================

@app.route(
    "/final-result/<int:result_id>"
)
def final_result(result_id):

    conn = get_db()

    original_result = conn.execute(
        """
        SELECT
            quiz_results.*,
            users.name AS user_name
        FROM quiz_results
        LEFT JOIN users
            ON quiz_results.user_id = users.id
        WHERE quiz_results.id = ?
        """,
        (result_id,)
    ).fetchone()

    if original_result is None:

        conn.close()

        flash(
            "Original result not found.",
            "error"
        )

        return redirect(
            url_for("challenge")
        )

    final_result_row = conn.execute(
        """
        SELECT *
        FROM final_test_results
        WHERE quiz_result_id = ?
        ORDER BY id DESC
        LIMIT 1
        """,
        (result_id,)
    ).fetchone()

    conn.close()

    if final_result_row is None:

        return redirect(
            url_for(
                "final_test",
                result_id=result_id
            )
        )

    weakest = final_result_row["weakest_area"]

    info = WEAK_AREA_INFO.get(
        weakest,
        {}
    )

    # Calculate the participant's score in the same weak area
    # before the focused test. This makes the before/after
    # comparison meaningful because both scores measure the
    # same safety category.
    previous_score_columns = {
        "Privacy": "privacy_score",
        "Account Security": "security_score",
        "Oversharing": "oversharing_score",
        "Location Safety": "location_score",
        "Digital Reputation": "reputation_score"
    }

    previous_score = int(
        original_result[
            previous_score_columns.get(weakest, "privacy_score")
        ] or 0
    )

    focused_score = int(final_result_row["score"] or 0)
    improvement = focused_score - previous_score

    # -----------------------------------------------------
    # Try normal template first.
    # -----------------------------------------------------

    try:

        return render_template(
            "final_result.html",
            result=original_result,
            final_result=final_result_row,
            weakest=weakest,
            weakest_info=info,
            previous_score=previous_score,
            improvement=improvement
        )

    except Exception:

        # -------------------------------------------------
        # Built-in fallback.
        # -------------------------------------------------

        score = final_result_row["score"]

        if score >= 80:

            message = (
                "Excellent! You demonstrated a strong "
                "understanding of this safety area."
            )

        elif score >= 60:

            message = (
                "Good job! You understand the basics, "
                "but there is still room to improve."
            )

        else:

            message = (
                "Keep learning! Review the safety advice "
                "for this area and try the test again."
            )

        return render_template_string(
            FINAL_RESULT_FALLBACK_HTML,
            result=original_result,
            final_result=final_result_row,
            weakest=weakest,
            weakest_info=info,
            message=message,
            previous_score=previous_score,
            improvement=improvement
        )


# =========================================================
# LEADERBOARD
# =========================================================

@app.route("/leaderboard")
def leaderboard():
    """Display the real challenge leaderboard.

    The existing database schema stores participant names in the users
    table, not in quiz_results. Multiple attempts by the same participant
    are grouped by name and only their best score is displayed. No demo
    participants are inserted.
    """

    conn = get_db()

    # The project database stores the participant's name in users.name.
    # Do NOT reference quiz_results.user_name because that column does not
    # exist in the existing database.
    rows = conn.execute(
        """
        SELECT
            MIN(q.id) AS result_id,
            MAX(TRIM(COALESCE(u.name, 'Guest'))) AS name,
            MAX(COALESCE(q.score, 0)) AS score
        FROM quiz_results AS q
        LEFT JOIN users AS u
            ON q.user_id = u.id
        WHERE TRIM(COALESCE(u.name, '')) != ''
        GROUP BY LOWER(TRIM(u.name))
        ORDER BY MAX(COALESCE(q.score, 0)) DESC, MIN(q.id) ASC
        """
    ).fetchall()

    # Identify the result that led the participant to the leaderboard.
    current_result_id = None

    result_parameter = request.args.get("result")
    if result_parameter:
        try:
            current_result_id = int(result_parameter)
        except (ValueError, TypeError):
            current_result_id = None

    if current_result_id is None:
        current_result_id = session.get("latest_result_id")

    current_name = None

    if current_result_id:
        current_row = conn.execute(
            """
            SELECT TRIM(COALESCE(u.name, '')) AS name
            FROM quiz_results AS q
            LEFT JOIN users AS u
                ON q.user_id = u.id
            WHERE q.id = ?
            """,
            (current_result_id,)
        ).fetchone()

        if current_row and current_row["name"]:
            current_name = current_row["name"].strip().lower()

    conn.close()

    board = []

    for index, row in enumerate(rows, start=1):
        name = (row["name"] or "Guest").strip()

        board.append({
            "rank": index,
            "name": name,
            "score": int(row["score"] or 0),
            "result_id": row["result_id"],
            "is_you": bool(
                current_name
                and name.lower() == current_name
            )
        })

    my_rank = None

    for person in board:
        if person["is_you"]:
            my_rank = person["rank"]
            break

    return render_template(
        "leaderboard.html",
        board=board,
        my_rank=my_rank
    )


# =========================================================
# ADMIN LOGIN
# =========================================================

@app.route(
    "/admin/login",
    methods=["GET", "POST"]
)
def admin_login():

    if request.method == "POST":

        username = request.form.get(
            "username"
        )

        password = request.form.get(
            "password"
        )

        conn = get_db()

        admin = conn.execute(
            """
            SELECT *
            FROM admin
            WHERE username = ?
            """,
            (username,)
        ).fetchone()

        conn.close()

        if (
            admin
            and check_password_hash(
                admin["password"],
                password
            )
        ):

            session["is_admin"] = True

            session["admin_username"] = username

            flash(
                "Logged in successfully.",
                "success"
            )

            return redirect(
                url_for("admin_dashboard")
            )

        flash(
            "Invalid username or password.",
            "error"
        )

    return render_template(
        "auth.html"
    )


# =========================================================
# ADMIN LOGOUT
# =========================================================

@app.route(
    "/admin/logout"
)
def admin_logout():

    session.pop(
        "is_admin",
        None
    )

    session.pop(
        "admin_username",
        None
    )

    flash(
        "Logged out.",
        "success"
    )

    return redirect(
        url_for("home")
    )


# =========================================================
# ADMIN DASHBOARD
# =========================================================

@app.route("/admin")
@admin_required
def admin_dashboard():

    conn = get_db()

    # Website challenge results remain available separately.
    total_participants = conn.execute(
        "SELECT COUNT(*) FROM quiz_results"
    ).fetchone()[0]

    avg_score_row = conn.execute(
        "SELECT AVG(score) FROM quiz_results"
    ).fetchone()[0]

    avg_score = round(avg_score_row) if avg_score_row else 0

    high_risk_users = conn.execute(
        """
        SELECT COUNT(*)
        FROM quiz_results
        WHERE risk_level = 'High Risk'
        """
    ).fetchone()[0]

    # -----------------------------------------------------
    # FOCUSED TEST / LEARNING IMPACT
    # -----------------------------------------------------
    # The focused test is tied to the original quiz result.
    # We keep the latest focused-test attempt for each result
    # so the dashboard never mixes multiple attempts together.
    focused_rows = conn.execute(
        """
        SELECT
            q.id AS result_id,
            TRIM(COALESCE(u.name, 'Guest')) AS name,
            COALESCE(q.score, 0) AS main_score,
            COALESCE(q.privacy_score, 0) AS privacy_score,
            COALESCE(q.security_score, 0) AS security_score,
            COALESCE(q.oversharing_score, 0) AS oversharing_score,
            COALESCE(q.location_score, 0) AS location_score,
            COALESCE(q.reputation_score, 0) AS reputation_score,
            f.weakest_area,
            f.score AS focused_score,
            f.total_questions,
            f.created_at AS focused_created_at
        FROM quiz_results AS q
        LEFT JOIN users AS u
            ON q.user_id = u.id
        LEFT JOIN final_test_results AS f
            ON f.id = (
                SELECT MAX(f2.id)
                FROM final_test_results AS f2
                WHERE f2.quiz_result_id = q.id
            )
        ORDER BY q.id DESC
        """
    ).fetchall()

    focused_completed = 0
    focused_score_total = 0
    improvement_total = 0
    improvement_count = 0
    previous_score_total = 0
    focused_participants = []

    category_columns = {
        "Privacy": "privacy_score",
        "Account Security": "security_score",
        "Oversharing": "oversharing_score",
        "Location Safety": "location_score",
        "Digital Reputation": "reputation_score"
    }

    for focused_row in focused_rows:

        focused_score = focused_row["focused_score"]
        weakest_area = focused_row["weakest_area"]

        if focused_score is not None and weakest_area:

            focused_completed += 1
            focused_score = int(focused_score or 0)
            focused_score_total += focused_score

            initial_score = int(
                focused_row[category_columns.get(weakest_area, "privacy_score")] or 0
            )

            improvement = focused_score - initial_score
            improvement_total += improvement
            improvement_count += 1
            previous_score_total += initial_score

            focused_participants.append({
                "result_id": focused_row["result_id"],
                "name": focused_row["name"] or "Guest",
                "main_score": int(focused_row["main_score"] or 0),
                "weakest_area": weakest_area,
                "previous_score": initial_score,
                "focused_score": focused_score,
                "improvement": improvement,
                "total_questions": int(focused_row["total_questions"] or 0),
                "created_at": focused_row["focused_created_at"] or ""
            })

    focused_average = (
        round(focused_score_total / focused_completed)
        if focused_completed
        else 0
    )

    average_improvement = (
        round(improvement_total / improvement_count)
        if improvement_count
        else 0
    )

    average_previous_score = (
        round(previous_score_total / improvement_count)
        if improvement_count
        else 0
    )

    focused_completion_rate = (
        round((focused_completed / total_participants) * 100)
        if total_participants
        else 0
    )

    risk_counts = {
        "High Risk": 0,
        "Moderate Risk": 0,
        "Low Risk": 0
    }

    risk_rows = conn.execute(
        """
        SELECT risk_level, COUNT(*) AS count
        FROM quiz_results
        GROUP BY risk_level
        """
    ).fetchall()

    for row in risk_rows:
        if row["risk_level"] in risk_counts:
            risk_counts[row["risk_level"]] = row["count"]

    total_risk = sum(risk_counts.values()) or 1

    risk_percent = {
        key: round((value / total_risk) * 100)
        for key, value in risk_counts.items()
    }

    conn.close()

    # REAL Google Form response analysis.
    survey = load_survey_data()

    return render_template(
        "admin_dashboard.html",
        total_participants=total_participants,
        avg_score=avg_score,
        high_risk_users=high_risk_users,
        quizzes_taken=total_participants,
        risk_percent=risk_percent,
        survey=survey,
        focused_completed=focused_completed,
        focused_average=focused_average,
        average_improvement=average_improvement,
        average_previous_score=average_previous_score,
        focused_completion_rate=focused_completion_rate,
        focused_participants=focused_participants
    )


# =========================================================
# BUILT-IN FINAL TEST FALLBACK
# =========================================================
#
# This is only used if final_test.html has not been created.
# Once we add your proper final_test.html, this will not be
# used.
# =========================================================

FINAL_TEST_FALLBACK_HTML = """

<!DOCTYPE html>

<html lang="en">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>Weak Area Test — Social Media Safety</title>

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    background: #05060d;
    color: #f4f4f5;
    font-family: Arial, sans-serif;
}

.wrapper {
    max-width: 850px;
    margin: auto;
    padding: 50px 20px 80px;
}

.header {
    text-align: center;
    margin-bottom: 35px;
}

.header h1 {
    font-size: 2.3rem;
    margin-bottom: 10px;
}

.header p {
    color: #a1a1aa;
    line-height: 1.6;
}

.card {
    background: #101221;
    border: 1px solid #292c43;
    border-radius: 20px;
    padding: 30px;
    margin-bottom: 20px;
}

.weak {
    background: rgba(239, 68, 68, 0.08);
    border: 1px solid rgba(239, 68, 68, 0.25);
    padding: 20px;
    border-radius: 14px;
    margin-bottom: 25px;
}

.weak strong {
    color: #f87171;
}

.question {
    margin-bottom: 30px;
}

.question h3 {
    line-height: 1.5;
    margin-bottom: 15px;
}

.option {
    display: block;
    padding: 14px;
    margin: 10px 0;
    background: #141628;
    border: 1px solid #30334d;
    border-radius: 10px;
    cursor: pointer;
}

.option:hover {
    border-color: #8b5cf6;
}

.option input {
    margin-right: 10px;
}

button {
    width: 100%;
    padding: 16px;
    border: none;
    border-radius: 12px;
    background: linear-gradient(
        90deg,
        #7c3aed,
        #a855f7
    );
    color: white;
    font-size: 1rem;
    font-weight: bold;
    cursor: pointer;
}

button:hover {
    transform: translateY(-1px);
}

</style>

</head>

<body>

<div class="wrapper">

<div class="header">

<h1>
    Focused Safety Test
</h1>

<p>
    Your original challenge identified
    <strong>{{ weakest }}</strong>
    as your weakest area.
</p>

</div>

<div class="card">

<div class="weak">

<strong>
    Focus Area: {{ weakest }}
</strong>

<p>
    {{ weakest_info.description }}
</p>

{% if weakest_info.advice %}

<ul>

{% for item in weakest_info.advice %}

<li>{{ item }}</li>

{% endfor %}

</ul>

{% endif %}

</div>

<form id="finalTestForm">

{% for question in questions %}

<div class="question">

<h3>
    {{ loop.index }}.
    {{ question.question }}
</h3>

{% for option in question.options %}

<label class="option">

<input
    type="radio"
    name="question_{{ loop.index0 }}"
    value="{{ loop.index0 }}"
    required
>

{{ option }}

</label>

{% endfor %}

</div>

{% endfor %}

<button type="submit">
    Finish Test
</button>

</form>

</div>

</div>

<script>

document
.getElementById("finalTestForm")
.addEventListener("submit", async function(event) {

    event.preventDefault();

    const answers = [];

    {% for question in questions %}

    const selected{{ loop.index }} =
        document.querySelector(
            'input[name="question_{{ loop.index0 }}"]:checked'
        );

    answers.push(
        selected{{ loop.index }}
            ? parseInt(selected{{ loop.index }}.value)
            : -1
    );

    {% endfor %}

    const response = await fetch(
        "/submit-final-test",
        {
            method: "POST",

            headers: {
                "Content-Type":
                    "application/json"
            },

            body: JSON.stringify({
                result_id:
                    {{ result["id"] }},

                answers:
                    answers
            })
        }
    );

    const data =
        await response.json();

    if (
        data.success &&
        data.redirect
    ) {

        window.location.href =
            data.redirect;

    } else {

        alert(
            data.error ||
            "Something went wrong."
        );

    }

});

</script>

</body>

</html>

"""


# =========================================================
# BUILT-IN FINAL RESULT FALLBACK
# =========================================================

FINAL_RESULT_FALLBACK_HTML = """

<!DOCTYPE html>

<html lang="en">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>Final Test Result — Social Media Safety</title>

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    background: #05060d;
    color: #f4f4f5;
    font-family: Arial, sans-serif;
}

.wrapper {
    max-width: 850px;
    margin: auto;
    padding: 70px 20px;
}

.card {
    background: #101221;
    border: 1px solid #292c43;
    border-radius: 22px;
    padding: 45px;
    text-align: center;
}

.score {
    font-size: 5rem;
    font-weight: 800;
    margin: 20px 0;
}

.area {
    color: #a855f7;
    font-weight: bold;
}

.message {
    color: #a1a1aa;
    line-height: 1.7;
}

.buttons {
    margin-top: 30px;
}

a {
    display: inline-block;
    padding: 14px 22px;
    border-radius: 12px;
    background: linear-gradient(
        90deg,
        #7c3aed,
        #a855f7
    );
    color: white;
    text-decoration: none;
    font-weight: bold;
}

</style>

</head>

<body>

<div class="wrapper">

<div class="card">

<h1>
    Test Complete 🎉
</h1>

<p>
    You completed the focused test for
</p>

<div class="area">
    {{ weakest }}
</div>

<div class="score">
    {{ final_result["score"] }}%
</div>

<p>
    {{ final_result["correct"] if "correct" in final_result.keys() else "" }}
</p>

<p class="message">
    {{ message }}
</p>

<div class="buttons">

<a href="/results/{{ result['id'] }}">
    Back to My Results
</a>

</div>

</div>

</div>

</body>

</html>

"""


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def page_not_found(error):

    return """
    <h1>404 - Page Not Found</h1>
    <p>The page you requested does not exist.</p>
    """, 404


# =========================================================
# RUN APPLICATION
# =========================================================

if __name__ == "__main__":

    init_db()

    app.run(
        debug=True,
        threaded=False
    )
