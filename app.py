import os
import io
import time
import json
import base64
import random
import hmac
import hashlib
import logging
import secrets as pysecrets
import sqlite3
import tempfile
import textwrap
import concurrent.futures
from datetime import datetime, timedelta
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from google.genai import types
import pypdf
from PIL import Image, ImageDraw

# محاولة استيراد مكتبات Word و PowerPoint بأمان تامة لضمان عدم انهيار السيرفر
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# محاولة استيراد مكتبة FPDF لتوليد ملفات الـ PDF بأمان
try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# محاولة استيراد مكتبات تشكيل النص العربي (ضرورية لعرض العربية بشكل صحيح داخل PDF)
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_SHAPING_AVAILABLE = True
except ImportError:
    ARABIC_SHAPING_AVAILABLE = False

# محاولة استيراد requests لدعم التكامل الاختياري مع Canva (Canva Connect API)
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# محاولة استيراد openpyxl لتصدير نموذج التصحيح التلقائي كملف Excel
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# === إصلاح: ميزة "استمع للورقة" (تحويل نص لصوت) للطلاب ذوي الإعاقة البصرية — كانت
# مطلوبة سابقاً لكن لم تكن موجودة فعلياً في الكود المنشور (تحقّقنا من الملف الأصلي
# ومن كل نسخة سابقة ولم نجدها)، فتمت إضافتها الآن من الصفر عبر مكتبة gTTS (مجانية
# ولا تحتاج مفتاح API، تعتمد فقط على اتصال إنترنت بسيط). ===
try:
    from gtts import gTTS
    GTTS_AVAILABLE = True
except ImportError:
    GTTS_AVAILABLE = False

# =========================================================================================
# === سجلّ أخطاء داخلي (Server-side logging) ===
# أي تفصيل خطأ حسّاس (رسالة اتصال قاعدة بيانات، استثناء داخلي من مكتبة، إلخ) يُطبع هنا
# فقط (يظهر بسجلات الخادم / "Manage app" -> Logs على Streamlit Cloud) ولا يُعرض أبداً
# لأي زائر أو معلم بواجهة التطبيق — الواجهة تعرض دائماً رسالة عامة آمنة بدلاً منه.
# =========================================================================================
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("edu_worksheet_adapt")


def log_internal_error(context, exc):
    """يطبع تفاصيل الخطأ الكاملة بسجلات الخادم فقط، دون كشفها للمستخدم النهائي."""
    logger.error("%s: %r", context, exc)


# =========================================================================================
# === طبقة قاعدة بيانات لحفظ "ذاكرة الطالب" — تدعم Supabase (PostgreSQL) كتخزين دائم
# لا يُمسح عند إعادة النشر أو نوم التطبيق، مع رجوع تلقائي لملف SQLite محلي مؤقت فقط
# إن لم يُضبط اتصال Supabase بعد (SUPABASE_DB_URL بالـ Secrets).
#
# === إصلاح أداء مهم (كان السبب الأكبر في بطء التطبيق) ===
# كان الكود القديم يفتح اتصال Postgres/Supabase جديد بالكامل (مصافحة شبكة + SSL)
# في كل استدعاء لأي دالة قاعدة بيانات — وهذا يحصل في كل ضغطة زر أو تغيير قائمة
# منسدلة لأن ستريمليت يعيد تشغيل السكربت بالكامل (rerun). الآن الاتصال يُفتح مرة
# واحدة فقط ويُعاد استخدامه عبر st.cache_resource، مع فحص خفيف (SELECT 1) للتأكد
# أنه لا يزال حياً وإعادة إنشائه تلقائياً فقط إذا انقطع فعلاً.
# =========================================================================================
try:
    import psycopg2
    PSYCOPG2_AVAILABLE = True
except ImportError:
    PSYCOPG2_AVAILABLE = False

SUPABASE_DB_URL = None
try:
    SUPABASE_DB_URL = st.secrets.get("SUPABASE_DB_URL", None)
except Exception:
    SUPABASE_DB_URL = os.getenv("SUPABASE_DB_URL")

USE_POSTGRES = bool(SUPABASE_DB_URL and PSYCOPG2_AVAILABLE)
DB_PATH = "edu_adapt_data.db"  # يُستخدم فقط كتخزين احتياطي مؤقت (غير دائم على الاستضافة السحابية)

# =========================================================================================
# === ربط الاشتراك المدفوع على Whop (منتج EWAS) بحسابات المعلمين: كل حساب جديد لازم
# يدخل "مفتاح ترخيص" (License Key) يستلمه من Whop بعد الدفع، ويتحقق التطبيق منه مباشرة
# عبر واجهة Whop API قبل السماح بإنشاء الحساب، وبشكل دوري (كل 24 ساعة تقريباً) عند
# الدخول للتأكد أن الاشتراك ما زال فعالاً (ولم يُلغَ أو ينتهِ). الحسابات القديمة التي
# أُنشئت قبل هذه الميزة (بدون مفتاح ترخيص محفوظ) تبقى تعمل بلا قيود لتفادي كسرها. ===
# =========================================================================================
WHOP_API_KEY = None
try:
    WHOP_API_KEY = st.secrets.get("WHOP_API_KEY", None)
except Exception:
    WHOP_API_KEY = os.getenv("WHOP_API_KEY")

WHOP_API_BASE = "https://api.whop.com/api/v1"
MONTHLY_WORKSHEET_LIMIT = 50  # الحد الأقصى لعدد أوراق العمل المولّدة شهرياً لكل حساب مشترك


@st.cache_resource(show_spinner=False)
def _get_pg_connection():
    """
    يفتح اتصال Postgres/Supabase مرة واحدة فقط طوال عمر التطبيق (بفضل
    st.cache_resource) بدل فتح اتصال جديد بكل استدعاء دالة. هذا هو الإصلاح
    الأهم لبطء التطبيق: فتح اتصال جديد بكل rerun كان يضيف تأخيراً ملحوظاً
    ومتكرراً بسبب مصافحة الشبكة وSSL مع خادم Supabase البعيد.
    """
    return psycopg2.connect(SUPABASE_DB_URL, sslmode="require")


def _ensure_pg_alive(conn):
    """يتحقق أن الاتصال المخزَّن ما زال حياً بفحص خفيف جداً، ولا يعيد إنشاءه إلا لو انقطع فعلاً."""
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
        return conn
    except Exception:
        try:
            conn.close()
        except Exception:
            pass
        _get_pg_connection.clear()
        return _get_pg_connection()


def get_db():
    if USE_POSTGRES:
        conn = _get_pg_connection()
        return _ensure_pg_alive(conn)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def release_db(conn):
    """
    استخدم هذه الدالة بدل conn.close() مباشرة بعد كل استخدام لقاعدة البيانات.
    بالنسبة لـ SQLite (تخزين مؤقت محلي) تُغلق الاتصال كالمعتاد لأنه رخيص الإنشاء.
    بالنسبة لـ Postgres/Supabase لا تُغلق الاتصال المشترك المخزَّن مؤقتاً — إبقاؤه
    مفتوحاً بين الاستدعاءات هو بيت القصيد من التخزين المؤقت أعلاه.
    """
    if not USE_POSTGRES:
        try:
            conn.close()
        except Exception:
            pass


def _q(sql):
    """يحوّل صيغة الـ placeholders من ? (SQLite) إلى %s (PostgreSQL) عند الحاجة."""
    return sql.replace("?", "%s") if USE_POSTGRES else sql


def _rows_to_dicts(cur, rows):
    if USE_POSTGRES:
        colnames = [desc[0] for desc in cur.description]
        return [dict(zip(colnames, row)) for row in rows]
    return [dict(r) for r in rows]


def _row_to_dict(cur, row):
    if row is None:
        return None
    if USE_POSTGRES:
        colnames = [desc[0] for desc in cur.description]
        return dict(zip(colnames, row))
    return dict(row)


def init_db():
    conn = get_db()
    cur = conn.cursor()
    id_column = "id SERIAL PRIMARY KEY" if USE_POSTGRES else "id INTEGER PRIMARY KEY AUTOINCREMENT"

    cur.execute("""
        CREATE TABLE IF NOT EXISTS teachers (
            teacher_name TEXT PRIMARY KEY,
            pin_hash TEXT NOT NULL,
            created_at TEXT,
            failed_attempts INTEGER DEFAULT 0,
            locked_until TEXT
        )
    """)
    # === ترقية الجداول القديمة: إضافة أعمدة قفل الحساب وأعمدة مفتاح ترخيص Whop لو الجدول
    # أُنشئ قبل هذه الإصلاحات ===
    for column_def in (
        "failed_attempts INTEGER DEFAULT 0", "locked_until TEXT",
        "license_key TEXT", "license_status TEXT", "license_checked_at TEXT"
    ):
        try:
            cur.execute(f"ALTER TABLE teachers ADD COLUMN {column_def}")
            conn.commit()
        except Exception:
            conn.rollback() if USE_POSTGRES else None
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS students (
            {id_column},
            teacher_name TEXT NOT NULL,
            full_name TEXT NOT NULL,
            grade TEXT,
            system TEXT,
            category TEXT,
            condition_text TEXT,
            created_at TEXT
        )
    """)
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS worksheet_history (
            {id_column},
            teacher_name TEXT NOT NULL,
            student_id INTEGER,
            student_name TEXT,
            subject TEXT,
            grade TEXT,
            adaptation_level TEXT,
            mode TEXT,
            adapted_text TEXT,
            answer_key_json TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    release_db(conn)


DB_INIT_ERROR = None
try:
    init_db()
except Exception as e:
    # === إصلاح أمني: لا نخزّن نص الاستثناء الخام (قد يحوي اسم مضيف/منفذ/بيانات اتصال
    # قاعدة البيانات) في متغيّر يُعرض لاحقاً لأي زائر قبل تسجيل الدخول. التفاصيل
    # الكاملة تُطبع فقط بسجلات الخادم، وDB_INIT_ERROR يصبح مجرد علم (True/False). ===
    log_internal_error("فشل تهيئة قاعدة البيانات", e)
    DB_INIT_ERROR = True
    if USE_POSTGRES:
        # فشل الاتصال الفعلي بـ Supabase رغم توفر الإعداد — نرجع تلقائياً للتخزين
        # المحلي المؤقت بدل أن ينهار التطبيق بالكامل
        USE_POSTGRES = False
        try:
            init_db()
        except Exception as e2:
            log_internal_error("فشل تهيئة قاعدة البيانات المحلية الاحتياطية أيضاً", e2)


# =========================================================================================
# === تجزئة كلمات المرور: PBKDF2-HMAC-SHA256 بملح عشوائي فريد لكل حساب (معيار NIST
# القياسي لتخزين كلمات المرور)، بدل sha256 بسيط بلا ملح (كان عرضة لجداول Rainbow
# Tables الجاهزة، وكان يعطي نفس القيمة لنفس كلمة المرور عبر كل الحسابات). صيغة
# التخزين: "pbkdf2_sha256$<عدد التكرارات>$<الملح بصيغة hex>$<الناتج بصيغة hex>".
# ===
# === ترقية تلقائية للحسابات القديمة: أي حساب مسجَّل بالنظام القديم (sha256 بلا ملح،
# يظهر كسلسلة hex طولها 64 محرفاً بالضبط بدون "$") يُتعرَّف عليه تلقائياً عند محاولة
# الدخول، ويُرقَّى فوراً للتجزئة الجديدة بعد أول دخول ناجح فقط، دون أي إزعاج للمعلم
# ودون أي حاجة لإعادة تسجيل. ===
# =========================================================================================
_PBKDF2_ITERATIONS = 260_000


def _hash_password(password_text, salt_hex=None):
    if salt_hex is None:
        salt_hex = pysecrets.token_hex(16)
    salt_bytes = bytes.fromhex(salt_hex)
    derived = hashlib.pbkdf2_hmac(
        "sha256", password_text.encode("utf-8"), salt_bytes, _PBKDF2_ITERATIONS
    )
    return f"pbkdf2_sha256${_PBKDF2_ITERATIONS}${salt_hex}${derived.hex()}"


def _is_legacy_plain_sha256(stored_hash):
    """يتعرّف على تنسيق التخزين القديم: sha256 بلا ملح (hex بطول 64 بدون أي '$')."""
    return "$" not in stored_hash and len(stored_hash) == 64


def _verify_password(password_text, stored_hash):
    """
    يتحقق من كلمة المرور مقابل القيمة المخزَّنة، بدعم التنسيقين معاً (الجديد والقديم).
    يعيد tuple: (is_valid: bool, needs_upgrade: bool).
    """
    if _is_legacy_plain_sha256(stored_hash):
        legacy_hash = hashlib.sha256(password_text.encode("utf-8")).hexdigest()
        is_valid = hmac.compare_digest(legacy_hash, stored_hash)
        return is_valid, is_valid  # يُرقّى فوراً لو كانت صحيحة

    try:
        _, iterations_str, salt_hex, _ = stored_hash.split("$")
        candidate = _hash_password(password_text, salt_hex=salt_hex)
        return hmac.compare_digest(candidate, stored_hash), False
    except Exception:
        return False, False


def _is_valid_password(password_text):
    """
    === إصلاح أمني: كلمة مرور من ٤ أرقام فقط (١٠,٠٠٠ احتمال) كانت سهلة التخمين جداً
    بهجوم تجربة شامل آلي (Brute-force). الآن يجب أن تكون ٦ خانات على الأقل، وتحتوي
    حرفاً واحداً على الأقل ورقماً واحداً على الأقل معاً (وليس أرقاماً فقط). ===
    """
    if len(password_text) < 6:
        return False
    has_letter = any(ch.isalpha() for ch in password_text)
    has_digit = any(ch.isdigit() for ch in password_text)
    return has_letter and has_digit


PASSWORD_REQUIREMENT_MSG = "كلمة المرور يجب أن تتكوّن من ٦ خانات على الأقل، وتحتوي حروفاً وأرقاماً معاً."

# === قفل الحساب المؤقت بعد محاولات دخول خاطئة متكررة (حماية من هجمات التخمين الآلي) ===
MAX_FAILED_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_MINUTES = 15

# === مدة صلاحية آخر تحقق من حالة الاشتراك قبل إعادة السؤال من Whop عند تسجيل الدخول،
# لتجنّب استدعاء الـ API في كل مرة (اشتراك ما بيتغيّر كل دقيقة). ===
LICENSE_RECHECK_HOURS = 24


def _whop_check_license(license_key):
    """
    يتحقق من صلاحية مفتاح ترخيص Whop عبر واجهة Whop API (GET /memberships/{license_key}).
    يعيد tuple: (is_active: bool, status: str|None, error_message: str|None).
    error_message تُعرض للمستخدم فقط لو تعذّر التحقق تقنياً (لا يعني بالضرورة أن المفتاح خطأ).
    """
    license_key = (license_key or "").strip()
    if not license_key:
        return False, None, "الرجاء إدخال مفتاح الترخيص."
    if not REQUESTS_AVAILABLE or not WHOP_API_KEY:
        log_internal_error("تعذّر التحقق من الترخيص", "WHOP_API_KEY أو مكتبة requests غير متاحة")
        return False, None, "خدمة التحقق من الاشتراك غير متاحة حالياً. حاول لاحقاً أو تواصل مع الدعم."
    try:
        resp = requests.get(
            f"{WHOP_API_BASE}/memberships/{license_key}",
            headers={"Authorization": f"Bearer {WHOP_API_KEY}"},
            timeout=10,
        )
        if resp.status_code == 404:
            return False, None, "مفتاح الترخيص غير صحيح. تأكد من نسخه كاملاً من رسالة/صفحة الشراء على Whop."
        if resp.status_code != 200:
            log_internal_error(f"فشل التحقق من مفتاح الترخيص (HTTP {resp.status_code})", resp.text[:300])
            return False, None, "تعذّر التحقق من مفتاح الترخيص حالياً. حاول مرة أخرى بعد قليل."
        data = resp.json()
        status = data.get("status")
        is_active = status in ("active", "trialing")
        return is_active, status, None
    except Exception as e:
        log_internal_error("خطأ أثناء الاتصال بخدمة Whop للتحقق من الترخيص", e)
        return False, None, "تعذّر الاتصال بخدمة التحقق من الاشتراك حالياً. حاول مرة أخرى بعد قليل."


def register_teacher(username, password, license_key):
    """
    إنشاء حساب معلم جديد. يرفض الطلب لو اسم المستخدم محجوز مسبقاً، أو لو مفتاح الترخيص
    غير صحيح/غير فعّال، أو مستخدم مسبقاً بحساب آخر (كل مفتاح = اشتراك واحد = حساب واحد).
    يعيد tuple: (success: bool, message: str)
    """
    username = username.strip()
    password = password.strip()
    license_key = (license_key or "").strip()

    if not username or not password:
        return False, "الرجاء إدخال اسم المستخدم وكلمة المرور معاً."
    if not _is_valid_password(password):
        return False, PASSWORD_REQUIREMENT_MSG
    if not license_key:
        return False, "الرجاء إدخال مفتاح الترخيص (License Key) الذي استلمته بعد الاشتراك في EWAS على Whop."

    is_active, status, err = _whop_check_license(license_key)
    if err:
        return False, err
    if not is_active:
        return False, "مفتاح الترخيص هذا غير فعّال حالياً (الاشتراك منتهي أو مُلغى). الرجاء التجديد من صفحة EWAS على Whop."

    conn = get_db()
    cur = conn.cursor()
    cur.execute(_q("SELECT teacher_name FROM teachers WHERE teacher_name = ?"), (username,))
    if cur.fetchone() is not None:
        release_db(conn)
        return False, "اسم المستخدم هذا محجوز مسبقاً. الرجاء اختيار اسم آخر أو تسجيل الدخول."

    cur.execute(_q("SELECT teacher_name FROM teachers WHERE license_key = ?"), (license_key,))
    if cur.fetchone() is not None:
        release_db(conn)
        return False, "هذا مفتاح الترخيص مستخدم مسبقاً بحساب آخر. كل مفتاح ترخيص يمكن استخدامه لحساب واحد فقط."

    cur.execute(
        _q("""INSERT INTO teachers
           (teacher_name, pin_hash, created_at, failed_attempts, license_key, license_status, license_checked_at)
           VALUES (?, ?, ?, 0, ?, ?, ?)"""),
        (username, _hash_password(password), datetime.now().isoformat(),
         license_key, status, datetime.now().isoformat())
    )
    conn.commit()
    release_db(conn)
    return True, "تم إنشاء الحساب بنجاح! يمكنك الآن تسجيل الدخول به."


def login_teacher(username, password):
    """
    تسجيل دخول معلم موجود مسبقاً. يرفض إن لم يوجد الحساب، أو كانت كلمة المرور خاطئة،
    أو كان الحساب مقفلاً مؤقتاً بسبب محاولات خاطئة متكررة.
    يعيد tuple: (success: bool, message: str)
    """
    username = username.strip()
    password = password.strip()

    if not username or not password:
        return False, "الرجاء إدخال اسم المستخدم وكلمة المرور معاً."

    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        _q("""SELECT pin_hash, failed_attempts, locked_until, license_key, license_status,
              license_checked_at FROM teachers WHERE teacher_name = ?"""),
        (username,)
    )
    row = _row_to_dict(cur, cur.fetchone())

    if row is None:
        release_db(conn)
        return False, "لا يوجد حساب بهذا الاسم. الرجاء إنشاء حساب جديد أولاً."

    locked_until_raw = row.get("locked_until")
    if locked_until_raw:
        try:
            locked_until = datetime.fromisoformat(locked_until_raw)
        except Exception:
            locked_until = None
        if locked_until and datetime.now() < locked_until:
            remaining_minutes = max(1, int((locked_until - datetime.now()).total_seconds() // 60) + 1)
            release_db(conn)
            return False, f"تم قفل الحساب مؤقتاً بسبب محاولات دخول خاطئة متكررة. حاول مرة أخرى بعد {remaining_minutes} دقيقة."

    is_valid, needs_upgrade = _verify_password(password, row["pin_hash"])

    if not is_valid:
        failed_attempts = (row.get("failed_attempts") or 0) + 1
        if failed_attempts >= MAX_FAILED_LOGIN_ATTEMPTS:
            locked_until = (datetime.now() + timedelta(minutes=LOCKOUT_DURATION_MINUTES)).isoformat()
            cur.execute(
                _q("UPDATE teachers SET failed_attempts = ?, locked_until = ? WHERE teacher_name = ?"),
                (failed_attempts, locked_until, username)
            )
            conn.commit()
            release_db(conn)
            return False, f"تم قفل الحساب مؤقتاً {LOCKOUT_DURATION_MINUTES} دقيقة بعد {MAX_FAILED_LOGIN_ATTEMPTS} محاولات خاطئة متتالية."
        cur.execute(
            _q("UPDATE teachers SET failed_attempts = ? WHERE teacher_name = ?"),
            (failed_attempts, username)
        )
        conn.commit()
        release_db(conn)
        return False, "كلمة المرور غير صحيحة."

    # === التحقق الدوري من حالة الاشتراك على Whop (فقط للحسابات التي أُنشئت بمفتاح ترخيص —
    # الحسابات القديمة بدون مفتاح تبقى تعمل بلا قيود). لا نستدعي Whop في كل دخول، بل فقط
    # لو مر أكثر من LICENSE_RECHECK_HOURS ساعة منذ آخر تحقق، لتقليل عدد الاستدعاءات. ===
    license_key = row.get("license_key")
    license_status = row.get("license_status")
    if license_key:
        needs_recheck = True
        checked_at_raw = row.get("license_checked_at")
        if checked_at_raw:
            try:
                needs_recheck = (datetime.now() - datetime.fromisoformat(checked_at_raw)) > timedelta(hours=LICENSE_RECHECK_HOURS)
            except Exception:
                needs_recheck = True
        if needs_recheck:
            is_active, fresh_status, err = _whop_check_license(license_key)
            if not err:
                license_status = fresh_status
                cur.execute(
                    _q("UPDATE teachers SET license_status = ?, license_checked_at = ? WHERE teacher_name = ?"),
                    (fresh_status, datetime.now().isoformat(), username)
                )
                conn.commit()
        if license_status not in ("active", "trialing"):
            release_db(conn)
            return False, "انتهى اشتراكك في EWAS أو تم إلغاؤه. الرجاء تجديد الاشتراك من صفحة المنتج على Whop للاستمرار."

    # نجاح الدخول: تصفير عدّاد المحاولات الخاطئة وفكّ القفل، وترقية التجزئة القديمة إن لزم
    new_hash = _hash_password(password) if needs_upgrade else row["pin_hash"]
    cur.execute(
        _q("UPDATE teachers SET failed_attempts = 0, locked_until = NULL, pin_hash = ? WHERE teacher_name = ?"),
        (new_hash, username)
    )
    conn.commit()
    release_db(conn)
    return True, "تم تسجيل الدخول بنجاح."


# =========================================================================================
# === دوال منطقية "نقية" (لا تعتمد على واجهة ستريمليت أو حالة الجلسة) خاصة بمعالجة نص
# الذكاء الاصطناعي — مرفوعة إلى مستوى الملف (بدل تعريفها داخل الكتلة الرئيسية للواجهة)
# حتى يمكن اختبارها آلياً (راجع tests/test_worksheet_helpers.py) دون الحاجة لتشغيل
# ستريمليت كاملاً أو توفير مفتاح GOOGLE_API_KEY وهمي. ===
# =========================================================================================
def parse_ai_sections(full_text):
    """
    تفصل استجابة الذكاء الاصطناعي الواحدة إلى ثلاثة أجزاء: نص ورقة العمل الرئيسي،
    نموذج الإجابات (JSON)، وقائمة المفردات الأساسية — دون الحاجة لاستدعاء إضافي منفصل
    للنموذج، توفيراً للوقت والتكلفة.
    """
    main_text = full_text
    answer_key = []
    vocab_words = []

    rest = ""
    if "### ANSWER_KEY_JSON ###" in full_text:
        main_text, rest = full_text.split("### ANSWER_KEY_JSON ###", 1)

    json_part, vocab_part = rest, ""
    if "### KEY_VOCAB ###" in rest:
        json_part, vocab_part = rest.split("### KEY_VOCAB ###", 1)

    try:
        json_part_clean = json_part.strip()
        # إزالة أي code fence من نوع ```json أو ``` بغض النظر عن مكانها
        json_part_clean = json_part_clean.replace("```json", "").replace("```JSON", "").replace("```", "")
        json_part_clean = json_part_clean.strip("` \n\t")
        if json_part_clean:
            try:
                answer_key = json.loads(json_part_clean)
            except Exception:
                # محاولة أخيرة: استخراج أول قائمة [ ... ] صالحة داخل النص حتى لو
                # كان هناك كلام إضافي قبلها أو بعدها لم يلتزم به النموذج بدقة
                start_idx = json_part_clean.find("[")
                end_idx = json_part_clean.rfind("]")
                if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
                    answer_key = json.loads(json_part_clean[start_idx:end_idx + 1])
            if not isinstance(answer_key, list):
                answer_key = []
    except Exception:
        answer_key = []

    if vocab_part.strip():
        vocab_words = [w.strip() for w in vocab_part.strip().split(",") if w.strip()][:6]

    return main_text.strip(), answer_key, vocab_words


def _format_exam_text_for_display(raw_exam_text):
    """
    يمرّ على نص الامتحان سطراً سطراً؛ أي سطر يبدأ بعبارة "الإجابة النموذجية:"
    يُحوَّل إلى Blockquote مميّز بصرياً (> ✅ **الإجابة النموذجية:** ...) بدل أن
    يبقى سطراً عادياً مطابقاً لبقية النص. بقية الأسطر (الأسئلة والعنوان) تبقى
    كما هي دون أي تعديل حتى لا نفسد تنسيق الذكاء الاصطناعي الأصلي.
    """
    formatted_lines = []
    for raw_line in raw_exam_text.split("\n"):
        line = raw_line.strip()
        if line.startswith("الإجابة النموذجية:"):
            answer_part = line[len("الإجابة النموذجية:"):].strip()
            formatted_lines.append(f"> ✅ **الإجابة النموذجية:** {answer_part}")
        else:
            formatted_lines.append(raw_line)
    return "\n".join(formatted_lines)


def get_students(teacher_name):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(_q("SELECT * FROM students WHERE teacher_name = ? ORDER BY full_name"), (teacher_name,))
    rows = _rows_to_dicts(cur, cur.fetchall())
    release_db(conn)
    return [
        {
            "id": r["id"], "full_name": r["full_name"], "grade": r["grade"],
            "system": r["system"], "category": r["category"], "condition": r["condition_text"]
        }
        for r in rows
    ]


def save_student(teacher_name, full_name, grade, system, category, condition_text):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        _q("INSERT INTO students (teacher_name, full_name, grade, system, category, condition_text, created_at) "
           "VALUES (?, ?, ?, ?, ?, ?, ?)"),
        (teacher_name, full_name, grade, system, category, condition_text, datetime.now().isoformat())
    )
    conn.commit()
    release_db(conn)


def update_student(student_id, grade, system, category, condition_text):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        _q("UPDATE students SET grade=?, system=?, category=?, condition_text=? WHERE id=?"),
        (grade, system, category, condition_text, student_id)
    )
    conn.commit()
    release_db(conn)


def save_worksheet_history(teacher_name, student_id, student_name, subject, grade, level, mode,
                            adapted_text, answer_key_json):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        _q("""INSERT INTO worksheet_history
           (teacher_name, student_id, student_name, subject, grade, adaptation_level, mode,
            adapted_text, answer_key_json, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""),
        (teacher_name, student_id, student_name, subject, grade, level, mode,
         adapted_text, answer_key_json, datetime.now().isoformat())
    )
    conn.commit()
    release_db(conn)


def get_monthly_usage_count(teacher_name):
    """
    يحسب عدد أوراق العمل التي ولّدها هذا المعلم خلال الشهر الحالي (تقويمياً)، لتطبيق
    سقف الاستخدام الشهري (MONTHLY_WORKSHEET_LIMIT) المرتبط بالاشتراك الشهري على Whop.
    """
    conn = get_db()
    cur = conn.cursor()
    month_prefix = datetime.now().strftime("%Y-%m")
    cur.execute(
        _q("SELECT COUNT(*) AS cnt FROM worksheet_history WHERE teacher_name = ? AND created_at LIKE ?"),
        (teacher_name, f"{month_prefix}%")
    )
    row = _row_to_dict(cur, cur.fetchone())
    release_db(conn)
    return row["cnt"] if row and row.get("cnt") is not None else 0


def get_student_history(teacher_name, student_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        _q("SELECT * FROM worksheet_history WHERE teacher_name = ? AND student_id = ? ORDER BY created_at DESC"),
        (teacher_name, student_id)
    )
    rows = _rows_to_dicts(cur, cur.fetchall())
    release_db(conn)
    return rows


if __name__ == "__main__":
    # إعداد صفحة ستريمليت مع العنوان الرسمي الأنيق والأيقونة
    st.set_page_config(
        page_title="نظام تكييف أوراق العمل بالذكاء الاصطناعي | Educational Worksheet Adaptation Platform",
        page_icon="https://cdn.jsdelivr.net/gh/riyadziadeh/-Educational-adapt-platform@main/store_icon.png",
        layout="centered"
    )

    # حقن الشعار الصحيح ضمن meta tags حتى تظهر الصورة الصحيحة عند مشاركة الرابط (واتساب/تيليجرام/إلخ)
    components.html("""
    <script>
    (function() {
      const setMeta = (property, content) => {
        let el = window.parent.document.querySelector(`meta[property="${property}"]`);
        if (!el) {
          el = window.parent.document.createElement('meta');
          el.setAttribute('property', property);
          window.parent.document.head.appendChild(el);
        }
        el.setAttribute('content', content);
      };
      setMeta('og:title', 'نظام تكييف أوراق العمل التربوية');
      setMeta('og:description', 'قم برفع ملف ورقة العمل وسيتم تحليلها وتكييفها تلقائياً باللغة المختارة');
      setMeta('og:image', 'https://cdn.jsdelivr.net/gh/riyadziadeh/-Educational-adapt-platform@main/store_icon.png');
    })();
    </script>
    """, height=0, width=0)

    # =========================================================================================
    # === إصلاح مربّعات "None" الغامضة التي تظهر أعلى الصفحة: المحاولة السابقة (إخفاء
    # stAppToolbar/stToolbar/stDecoration بـ CSS فقط) لم تحل المشكلة بشكل كامل — يبدو أن
    # هذه العناصر تُنشأ أو تُعاد إنشاؤها ديناميكياً من طرف Streamlit نفسه بعد تحميل
    # الصفحة، فقاعدة CSS الثابتة لا تلحقها دائماً. الحل الأشمل هنا: سكربت JS يبحث بشكل
    # دوري (ومع أي تغيير بالصفحة عبر MutationObserver) عن أي عنصر "ورقة" (بلا عناصر
    # فرعية) نصّه الكامل هو "None" بالضبط، ويخفيه فوراً — بغضّ النظر عن مصدره الحقيقي.
    # هذا آمن لأن واجهة التطبيق كلها بالعربية ولا يوجد أي نص حقيقي بالتطبيق يساوي "None". ===
    # =========================================================================================
    components.html("""
    <script>
    (function() {
      function hideNoneArtifacts() {
        try {
          const doc = window.parent.document;
          const candidates = doc.querySelectorAll('div, span, p, small, label');
          candidates.forEach((el) => {
            if (el.children.length === 0 && el.textContent.trim() === 'None') {
              el.style.display = 'none';
            }
          });
        } catch (e) { /* تجاهل أي خطأ صلاحيات وصول بين الإطارات */ }
      }
      hideNoneArtifacts();
      try {
        const observer = new MutationObserver(hideNoneArtifacts);
        observer.observe(window.parent.document.body, { childList: true, subtree: true });
      } catch (e) { /* تجاهل */ }
      setInterval(hideNoneArtifacts, 1000);
    })();
    </script>
    """, height=0, width=0)

    # =========================================================================================
    # === هوية الألوان الرسمية للتطبيق: كحلي غامق + كحلي/أزرق فاتح + ذهبي + أبيض ===
    # =========================================================================================
    NAVY_DARK = "#101B2D"      # الأزرق الكحلي الغامق (خلفية الهيدر والبطاقات المختارة)
    NAVY_LIGHT = "#EAF1FB"     # الأزرق الفاتح جداً (خلفية الشبكة العامة)
    BLUE_ACCENT = "#2E6FBB"    # أزرق متوسط لدعم التدرجات والحدود
    GOLD = "#F1C40F"           # الذهبي (لون التمييز والعناصر النشطة)
    WHITE = "#FFFFFF"

    st.markdown(textwrap.dedent(f"""
        <style>
        /* استيراد خط Cairo العصري (يدعم العربية بشكل ممتاز) لطابع بصري أحدث لعام 2026 */
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap');

        html, body, [class*="css"], .stApp, .stMarkdown, p, h1, h2, h3, h4, label, button {{
            font-family: 'Cairo', -apple-system, sans-serif !important;
        }}

        /* === إصلاح تجريبي لمربّعي "None" الغامضين أعلى الصفحة: كانت هذه القاعدة تُخفي
        فقط زر أيقونة GitHub (stToolbarActions) داخل شريط Streamlit العلوي، لكن قد يترك
        هذا عناصر نصية أخرى تابعة لنفس الشريط (مثل عدّاد أو تسمية فشل تحميلها) ظاهرة
        بمفردها بدون الأيقونة المرتبطة بها. الآن نُخفي الشريط العلوي الافتراضي بالكامل
        (stAppToolbar / stToolbar / stDecoration) بدل جزء منه فقط — لا يؤثر على أيقونة
        فتح القائمة الجانبية المخصّصة (stSidebarCollapsedControl) لأنها عنصر منفصل تماماً
        ومُنسَّقة بقاعدة أخرى أدناه. ===== */
        .stAppToolbar,
        [data-testid="stToolbar"],
        [data-testid="stDecoration"] {{
            display: none !important;
            visibility: hidden !important;
        }}

        /* ===== تكبير وتمييز أيقونة فتح القائمة الجانبية (حساب المعلم) لتصبح واضحة
        كزر حقيقي بدل سهم ">" صغير قد يمر دون ملاحظة، مع إضافة نص "تسجيل الدخول وإدارة
        الطلاب" بجانبها مباشرة. === إصلاح: الاعتماد السابق على سكربت JS (components.html)
        لإضافة هذا النص لم يكن يعمل بثبات لأن الإطار (iframe) الخاص بالمكوّن قد لا يملك
        وصولاً مضموناً لعنصر الصفحة الأصلية حسب بيئة الاستضافة. الحل الأكثر ثباتاً هو
        إضافة النص كـ CSS Pseudo-element (::after) مباشرة داخل نفس زر ">>"، فيظهر دائماً
        بمجرد تحميل الصفحة دون أي اعتماد على تنفيذ سكربت خارجي. ===== */
        [data-testid="stSidebarCollapsedControl"] {{
            background: linear-gradient(120deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%) !important;
            border-radius: 14px !important;
            padding: 8px 14px !important;
            box-shadow: 0px 4px 16px rgba(16,27,45,0.40) !important;
            display: flex !important;
            flex-direction: row-reverse !important;
            align-items: center !important;
            gap: 8px !important;
            width: auto !important;
            min-width: unset !important;
        }}
        [data-testid="stSidebarCollapsedControl"] svg {{
            width: 24px !important;
            height: 24px !important;
            color: {GOLD} !important;
            fill: {GOLD} !important;
            flex-shrink: 0 !important;
        }}
        [data-testid="stSidebarCollapsedControl"]::after {{
            content: "تسجيل الدخول وإدارة الطلاب";
            color: {GOLD} !important;
            font-weight: 800 !important;
            font-size: 12.5px !important;
            white-space: nowrap !important;
            font-family: 'Cairo', -apple-system, sans-serif !important;
            direction: rtl !important;
        }}
        [data-testid="stSidebarNavCollapseIcon"], button[kind="header"] {{
            color: {GOLD} !important;
        }}

        /* خلفية متدرجة ناعمة حديثة بدل اللون الفلات القديم، مع طبقة زجاجية خفيفة */
        .stApp {{
            background: linear-gradient(160deg, #F4F8FF 0%, {NAVY_LIGHT} 45%, #E4ECFB 100%);
            background-attachment: fixed;
        }}

        /* تأثير الحركة الانسيابية (Animation) لصندوق الشكر */
        @keyframes fadeInScale {{
            0% {{ opacity: 0; transform: scale(0.95); }}
            100% {{ opacity: 1; transform: scale(1); }}
        }}
        .animated-box {{
            animation: fadeInScale 0.8s ease-in-out;
        }}

        /* شريط علوي كحلي غامق يشبه هيدر التطبيقات الحديثة، بحواف أدور وظل أنعم */
        .app-topbar {{
            background: linear-gradient(120deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%);
            border-radius: 26px;
            padding: 18px 22px;
            margin-bottom: 24px;
            box-shadow: 0px 10px 30px rgba(16,27,45,0.22);
        }}
        .app-topbar .search-fake {{
            background-color: rgba(255,255,255,0.92);
            backdrop-filter: blur(6px);
            border-radius: 16px;
            padding: 12px 18px;
            color: #7d8a9a;
            font-weight: 700;
            text-align: right;
            font-size: 15px;
        }}

        /* عناوين الحقول والقوائم بخط عريض وواضح جداً */
        .stSelectbox label p, .stFileUploader label p, div[data-baseweb="select"] label, label, .stCheckbox label p {{
            font-weight: 900 !important;
            font-size: 17px !important;
            color: {NAVY_DARK} !important;
        }}

        /* صناديق الاختيار (Selectbox) و رافع الملفات بحواف أدور وحدود ذهبية ناعمة */
        div[data-baseweb="select"] > div, div.stFileUploader > div {{
            background-color: rgba(255,255,255,0.85) !important;
            backdrop-filter: blur(8px);
            border-radius: 18px !important;
            border: 2px solid {GOLD}99 !important;
        }}
        div[data-baseweb="select"] > div * {{
            color: {NAVY_DARK} !important;
        }}

        /* بطاقات الشبكة بطابع زجاجي عصري (Glassmorphism) */
        div[data-testid="stButton"] > button {{
            width: 100%;
            min-height: 108px;
            height: auto !important;
            padding: 16px 12px !important;
            border-radius: 22px !important;
            border: 1.5px solid rgba(46,111,187,0.18) !important;
            background: rgba(255,255,255,0.75) !important;
            backdrop-filter: blur(10px);
            color: {NAVY_DARK} !important;
            font-weight: 800 !important;
            font-size: 16px !important;
            line-height: 1.6 !important;
            box-shadow: 0px 6px 18px rgba(16,27,45,0.10);
            transition: all 0.22s cubic-bezier(0.22, 1, 0.36, 1);
            white-space: pre-line !important;
        }}
        div[data-testid="stButton"] > button:hover {{
            border: 1.5px solid {GOLD} !important;
            transform: translateY(-3px) scale(1.01);
            box-shadow: 0px 10px 24px rgba(16,27,45,0.16);
        }}

        /* ===== زر "ابدأ تكييف ورقة العمل" الرئيسي: عرض كامل من الطرف للطرف (وليس مربعاً صغيراً) ===== */
        div[class*="st-key-start-ai-button"] {{
            width: 100% !important;
            display: block !important;
        }}
        div[class*="st-key-start-ai-button"] div[data-testid="stButton"] {{
            width: 100% !important;
            display: block !important;
        }}
        div[class*="st-key-start-ai-button"] div[data-testid="stButton"] > button {{
            width: 100% !important;
            display: block !important;
            background: linear-gradient(120deg, {GOLD} 0%, #FFD84D 100%) !important;
            color: {NAVY_DARK} !important;
            font-weight: 900 !important;
            font-size: 19px !important;
            line-height: 1.6 !important;
            height: auto !important;
            min-height: 90px !important;
            padding: 18px 14px !important;
            border-radius: 28px !important;
            border: none !important;
            box-shadow: 0px 12px 28px rgba(241,196,15,0.35);
            white-space: normal !important;
            overflow: visible !important;
            word-wrap: break-word !important;
            transition: all 0.25s cubic-bezier(0.22, 1, 0.36, 1) !important;
        }}
        div[class*="st-key-start-ai-button"] div[data-testid="stButton"] > button:hover {{
            transform: translateY(-3px) scale(1.01);
            box-shadow: 0px 16px 34px rgba(241,196,15,0.45);
            color: {NAVY_DARK} !important;
        }}
        div[class*="st-key-start-ai-button"] div[data-testid="stButton"] > button p {{
            color: {NAVY_DARK} !important;
            font-weight: 900 !important;
            font-size: 19px !important;
            white-space: normal !important;
            overflow: visible !important;
            word-wrap: break-word !important;
            margin: 0 !important;
        }}

        .grid-title {{
            font-weight: 900;
            font-size: 19px;
            color: {NAVY_DARK};
            text-align: right;
            margin: 6px 0 2px 0;
        }}

        /* إجبار أعمدة ستريمليت على البقاء بجانب بعضها أفقياً حتى على شاشات الجوال الضيقة */
        div[data-testid="stHorizontalBlock"] {{
            flex-direction: row !important;
            flex-wrap: wrap !important;
            align-items: flex-start !important;
        }}
        div[data-testid="stColumn"] {{
            min-width: 0 !important;
        }}

        /* ===== تحويل أزرار المواد الدراسية إلى دوائر حقيقية بحجم مناسب للأيقونة (٤ بجانب بعض) =====
           يتم الاستهداف عبر كلاس الحاوية (st-key-circlebtn-...) بدل الاعتماد على ترتيب العناصر،
           لضمان عمل التصميم بشكل موثوق بغض النظر عن التغييرات الداخلية بهيكل DOM في ستريمليت. */
        div[class*="st-key-circlebtn-"] {{
            display: flex !important;
            justify-content: center !important;
        }}
        div[class*="st-key-circlebtn-"] div[data-testid="stButton"] {{
            display: flex !important;
            justify-content: center !important;
            width: 100% !important;
        }}
        div[class*="st-key-circlebtn-"] button {{
            box-sizing: border-box !important;
            width: 68px !important;
            height: 68px !important;
            min-width: 68px !important;
            min-height: 68px !important;
            max-width: 68px !important;
            max-height: 68px !important;
            aspect-ratio: 1 / 1 !important;
            flex: none !important;
            border-radius: 50% !important;
            background: rgba(255,255,255,0.65) !important;
            backdrop-filter: blur(8px);
            border: 2px solid rgba(46,111,187,0.25) !important;
            font-size: 22px !important;
            line-height: 1 !important;
            padding: 0 !important;
            margin: 0 auto !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            box-shadow: 0px 6px 16px rgba(16,27,45,0.10);
            transition: all 0.22s cubic-bezier(0.22, 1, 0.36, 1) !important;
        }}
        div[class*="st-key-circlebtn-"] button p {{
            margin: 0 !important;
            line-height: 1 !important;
        }}
        div[class*="st-key-circlebtn-"] button:hover {{
            border: 2px solid {GOLD} !important;
            transform: scale(1.1);
            box-shadow: 0px 10px 22px rgba(16,27,45,0.16);
        }}
        div[class*="st-key-circlebtn-"][class*="-on"] button {{
            border: 2px solid {GOLD} !important;
            background: linear-gradient(150deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%) !important;
            box-shadow: 0px 8px 20px rgba(46,111,187,0.35);
        }}
        .circle-caption {{
            text-align: center;
            font-weight: 800;
            font-size: 11.5px;
            color: {NAVY_DARK};
            margin: 4px auto 14px auto;
            line-height: 1.3;
            max-width: 92px;
        }}

        /* ===== صندوق زجاجي واحد يلف كامل شبكة دوائر المواد الدراسية (Glassmorphism) ===== */
        div[class*="st-key-subjects-white-card"] {{
            background: rgba(255,255,255,0.55) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid rgba(255,255,255,0.6);
            border-radius: 28px !important;
            padding: 20px 16px 10px 16px !important;
            margin-bottom: 20px !important;
            box-shadow: 0px 10px 28px rgba(16,27,45,0.10);
        }}

        /* ===== صندوق قسم إدارة الطلاب (زجاجي أيضاً) ===== */
        div[class*="st-key-student-mgmt-card"] {{
            background: rgba(255,255,255,0.55) !important;
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border: 1px solid rgba(255,255,255,0.6);
            border-radius: 24px !important;
            padding: 16px !important;
            margin-bottom: 18px !important;
            box-shadow: 0px 8px 22px rgba(16,27,45,0.08);
        }}

        /* ===== بطاقة "تحويل إلى امتحان تقييمي": بارزة بحدود ذهبية لتُلفت الانتباه كخطوة
        اختيارية مهمة بعد التكييف مباشرة ===== */
        div[class*="st-key-exam-conversion-card"] {{
            background: rgba(255,255,255,0.7) !important;
            backdrop-filter: blur(12px);
            border: 2px solid {GOLD}aa;
            border-radius: 22px !important;
            padding: 18px 20px !important;
            margin: 14px 0 20px 0 !important;
            box-shadow: 0px 8px 22px rgba(241,196,15,0.18);
        }}
        div[class*="st-key-exam-conversion-card"] button {{
            background: linear-gradient(120deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%) !important;
            color: {WHITE} !important;
            font-weight: 800 !important;
            border: none !important;
        }}

        /* ===== بطاقة عرض ورقة العمل المكيّفة: تباعد أسطر مريح + تمييز واضح للعناوين
        والأسئلة الغامقة والفواصل بين التمارين، بدل نص متلاصق متعب للقراءة ===== */
        div[class*="st-key-worksheet-output-card"] {{
            background: rgba(255,255,255,0.78) !important;
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(46,111,187,0.16);
            border-radius: 24px !important;
            padding: 30px 28px !important;
            margin: 10px 0 24px 0 !important;
            box-shadow: 0px 10px 26px rgba(16,27,45,0.08);
        }}
        div[class*="st-key-worksheet-output-card"] p {{
            line-height: 2.1 !important;
            font-size: 17px !important;
            color: {NAVY_DARK};
            margin-bottom: 14px !important;
            direction: rtl !important;
            unicode-bidi: plaintext !important;
            text-align: right !important;
        }}
        /* === إصلاح: كانت خلفية ملوّنة (Highlight Box) توضع خلف كل نص **عريض**، وعندما
        يكون هذا النص طويلاً ويلتف على أكثر من سطر (خصوصاً مع خلط عربي/أرقام/رموز
        رياضية مثل a= 2 وb= -7)، كانت الخلفية تتقطّع بشكل غير متناسق مع التفاف السطر
        فيظهر الشكل "ملخبطاً" وكأن ترتيب الكلمات تبدّل. الحل: إبقاء النص عريضاً وبلون
        مميّز بدون أي صندوق خلفية يلتف معه، وهذا وحده يحل معظم المشكلة البصرية. ===== */
        div[class*="st-key-worksheet-output-card"] strong {{
            color: {BLUE_ACCENT} !important;
            font-weight: 900 !important;
            background: none !important;
            padding: 0 !important;
            border-radius: 0 !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-worksheet-output-card"] em {{
            color: {BLUE_ACCENT} !important;
        }}
        div[class*="st-key-worksheet-output-card"] h1,
        div[class*="st-key-worksheet-output-card"] h2,
        div[class*="st-key-worksheet-output-card"] h3 {{
            color: {BLUE_ACCENT} !important;
            text-align: right;
            margin-top: 10px !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-worksheet-output-card"] hr {{
            border: none;
            border-top: 2px dashed rgba(46,111,187,0.35);
            margin: 24px 0 !important;
        }}
        div[class*="st-key-worksheet-output-card"] ul,
        div[class*="st-key-worksheet-output-card"] ol {{
            line-height: 2 !important;
            font-size: 16.5px !important;
            direction: rtl !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-worksheet-output-card"] li {{
            unicode-bidi: plaintext !important;
        }}

        /* ===== إصلاح (٢): تنسيق مخصص ومرتّب لبطاقة "الامتحان التقييمي" الناتج بعد
        التكييف — هوية بصرية خاصة (أخضر/تيل) تميّزه عن ورقة العمل الأصلية، مع تباعد
        أسطر مريح للأسئلة، وإبراز واضح لكل "إجابة نموذجية" داخل صندوق مستقل ملوّن
        (عبر تحويلها إلى Blockquote في المعالجة النصية أدناه) بدل أن تظهر مطابقة
        لشكل نص السؤال العادي. ===== */
        div[class*="st-key-exam-output-card"] {{
            background: rgba(255,255,255,0.82) !important;
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 2px solid rgba(46,187,109,0.35);
            border-radius: 24px !important;
            padding: 28px 26px !important;
            margin: 10px 0 22px 0 !important;
            box-shadow: 0px 10px 26px rgba(16,27,45,0.08);
        }}
        div[class*="st-key-exam-output-card"] p {{
            line-height: 2.1 !important;
            font-size: 16.5px !important;
            color: {NAVY_DARK};
            margin-bottom: 12px !important;
            direction: rtl !important;
            unicode-bidi: plaintext !important;
            text-align: right !important;
        }}
        div[class*="st-key-exam-output-card"] strong {{
            color: #1F9D63 !important;
            font-weight: 900 !important;
            background: none !important;
            padding: 0 !important;
            border-radius: 0 !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-exam-output-card"] em {{
            color: #1F9D63 !important;
        }}
        div[class*="st-key-exam-output-card"] h1,
        div[class*="st-key-exam-output-card"] h2,
        div[class*="st-key-exam-output-card"] h3 {{
            color: #1F9D63 !important;
            text-align: right;
            margin-top: 8px !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-exam-output-card"] hr {{
            border: none;
            border-top: 2px dashed rgba(46,187,109,0.35);
            margin: 22px 0 !important;
        }}
        div[class*="st-key-exam-output-card"] ul,
        div[class*="st-key-exam-output-card"] ol {{
            line-height: 2 !important;
            font-size: 16px !important;
            direction: rtl !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-exam-output-card"] li {{
            unicode-bidi: plaintext !important;
        }}
        /* صندوق "الإجابة النموذجية" المميّز — يأتي من تحويل السطر إلى Blockquote (>)
        في دالة _format_exam_text_for_display حتى يبرز بصرياً عن نص السؤال */
        div[class*="st-key-exam-output-card"] blockquote {{
            background: rgba(46,187,109,0.12) !important;
            border-right: 4px solid #1F9D63 !important;
            border-left: none !important;
            margin: 6px 0 20px 0 !important;
            padding: 10px 16px !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            color: #0F6B41 !important;
            direction: rtl !important;
            unicode-bidi: plaintext !important;
        }}
        div[class*="st-key-exam-output-card"] blockquote p {{
            margin-bottom: 0 !important;
            color: #0F6B41 !important;
        }}

        .selection-summary {{
            background: linear-gradient(120deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%);
            color: {GOLD};
            border-radius: 20px;
            padding: 12px 18px;
            text-align: center;
            font-weight: 800;
            margin: 12px 0 20px 0;
            box-shadow: 0px 8px 20px rgba(16,27,45,0.18);
        }}

        /* ===== لمسات إضافية واضحة لطابع 2026: كرات ضوئية متوهجة خلف المحتوى + عنوان
        متدرج اللون + حركة ظهور تدريجي للصفحة كاملة، حتى يبين التحديث فوراً من أول ثانية ===== */
        .stApp::before, .stApp::after {{
            content: "";
            position: fixed;
            border-radius: 50%;
            filter: blur(70px);
            z-index: 0;
            pointer-events: none;
            opacity: 0.35;
        }}
        .stApp::before {{
            width: 280px;
            height: 280px;
            top: -80px;
            right: -60px;
            background: radial-gradient(circle, {GOLD} 0%, transparent 70%);
        }}
        .stApp::after {{
            width: 320px;
            height: 320px;
            bottom: -100px;
            left: -80px;
            background: radial-gradient(circle, {BLUE_ACCENT} 0%, transparent 70%);
        }}
        .stApp > div {{
            position: relative;
            z-index: 1;
        }}

        @keyframes pageFadeIn {{
            0% {{ opacity: 0; transform: translateY(10px); }}
            100% {{ opacity: 1; transform: translateY(0); }}
        }}
        .main .block-container {{
            animation: pageFadeIn 0.5s ease-out;
        }}

        /* عنوان النظام الرئيسي بأسلوب متدرج اللون (Gradient Text) بدل اللون الفلات */
        h1 {{
            background: linear-gradient(90deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 60%, {GOLD} 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            letter-spacing: 0.5px;
        }}
        h2 {{
            color: {BLUE_ACCENT} !important;
            letter-spacing: 0.3px;
        }}
        </style>
    """).strip(), unsafe_allow_html=True)

    # =========================================================================================
    # === تسجيل دخول مبسّط للمعلم عبر الشريط الجانبي، لفصل بيانات كل معلم (طلابه وسجل
    # أوراقه) عن غيره في نفس قاعدة البيانات المشتركة — يناسب استخدام التطبيق من عدة
    # معلمين في نفس المدرسة دون الحاجة لنظام حسابات معقّد. ===
    # =========================================================================================
    with st.sidebar:
        st.markdown("### 👩‍🏫 حساب المعلم / Teacher Account")

        if USE_POSTGRES and not DB_INIT_ERROR:
            st.caption("🟢 التخزين دائم (متصل بقاعدة بيانات Supabase) — بياناتك لن تُفقد عند إعادة النشر.")
        elif DB_INIT_ERROR:
            # === إصلاح أمني: لا نعرض نص خطأ الاتصال الخام (قد يكشف اسم مضيف/بيانات
            # خادم قاعدة البيانات) لأي زائر — حتى قبل تسجيل الدخول. رسالة عامة فقط هنا،
            # والتفاصيل الكاملة تُطبع بسجلات الخادم (راجع log_internal_error). ===
            st.caption("🔴 تعذّر الاتصال بقاعدة البيانات الدائمة حالياً.")
            st.caption("سيتم استخدام تخزين مؤقت محلياً حتى يُحل الاتصال.")
        else:
            st.caption("🟡 التخزين مؤقت حالياً (لم يُضبط SUPABASE_DB_URL بعد) — البيانات قد تُفقد عند إعادة النشر.")
        if "teacher_name" not in st.session_state:
            st.session_state.teacher_name = ""

        if st.session_state.teacher_name:
            st.success(f"مرحباً {st.session_state.teacher_name} 👋")
            if st.button("🚪 تسجيل خروج / Logout", key="logout_btn"):
                st.session_state.teacher_name = ""
                st.rerun()
        else:
            # === إصلاح: نموذج تسجيل الدخول/إنشاء الحساب انتقل من الشريط الجانبي إلى
            # منتصف الصفحة الرئيسية مباشرة (أنظر أسفل قسم "يجب تسجيل الدخول أولاً").
            # السبب: على شاشات الموبايل الشريط الجانبي يكون مطوياً تلقائياً وبدون زر
            # واضح لفتحه في بعض المتصفحات، فكان المستخدم يرى فقط رسالة "يجب تسجيل
            # الدخول" بدون أي طريقة فعلية للوصول لنموذج الدخول. الآن النموذج نفسه
            # جزء من الصفحة الرئيسية، ظاهر مباشرة بدون الحاجة لفتح أي قائمة جانبية. ===
            st.caption("👈 نموذج تسجيل الدخول / إنشاء الحساب موجود في وسط الصفحة الرئيسية.")

    # =========================================================================================
    # === إزالة الخلفية البيضاء من صورة الشعار تلقائياً وتحويلها لشفافة، حتى يندمج الشعار
    # بصرياً مع خلفية التطبيق بدل الظهور داخل مربع أبيض واضح الحواف. النتيجة مخزّنة
    # مؤقتاً (cache) حتى لا تُعاد المعالجة في كل rerun. ===
    # === إصلاح: نقلنا عرض الشعار ليكون قبل فحص تسجيل الدخول (بدل بعده) — بهذا الشكل
    # الشعار يظهر أيضاً لأي زائر غير مسجّل على صفحة "يجب تسجيل الدخول أولاً"، لا يظهر
    # فقط بعد الدخول. ===
    # =========================================================================================
    @st.cache_data(show_spinner=False)
    def _load_logo_with_transparent_background(path, white_threshold=245):
        image = Image.open(path).convert("RGBA")
        datas = image.getdata()
        new_data = []
        for item in datas:
            if item[0] >= white_threshold and item[1] >= white_threshold and item[2] >= white_threshold:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        image.putdata(new_data)
        return image

    def _render_logo():
        """يعرض شعار التطبيق في منتصف الصفحة — يُستخدم قبل تسجيل الدخول وبعده."""
        col_logo1, col_logo2, col_logo3 = st.columns([0.5, 3, 0.5])
        with col_logo2:
            logo_loaded = False
            logo_filenames = [
                "new_logo.png",
                "Educ_Worksheet_Adapt_Icon_(Square).png",
                "logo.png", "logo.jpg", "Logo.png", "Logo.JPG"
            ]
            for filename in logo_filenames:
                if os.path.exists(filename):
                    try:
                        image = _load_logo_with_transparent_background(filename)
                    except Exception:
                        image = Image.open(filename)
                    st.image(image, use_container_width=True)
                    logo_loaded = True
                    break
            if not logo_loaded:
                st.warning("الرجاء التأكد من رفع صورة الأيقونة باسم new_logo.png في نفس مجلد المشروع.")

    _render_logo()

    # =========================================================================================
    # === إصلاح أمني مهم: قبل هذا الإصلاح كان بإمكان أي زائر استخدام التطبيق (رفع ورقة
    # عمل والحصول على تكييف كامل) دون تسجيل دخول على الإطلاق — الكود كان يستخدم اسم
    # "معلم_عام" كحساب افتراضي لأي زائر غير مسجّل. الآن: أي زائر غير مسجّل يرى نموذج
    # تسجيل دخول/إنشاء حساب واضح في وسط الصفحة الرئيسية مباشرة (وليس في الشريط الجانبي
    # فقط — لأن الشريط الجانبي على الموبايل يكون مطوياً بدون طريقة واضحة لفتحه)، ولا
    # يصل لأي جزء من واجهة التكييف الفعلية. st.stop() يوقف تنفيذ باقي الصفحة بعد عرض
    # هذا النموذج مباشرة. ===
    # =========================================================================================
    if not st.session_state.teacher_name:
        st.markdown(textwrap.dedent("""
            <div style="text-align:center; max-width:620px; margin:0 auto; padding: 30px 20px 6px 20px;
                        font-family:'Cairo', -apple-system, sans-serif;">
                <div style="font-size:48px; line-height:1; margin-bottom:10px;">🔒</div>

                <h2 style="margin:0 0 4px 0; color:#101B2D; font-weight:800;">يجب تسجيل الدخول أولاً</h2>
                <h3 style="margin:0 0 22px 0; color:#2E6FBB; font-weight:600;">You must log in first</h3>

                <p style="font-size:17px; color:#222; line-height:2; margin:0;">
                    سجّل دخولك بحسابك الحالي<br>
                    أو أنشئ حساباً جديداً بمفتاح ترخيص Whop<br>
                    من النموذج أدناه للاستمرار.
                </p>

                <div style="height:1px; background:#E3E9F2; margin:20px auto; width:80%;"></div>

                <p style="font-size:15px; color:#666; line-height:2; margin:0;">
                    Log in with your existing account,<br>
                    or create a new one using your Whop license key,<br>
                    from the form below to continue.
                </p>
            </div>
        """, unsafe_allow_html=True)

        _login_col_l, _login_col_mid, _login_col_r = st.columns([1, 2.4, 1])
        with _login_col_mid:
            auth_mode = st.radio(
                "اختر / Choose:",
                ["🔐 تسجيل دخول / Login", "🆕 إنشاء حساب جديد / Create Account"],
                key="auth_mode_radio",
                horizontal=True,
            )

            st.write("")  # مسافة فاصلة صغيرة أنيقة بين الاختيار والنموذج

            if auth_mode == "🔐 تسجيل دخول / Login":
                st.markdown("#### 🔐 تسجيل الدخول")
                st.caption("أدخل اسم المستخدم وكلمة المرور اللي سجّلت فيهم حسابك.")
                # === إصلاح: لفّ الحقول وزر الدخول داخل st.form. سابقاً كانت الحقول
                # الثلاثة والزر عناصر منفصلة خارج أي form — كتابة أي حرف بحقل كلمة
                # المرور تُطلق rerun فوري لكامل الصفحة، وعلى بعض المتصفحات (خصوصاً
                # الجوال عند ظهور لوحة المفاتيح الافتراضية وتغيّر حجم الشاشة أثناء ذلك)
                # كان هذا يتسبب أحياناً بعدم ظهور/اختفاء الزر لحظة الكتابة لأن الصفحة
                # تُعاد رسمتها بالكامل في كل ضغطة. الآن مع st.form: الحقول لا تُطلق أي
                # rerun أثناء الكتابة إطلاقاً، ويبقى الزر ثابتاً دائماً، ولا يحدث أي
                # تنفيذ إلا عند الضغط الفعلي على زر الإرسال. ===
                with st.form("login_form", clear_on_submit=False):
                    login_username = st.text_input("اسم المستخدم / Username:", key="login_username_input")
                    login_password = st.text_input(
                        "كلمة المرور / Password:",
                        key="login_password_input", type="password", max_chars=32
                    )
                    st.write("")
                    login_submitted = st.form_submit_button("🔐 دخول / Login", use_container_width=True)
                if login_submitted:
                    success, message = login_teacher(login_username, login_password)
                    if success:
                        st.session_state.teacher_name = login_username.strip()
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)

            else:
                st.markdown("#### 🆕 إنشاء حساب جديد")
                st.caption(f"اختر اسم مستخدم جديد وكلمة مرور. {PASSWORD_REQUIREMENT_MSG}")
                st.info(
                    "🔑 لازم يكون معك مفتاح ترخيص (License Key) استلمته بعد الاشتراك في EWAS على Whop "
                    "— بتوصلك بإيميل التأكيد بعد الدفع، أو من قسم \"Software\" بحسابك على Whop.",
                    icon="🔑",
                )
                st.write("")
                # === نفس إصلاح st.form أعلاه، مطبّق هنا أيضاً على نموذج إنشاء الحساب. ===
                with st.form("register_form", clear_on_submit=False):
                    register_username = st.text_input("اسم المستخدم الجديد / New Username:", key="register_username_input")
                    register_password = st.text_input(
                        "كلمة المرور (٦ خانات على الأقل: حروف وأرقام) / Password:",
                        key="register_password_input", type="password", max_chars=32
                    )
                    register_password_confirm = st.text_input(
                        "تأكيد كلمة المرور / Confirm Password:",
                        key="register_password_confirm_input", type="password", max_chars=32
                    )
                    register_license_key = st.text_input(
                        "مفتاح الترخيص (License Key) من Whop:",
                        key="register_license_key_input"
                    )
                    st.write("")
                    register_submitted = st.form_submit_button("🆕 إنشاء الحساب / Create Account", use_container_width=True)
                if register_submitted:
                    if register_password != register_password_confirm:
                        st.error("كلمة المرور وتأكيدها غير متطابقين.")
                    else:
                        with st.spinner("جاري التحقق من مفتاح الترخيص مع Whop..."):
                            success, message = register_teacher(register_username, register_password, register_license_key)
                        if success:
                            st.session_state.teacher_name = register_username.strip()
                            st.success(message)
                            st.rerun()
                        else:
                            st.error(message)

        st.stop()

    # === ملاحظة: إزالة صندوق الشرح الأصفر الكبير أصبحت نهائية — نص "تسجيل الدخول
    # وإدارة الطلاب" أصبح يظهر بجانب سهم ">>" مباشرة عبر CSS (::after) المُعرَّف أعلاه
    # في كتلة <style>، وهو أسلوب أكثر ثباتاً من الاعتماد على سكربت JS خارجي لأنه لا
    # يحتاج أي وصول لعناصر الصفحة الأصلية من داخل إطار iframe منفصل. ===

    # شريط علوي كحلي بأسلوب "شريط البحث" الموجود في التطبيقات، للزينة وربط الهوية البصرية بالتصميم المطلوب
    st.markdown(f"""
        <div class="app-topbar">
            <div class="search-fake">🔍 &nbsp; اختر بيانات ورقة العمل من الشبكة أدناه</div>
        </div>
    """, unsafe_allow_html=True)

    # العنوان الرئيسي للنظام تحت الشعار مباشرة
    st.markdown("""
        <div style="text-align: center;">
            <h1 style="font-size: 28px; margin-bottom: 0; font-weight: 900;">نظام تكييف أوراق العمل التربوية</h1>
            <h2 style="font-size: 22px; margin-top: 5px; font-weight: 900;">Educational Worksheet Adaptation System</h2>
        </div>
    """, unsafe_allow_html=True)

    st.write("قم برفع ملف ورقة العمل وسيتم تحليلها وتكييفها تلقائياً باللغة المختارة مع خيارات التحميل المتعددة.")

    # =========================================================================================
    # --- الموسيقى الخلفية اختيارية بالكامل (Opt-in) بدل التشغيل التلقائي، لأن جزءاً كبيراً
    # من مستخدمي هذا النظام هم طلاب لديهم حساسية حسية (مثل اضطراب طيف التوحد)، وصوت يعمل
    # من تلقاء نفسه قد يكون مزعجاً أو مربكاً لهم. ---
    # =========================================================================================
    audio_file_path = None
    for music_name in ["music.mp3", "Music.mp3", "MUSIC.MP3", "music.WAV", "music.ogg"]:
        if os.path.exists(music_name):
            audio_file_path = music_name
            break

    with st.expander("🎵 إعدادات الصوت (اختياري) / Audio Settings (Optional)"):
        enable_bg_music = st.checkbox("تشغيل موسيقى خلفية هادئة أثناء الاستخدام", value=False)
        enable_ding = st.checkbox("تشغيل تنبيه صوتي قصير عند اكتمال تكييف الورقة", value=True)
        if enable_bg_music:
            if audio_file_path:
                st.audio(audio_file_path, format="audio/mp3", loop=True)
            else:
                st.audio("https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3", format="audio/mp3", loop=True)

    # =========================================================================================
    # === جديد: إعدادات الأداء والسرعة ===
    # توليد صور بالذكاء الاصطناعي (لغلاف العرض التقديمي وبطاقات PECS) كان يستدعي نموذج
    # توليد الصور من جوجل حتى 10 مرات متتالية لكل ورقة عمل — وهذا كان السبب الرئيسي
    # الثاني في بطء التطبيق (بعد اتصال قاعدة البيانات). الآن هذا اختياري ومطفأ افتراضياً؛
    # عند إطفائه تُستخدم أيقونات مرسومة محلياً فوراً (بدون أي اتصال إنترنت إضافي).
    # =========================================================================================
    with st.expander("⚙️ إعدادات الأداء والسرعة / Performance Settings"):
        enable_ai_images = st.checkbox(
            "توليد صور توضيحية بالذكاء الاصطناعي داخل عرض PowerPoint وبطاقات PECS "
            "(شكل أجمل لكنه أبطأ بشكل ملحوظ لأنه يستدعي نموذج توليد الصور عدة مرات "
            "لكل ورقة عمل). اتركه مطفأً للحصول على أسرع أداء ممكن — سيتم استخدام "
            "أيقونات مرسومة محلياً بدلاً منها.",
            value=False,
            key="enable_ai_images_checkbox"
        )

    st.markdown("---")

    # =========================================================================================
    # صوت تنبيه صغير (طنة) عند اكتمال تجهيز ورقة العمل — مضمّن مباشرة بالكود (Base64) بدون إنترنت
    # =========================================================================================
    DING_SOUND_B64 = "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1lIaUrBmER4NSpagUAtMREs9EoiKytVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVX/+1LE7gEMjDULAJuESVKGYVQdJRhVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVf/7UsTRg8vlRvYE4EkAAAA0gAAABFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV"

    def play_ready_ding():
        """تشغيل نغمة تنبيه قصيرة فور جاهزية ورقة العمل (تعمل فقط إذا فعّلها المستخدم)."""
        components.html(f"""
            <audio autoplay>
                <source src="data:audio/mp3;base64,{DING_SOUND_B64}" type="audio/mp3">
            </audio>
        """, height=0, width=0)

    # =========================================================================================
    # جلب مفتاح الـ API بمرونة تامة (سواء من الأسرار أو من متغيرات البيئة)
    # =========================================================================================
    api_key = None
    try:
        api_key = st.secrets["GOOGLE_API_KEY"]
    except Exception:
        api_key = os.getenv("GOOGLE_API_KEY")

    # إعدادات تكامل Canva الاختيارية (Canva Connect API - Autofill API)
    CANVA_API_TOKEN = None
    CANVA_BRAND_TEMPLATE_ID = None
    try:
        CANVA_API_TOKEN = st.secrets.get("CANVA_API_TOKEN", None)
        CANVA_BRAND_TEMPLATE_ID = st.secrets.get("CANVA_BRAND_TEMPLATE_ID", None)
    except Exception:
        CANVA_API_TOKEN = os.getenv("CANVA_API_TOKEN")
        CANVA_BRAND_TEMPLATE_ID = os.getenv("CANVA_BRAND_TEMPLATE_ID")

    CANVA_INTEGRATION_ENABLED = bool(CANVA_API_TOKEN and CANVA_BRAND_TEMPLATE_ID and REQUESTS_AVAILABLE)

    if not api_key:
        st.error("الرجاء ضبط مفتاح GOOGLE_API_KEY في إعدادات الأمان (Secrets) أو متغيرات البيئة لتشغيل النظام.")
    else:
        client = genai.Client(api_key=api_key)

        grades = [
            "الصف الأول / Grade 1", "الصف الثاني / Grade 2", "الصف الثالث / Grade 3",
            "الصف الرابع / Grade 4", "الصف الخامس / Grade 5", "الصف السادس / Grade 6",
            "الصف السابع / Grade 7", "الصف الثامن / Grade 8", "الصف التاسع / Grade 9",
            "الصف العاشر / Grade 10"
        ]

        educational_systems = [
            "وطني (National)",
            "دولي (International)"
        ]

        jordan_governorates = [
            "العاصمة (عمان) / Capital (Amman)", "إربد / Irbid", "الزرقاء / Zarqa",
            "البلقاء / Balqa", "المفرق / Mafraq", "الكرك / Karak",
            "مادبا / Madaba", "جرش / Jerash", "عجلون / Ajloun",
            "معان / Ma'an", "الطفيلة / Tafilah", "العقبة / Aqaba"
        ]

        # كل مادة مرتبطة بأيقونة (إيموجي) لعرضها في شبكة البطاقات على طراز الصورة المرفقة
        subjects_with_icons = [
            ("➗", "الرياضيات / Mathematics"),
            ("🔬", "العلوم / Science"),
            ("📖", "اللغة العربية / Arabic"),
            ("🔤", "اللغة الإنجليزية / English"),
            ("🇫🇷", "اللغة الفرنسية / French"),
            ("🕌", "التربية الإسلامية / Islamic Ed."),
            ("✝️", "التربية المسيحية / Christian Ed."),
            ("🌍", "الدراسات الاجتماعية / Social St."),
            ("🗺️", "الجغرافيا / Geography"),
            ("📜", "التاريخ / History"),
            ("⚛️", "الفيزياء / Physics"),
            ("🧪", "الكيمياء / Chemistry"),
            ("🧬", "الأحياء / Biology"),
            ("💻", "الحاسوب / Computer Sci."),
            ("🎨", "الفنون / Arts"),
            ("🎵", "الموسيقى / Music"),
            ("🎭", "المسرح والدراما / Drama"),
            ("⚽", "التربية الرياضية / PE"),
            ("💼", "الأعمال وريادة الأعمال / Business"),
            ("🏛️", "التربية الوطنية والمدنية / Civics"),
            ("🧠", "علم النفس / Psychology"),
            ("👥", "علم الاجتماع / Sociology"),
            ("🛠️", "التصميم والتكنولوجيا / Design & Tech"),
            ("🍳", "الاقتصاد المنزلي / Home Economics"),
        ]
        subjects = [s[1] for s in subjects_with_icons]

        languages = [
            "ثنائي اللغة (عربي / إنجليزي) - Bilingual (Arabic / English)",
            "ثنائي اللغة (عربي / فرنسي) - Bilingual (Arabic / French)",
            "اللغة الفرنسية بالكامل - Pure French (Français)",
            "اللغة الإنجليزية بالكامل - Pure English",
            "اللغة العربية بالكامل - Pure Arabic"
        ]

        special_conditions_categories = {
            "1. الإعاقات الحسية والجسدية / Sensory & Physical Disabilities": [
                "الإعاقة البصرية (كف تام أو ضعف بصر شديد / بريل ومطبوعات كبيرة) / Visual Impairment (Blind/Low Vision - Braille & Large Print)",
                "الإعاقة السمعية (صمم تام أو ضعف سمعي بحاجة لمعينات/إشارة) / Hearing Impairment (Deaf/Hard of Hearing)",
                "الإعاقة الحركية أو الجسدية (شلل، ضمور عضلات، بتر أطراف، تشوهات) / Physical & Motor Disabilities",
                "الإعاقة الحسية المزدوجة (الصم-المكفوفين) / Deaf-Blindness"
            ],
            "2. الاضطرابات النمائية وصعوبات التعلم / Developmental Disorders & Learning Difficulties": [
                "صعوبات التعلم المحددة (ديسليكسيا، عسر كتابة، صعوبة حساب) / Specific Learning Difficulties (Dyslexia, Dysgraphia, Dyscalculia)",
                "اضطراب طيف التوحد (ASD) / Autism Spectrum Disorder (ASD)",
                "اضطراب فرط الحركة ونقص الانتباه (ADHD) / ADHD",
                "اضطرابات النطق واللغة والتواصل / Speech, Language & Communication Disorders"
            ],
            "3. الإعاقات الذهنية والسلوكية / Intellectual & Behavioral Disabilities": [
                "الإعاقة الذهنية / العقلية (بسيطة، متوسطة، شديدة) / Intellectual Disability (Mild, Moderate, Severe)",
                "الاضطرابات الانفعالية والسلوكية (قلق شديد، اكتئاب، مخاوف مدرسية) / Emotional & Behavioral Disorders",
                "الإعاقات المتعددة (أكثر من إعاقة معاً) / Multiple Disabilities"
            ],
            "4. الإعاقات والحالات الصحية المزمنة / Chronic Health Conditions": [
                "الأمراض المزمنة المحتاجة لمتابعة (سكري، ربو شديد، صرع) / Chronic Illnesses (Diabetes, Asthma, Epilepsy)",
                "مرضى السرطان (برامج استكمال وعلاجات مستمرة) / Cancer Support Programs",
                "حالات الفشل الكلوي (غسيل دوري) / Kidney Failure & Dialysis Care"
            ],
            "5. فئة الموهبة والتفوق / Giftedness & Talent": [
                "الطلبة الموهوبون والمتفوقون (برامج إثراء معرفي وتسريع أكاديمي) / Gifted & Talented Students (Enrichment & Acceleration)"
            ]
        }

        # =====================================================================================
        # === بنك إرشادات تكييف معتمدة لكل فئة حالة خاصة، تُحقن ضمن الطلب المرسل للذكاء
        # الاصطناعي كأساس ومرجع أسلوبي، بدل الاعتماد فقط على وصف الحالة النصي. هذا يرفع
        # جودة واتساق التكييف بدل أن يبدأ النموذج من الصفر في كل مرة. ===
        # =====================================================================================
        ADAPTATION_TEMPLATE_HINTS = {
            "1. الإعاقات الحسية والجسدية / Sensory & Physical Disabilities":
                "استخدم أوصافاً لفظية غنية بدل الاعتماد على الشكل البصري فقط، كبّر حجم النص المقترح، "
                "بسّط أي جداول إلى نقاط متسلسلة واضحة، وأضف وصفاً نصياً بديلاً لأي عنصر بصري.",
            "2. الاضطرابات النمائية وصعوبات التعلم / Developmental Disorders & Learning Difficulties":
                "قسّم كل مهمة إلى خطوات قصيرة مرقّمة، استخدم جملاً قصيرة ومباشرة، كرّر التعليمة الواحدة "
                "بصياغتين مختلفتين عند الحاجة، وأضف إشارة إلى رمز بصري توضيحي بجانب كل تعليمة.",
            "3. الإعاقات الذهنية والسلوكية / Intellectual & Behavioral Disabilities":
                "بسّط المفردات إلى الحد الأدنى الممكن، اربط كل سؤال بمثال حياتي مألوف للطالب، وقلّل عدد "
                "الأسئلة مقابل زيادة المساحة البصرية والوقت المتاح بين الفقرات.",
            "4. الإعاقات والحالات الصحية المزمنة / Chronic Health Conditions":
                "صمّم الورقة بحيث يمكن إنجازها على أكثر من جلسة قصيرة، أضف ملاحظة تشجيعية في ختامها، "
                "وتجنّب أي صياغة تعطي إحساساً بضغط زمني.",
            "5. فئة الموهبة والتفوق / Giftedness & Talent":
                "أضف تحدياً معرفياً إثرائياً اختيارياً بعد كل سؤال أساسي، اربط المحتوى بتطبيق واقعي أكثر "
                "تقدماً، وشجّع التفكير الناقد عبر أسئلة مفتوحة النهاية.",
        }

        # === تعليمات تنسيق ثابتة (Markdown) تُستخدم عند توليد ورقة العمل وعند توليد
        # الامتحان التقييمي لاحقاً على حد سواء، لضمان نفس مستوى الوضوح في الاثنين. ===
        FORMATTING_INSTRUCTIONS = """
            التزم حرفياً بقواعد التنسيق التالية أثناء الكتابة، لضمان الوضوح الكامل للقارئ
            (معلم أو طالب) دون أي إرهاق بصري:
            - اكتب العنوان الرئيسي كعنوان Markdown من المستوى الأول: # العنوان.
            - إن وجدت تعليمات عامة قبل الأسئلة (مثل "أجب عما يلي")، اكتبها بخط مائل
              *هكذا* في سطر مستقل قبل أول سؤال.
            - اكتب رقم وصياغة كل سؤال أو تمرين بخط عريض فقط، بالشكل: **السؤال ١: ...نص السؤال...**
            - اترك سطراً فارغاً كاملاً بعد كل سؤال، ثم سطراً فارغاً آخر قبل بدء السؤال التالي.
            - افصل بين كل سؤال/تمرين رئيسي والذي يليه بخط فاصل أفقي مستقل مكوّن من ثلاث
              شرطات فقط (---) على سطر خاص به وحده.
            - للاختيار من متعدد، اكتب كل خيار في سطر مستقل يبدأ بحرف أو رمز واضح
              (أ- ، ب- ، ج- ...)، ولا تكتب الخيارات متلاصقة في سطر واحد.
            - لا تكتب فقرات طويلة متراصة؛ اكسر كل فكرة أو خطوة في سطر أو فقرة قصيرة
              مستقلة، مع مسافة بصرية واضحة بين الفقرات.
            - إن وجدت مساحة مخصصة لكتابة إجابة الطالب، أشر إليها بوضوح بسطر يحتوي على
              نقاط توضيحية (مثال: الإجابة: ......................................).
        """

        adaptation_levels_with_icons = [
            ("⚖️", "تكييف متوازن وشامل\nBalanced Adaptation"),
            ("🧩", "تبسيط وتسهيل شديد للمفاهيم\nDeep Simplification"),
            ("🌟", "إثراء معرفي متقدم للموهوبين\nAdvanced Enrichment"),
            ("🖐️", "دمج بصري وحسي مكثف\nSensory & Visual Integration"),
        ]
        adaptation_levels_full = [
            "تكييف متوازن وشامل (Balanced Adaptation)",
            "تبسيط وتسهيل شديد للمفاهيم (Deep Simplification)",
            "إثراء معرفي متقدم للموهوبين (Advanced Enrichment)",
            "دمج بصري والحسي مكثف (Sensory & Visual Integration)"
        ]

        # =====================================================================================
        # === مكوّن شبكة بطاقات قابلة للنقر يحاكي واجهة التطبيق المرفقة (أيقونة + عنوان) ===
        # =====================================================================================
        def render_icon_grid(title, items, state_key, columns_per_row=4, default_index=0):
            """
            يعرض شبكة بطاقات (أيقونة + نص) بعدد أعمدة محدد، ويحفظ الاختيار في st.session_state.
            items: قائمة من tuples (emoji, label)
            """
            if state_key not in st.session_state:
                st.session_state[state_key] = default_index

            st.markdown(f'<div class="grid-title">{title}</div>', unsafe_allow_html=True)

            for row_start in range(0, len(items), columns_per_row):
                row_items = items[row_start: row_start + columns_per_row]
                cols = st.columns(len(row_items))
                for col, (idx_in_row, (emoji, label)) in zip(cols, enumerate(row_items)):
                    real_idx = row_start + idx_in_row
                    is_selected = st.session_state[state_key] == real_idx
                    btn_label = f"{'✅ ' if is_selected else ''}{emoji}\n{label}"
                    with col:
                        if st.button(btn_label, key=f"{state_key}_btn_{real_idx}"):
                            st.session_state[state_key] = real_idx
                            st.rerun()

            return st.session_state[state_key]

        # =====================================================================================
        # === شبكة دوائر أنيقة خاصة بالمواد الدراسية (٤ دوائر بجانب بعضها في كل صف) ===
        # =====================================================================================
        def render_subject_circles(title, items, state_key, columns_per_row=4, default_index=0):
            """
            يعرض المواد الدراسية على شكل دوائر (أيقونة داخل الدائرة + اسم المادة أسفلها)،
            ويحفظ الاختيار في st.session_state. items: قائمة من tuples (emoji, label)
            """
            if state_key not in st.session_state:
                st.session_state[state_key] = default_index

            st.markdown(f'<div class="grid-title">{title}</div>', unsafe_allow_html=True)

            for row_start in range(0, len(items), columns_per_row):
                row_items = items[row_start: row_start + columns_per_row]
                cols = st.columns(len(row_items))
                for col, (idx_in_row, (emoji, label)) in zip(cols, enumerate(row_items)):
                    real_idx = row_start + idx_in_row
                    is_selected = st.session_state[state_key] == real_idx
                    with col:
                        on_off = "on" if is_selected else "off"
                        container_key = f"circlebtn-{state_key}-{real_idx}-{on_off}"
                        try:
                            circle_container = st.container(key=container_key)
                        except TypeError:
                            circle_container = st.container()
                        with circle_container:
                            if st.button(emoji, key=f"{state_key}_circle_{real_idx}"):
                                st.session_state[state_key] = real_idx
                                st.rerun()
                        check_mark = "✅ " if is_selected else ""
                        st.markdown(f'<div class="circle-caption">{check_mark}{label}</div>', unsafe_allow_html=True)

            return st.session_state[state_key]

        # ---------------- صندوق أبيض واحد يلف شبكة دوائر اختيار المادة الدراسية بالكامل ----------------
        try:
            subjects_white_card = st.container(key="subjects-white-card")
        except TypeError:
            subjects_white_card = st.container()

        with subjects_white_card:
            subject_idx = render_subject_circles(
                "📚 اختر المادة الدراسية / Select Subject",
                subjects_with_icons,
                "subject_idx",
                columns_per_row=4,
                default_index=0
            )
        selected_subject = subjects[subject_idx]

        # ---------------- شبكة اختيار مستوى ونوع التكييف ----------------
        level_idx = render_icon_grid(
            "🎯 اختر مستوى وطبيعة التكييف / Select Adaptation Level",
            adaptation_levels_with_icons,
            "level_idx",
            columns_per_row=1,
            default_index=0
        )
        selected_level = adaptation_levels_full[level_idx]

        # ---------------- شبكة اختيار وضع الإخراج (تكييف الأصل أو توليد بديل) ----------------
        mode_items = [
            ("📝", "تكييف الورقة الأصلية"),
            ("🆕", "ورقة بديلة + بنك أسئلة"),
        ]
        mode_idx = render_icon_grid(
            "🗂️ اختر وضع التوليد / Select Output Mode",
            mode_items,
            "mode_idx",
            columns_per_row=2,
            default_index=0
        )
        generate_alternative = (mode_idx == 1)

        st.markdown(f"""
            <div class="selection-summary">
                المختار حالياً: {subjects_with_icons[subject_idx][1]} &nbsp;|&nbsp;
                {selected_level} &nbsp;|&nbsp;
                {mode_items[mode_idx][1]}
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # =====================================================================================
        # === قسم إدارة الطلاب — اختيار طالب موجود أو إضافة طالب جديد، مربوط باسم المعلم
        # المسجّل في الشريط الجانبي. عند اختيار طالب موجود، تُستخدم بياناته كقيم افتراضية
        # لحقول الصف/النظام/الفئة/الحالة أدناه بدل البدء من الصفر في كل مرة. ===
        # =====================================================================================
        current_teacher = (st.session_state.get("teacher_name") or "معلم_عام").strip()

        try:
            student_mgmt_card = st.container(key="student-mgmt-card")
        except TypeError:
            student_mgmt_card = st.container()

        with student_mgmt_card:
            st.markdown('<div class="grid-title">👨‍🎓 إدارة الطالب / Student Management</div>', unsafe_allow_html=True)

            students_list = get_students(current_teacher)
            student_names_options = ["➕ بدون ربط بطالب / No Student Link", "✏️ طالب جديد / New Student"] + [
                f"{s['full_name']} ({s['grade']})" for s in students_list
            ]
            student_choice = st.selectbox("الطالب / Student:", student_names_options, key="student_choice")

            selected_student_record = None
            new_student_name = ""
            if student_choice == student_names_options[1]:
                new_student_name = st.text_input("اسم الطالب الجديد / New Student Name:", key="new_student_name_input")
            elif student_choice not in (student_names_options[0], student_names_options[1]):
                _sel_idx = student_names_options.index(student_choice) - 2
                selected_student_record = students_list[_sel_idx]

        def _idx_or_default(options_list, value, default=0):
            try:
                return options_list.index(value)
            except (ValueError, TypeError):
                return default

        # ---------------- بقية الحقول عبر قوائم منسدلة بنفس الهوية اللونية (كحلي/ذهبي/أبيض) ----------------
        grade_default_idx = _idx_or_default(grades, selected_student_record["grade"]) if selected_student_record else 0
        selected_grade = st.selectbox("اختر الصف الدراسي / Select Grade:", grades, index=grade_default_idx)

        system_default_idx = _idx_or_default(educational_systems, selected_student_record["system"]) if selected_student_record else 0
        selected_system = st.selectbox("اختر النظام التعليمي / Select Educational System:", educational_systems, index=system_default_idx)

        selected_language = st.selectbox("اختر لغة التكييف والمخرجات / Select Output Language / Langue:", languages)
        selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن / Select Governorate in Jordan:", jordan_governorates)

        category_options = list(special_conditions_categories.keys())
        category_default_idx = _idx_or_default(category_options, selected_student_record["category"]) if selected_student_record else 0
        selected_category = st.selectbox("اختر فئة الحالة الخاصة / Select Special Condition Category:", category_options, index=category_default_idx)

        condition_options = special_conditions_categories[selected_category]
        condition_default_idx = 0
        if selected_student_record and selected_student_record.get("category") == selected_category:
            condition_default_idx = _idx_or_default(condition_options, selected_student_record["condition"])
        selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة / Select Specific Condition:", condition_options, index=condition_default_idx)

        # عرض شفاف لإرشاد التكييف المعتمد لهذه الفئة حتى يطّلع عليه المعلم مباشرة
        st.caption(f"🗂️ إرشاد تكييف معتمد لهذه الفئة: {ADAPTATION_TEMPLATE_HINTS.get(selected_category, '')}")

        # ---------------- أزرار حفظ / تحديث بيانات الطالب (تُستخدم بعد اختيار كل الحقول أعلاه) ----------------
        if student_choice == student_names_options[1] and new_student_name.strip():
            if st.button("💾 حفظ الطالب الجديد / Save New Student", key="save_new_student_btn"):
                save_student(current_teacher, new_student_name.strip(), selected_grade, selected_system,
                             selected_category, selected_condition)
                st.success("تم حفظ بيانات الطالب بنجاح.")
                st.rerun()
        elif selected_student_record:
            if st.button("💾 تحديث بيانات الطالب / Update Student Info", key="update_student_btn"):
                update_student(selected_student_record["id"], selected_grade, selected_system,
                                selected_category, selected_condition)
                st.success("تم تحديث بيانات الطالب.")

        # =====================================================================================
        # === تقرير متابعة الطالب — يعرض سجل كل أوراق العمل السابقة المرتبطة بطالب معيّن،
        # مع إمكانية تصدير تقرير PDF من صفحة واحدة لمشاركته مع قسم الإرشاد الطلابي أو ولي الأمر. ===
        # =====================================================================================
        with st.expander("📊 تقرير متابعة الطالب / Student Progress Report"):
            if students_list:
                report_student_choice = st.selectbox(
                    "اختر الطالب لعرض سجله / Select Student:",
                    [s["full_name"] for s in students_list],
                    key="report_student_select"
                )
                _chosen_report_student = next(s for s in students_list if s["full_name"] == report_student_choice)
                history_rows = get_student_history(current_teacher, _chosen_report_student["id"])
                if history_rows:
                    st.dataframe(
                        [
                            {
                                "التاريخ": r["created_at"][:16].replace("T", " "),
                                "المادة": r["subject"],
                                "الصف": r["grade"],
                                "مستوى التكييف": r["adaptation_level"],
                                "الوضع": r["mode"],
                            }
                            for r in history_rows
                        ],
                        use_container_width=True
                    )
                else:
                    st.info("لا يوجد سجل أوراق سابق لهذا الطالب بعد.")
            else:
                st.info("لم تتم إضافة أي طلاب بعد. أضف طالباً من قسم إدارة الطالب أعلاه.")

        # === إصلاح: أضفنا key ثابت وصريح لهذا العنصر. بدون key، يعتمد ستريمليت على
        # ترتيب/موضع العنصر بالصفحة لتحديد "هويته" بين كل rerun، وهذه الصفحة فيها
        # عشرات الأزرار (شبكة المواد، مستوى التكييف، إدارة الطالب...) وكل واحد منها
        # يستدعي st.rerun() عند الضغط عليه — أي تغيّر بسيط بعدد/ترتيب العناصر قبل
        # رافع الملفات (مثلاً ظهور/اختفاء زر "حفظ الطالب" حسب الاختيار) قد يجعل
        # ستريمليت يعتبره عنصراً "جديداً" ويهمل الملف الذي اختاره المستخدم للتو. الـ
        # key الثابت يضمن أن هوية العنصر لا تتغير أبداً مهما تغيّر ما حوله. ===
        uploaded_file = st.file_uploader(
            "قم بتمرير أو رفع ملف ورقة العمل (PDF أو Word أو TXT) / Upload Worksheet File:",
            type=["pdf", "docx", "txt"],
            key="worksheet_file_uploader",
        )

        MAX_INPUT_CHARS = 20000

        extracted_content = ""
        if uploaded_file is not None:
            # === إصلاح: تأكيد استلام فوري قبل أي معالجة — إن لم تظهر هذه الرسالة أصلاً
            # بعد اختيار الملف، فالمشكلة في نقل الملف نفسه من المتصفح للخادم (شبكة/متصفح
            # الجوال) وليست بمنطق قراءة الملف بالأسفل، وهذا يسهّل تشخيص أي عطل مستقبلي. ===
            st.caption(f"📎 تم استلام الملف: {uploaded_file.name} ({uploaded_file.size:,} بايت)")

            file_extension = uploaded_file.name.split(".")[-1].lower()
            # === إصلاح: نستخدم st.toast() بالإضافة إلى الرسالة العادية، لأن صفحة التطبيق
            # طويلة جداً وقد تظهر رسالة النجاح/الخطأ أسفل الشاشة دون أن يلاحظها المستخدم
            # (خصوصاً على الجوال بعد إغلاق نافذة اختيار الملف). st.toast يظهر كإشعار عائم
            # مؤقت في الزاوية بغض النظر عن مكان التمرير الحالي بالصفحة. ===
            try:
                if file_extension == "txt":
                    raw_bytes = uploaded_file.getvalue()
                    # === إصلاح: بعض ملفات الـ TXT (خصوصاً المحفوظة من Windows بالعربية)
                    # لا تكون UTF-8، فكانت تفشل بخطأ ترميز (UnicodeDecodeError) يظهر
                    # كرسالة تقنية غير مفهومة. الآن نجرّب عدة ترميزات شائعة قبل الاستسلام. ===
                    extracted_content = None
                    for encoding_name in ("utf-8", "utf-8-sig", "cp1256", "windows-1256"):
                        try:
                            extracted_content = raw_bytes.decode(encoding_name)
                            break
                        except UnicodeDecodeError:
                            continue
                    if extracted_content is None:
                        extracted_content = raw_bytes.decode("utf-8", errors="replace")

                elif file_extension == "docx":
                    if not DOCX_AVAILABLE:
                        st.warning("⚠️ دعم ملفات Word (.docx) غير مفعّل على الخادم حالياً. الرجاء رفع الملف بصيغة PDF أو TXT بدلاً منه.")
                        st.toast("⚠️ دعم ملفات Word غير متاح حالياً", icon="⚠️")
                    else:
                        doc = Document(uploaded_file)
                        extracted_content = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

                elif file_extension == "pdf":
                    pdf_reader = pypdf.PdfReader(uploaded_file)
                    extracted_content = ""
                    for page in pdf_reader.pages:
                        text = page.extract_text()
                        if text:
                            extracted_content += text + "\n"

                if extracted_content.strip():
                    st.success(f"تم قراءة الملف بنجاح / File successfully read: {uploaded_file.name}")
                    st.toast(f"✅ تم قراءة الملف بنجاح: {uploaded_file.name}", icon="✅")
                    if len(extracted_content) > MAX_INPUT_CHARS:
                        st.info(
                            f"⚠️ الملف المرفوع يحتوي على {len(extracted_content):,} حرفاً، وسيتم إرسال أول "
                            f"{MAX_INPUT_CHARS:,} حرف فقط منه إلى نموذج الذكاء الاصطناعي بسبب حدود المعالجة الحالية. "
                            "لضمان تغطية الورقة كاملة، يُفضّل تقسيم الملفات الطويلة إلى أجزاء أصغر قبل الرفع."
                        )
                elif file_extension != "docx" or DOCX_AVAILABLE:
                    st.warning("⚠️ الملف المرفوع لا يحتوي على نص قابل للقراءة المباشرة. سيتم الاعتماد على معلومات النظام والعنوان لتوليد ورقة العمل.")
                    st.toast("⚠️ لم يُستخرج نص من الملف", icon="⚠️")
            except Exception as e:
                # === إصلاح أمني: لا نعرض نص الاستثناء الخام للمستخدم (قد يكشف تفاصيل
                # داخلية عن مكتبات الخادم)، بل رسالة عامة واضحة، مع تسجيل التفصيل الكامل
                # بسجلات الخادم فقط لتشخيصه لاحقاً. ===
                log_internal_error(f"فشل قراءة الملف المرفوع ({uploaded_file.name})", e)
                st.error(
                    "⚠️ حدث خطأ أثناء قراءة هذا الملف. تأكد أنه ملف PDF أو Word (.docx) أو نصي (.txt) سليم "
                    "وغير تالف أو محمي بكلمة مرور، ثم أعد المحاولة."
                )
                st.toast("❌ تعذّرت قراءة الملف", icon="❌")

        # =====================================================================================
        # === دعم كامل لاتجاه RTL الصحيح في مستندات Word (محاذاة يمين + خاصية bidi فعلية) ===
        # =====================================================================================
        def _set_paragraph_rtl(paragraph):
            """يضبط الفقرة لتكون بمحاذاة اليمين وباتجاه RTL فعلي (وليس فقط محاذاة بصرية)."""
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pPr = paragraph._p.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            bidi.set(qn('w:val'), "1")
            pPr.append(bidi)

        def _add_bold_line(doc, raw_line):
            """يكتب سطراً كاملاً بخط عريض (يُستخدم لأسطر الأسئلة **...** في ملف Word)."""
            p = doc.add_paragraph()
            run = p.add_run(raw_line.strip("* ").strip())
            run.bold = True
            _set_paragraph_rtl(p)
            return p

        def create_word_file(text):
            """
            غلاف آمن: يلتقط أي استثناء غير متوقع أثناء بناء ملف Word بدل ترك التطبيق
            ينهار بالكامل (نفس مبدأ الحماية المطبّق على PDF وPowerPoint). ===
            """
            if not DOCX_AVAILABLE:
                return None
            try:
                return _create_word_file_impl(text)
            except Exception as e:
                log_internal_error("فشل توليد ملف Word", e)
                return None

        def _create_word_file_impl(text):
            """
            === تحسين: يفسّر تنسيق Markdown الخفيف الذي يطلبه التطبيق من الذكاء الاصطناعي
            (# عناوين، **أسئلة بخط عريض**، --- كفاصل بين التمارين) ويحوّله لتنسيق حقيقي
            داخل ملف Word بدل نسخ النص كأسطر عادية متطابقة الشكل — لتظل الورقة واضحة
            ومنظمة حتى بعد تحميلها وطباعتها، وليس فقط عند عرضها على الشاشة. ===
            """
            doc = Document()
            heading = doc.add_heading('ورقة العمل المطورة (التربية الخاصة) / Adapted Worksheet', 0)
            _set_paragraph_rtl(heading)

            for raw_line in text.split('\n'):
                line = raw_line.strip()

                if not line:
                    doc.add_paragraph("")
                    continue

                if line == "---":
                    divider = doc.add_paragraph("―" * 25)
                    _set_paragraph_rtl(divider)
                    continue

                # === سطر "إجابة نموذجية" مميّز بصرياً (Blockquote) — يأتي فقط من ورقة
                # الامتحان بعد معالجتها في _format_exam_text_for_display؛ هنا نزيل رموز
                # الـ Markdown الخام (> و ** و ✅) ونكتبه كسطر عريض واحد مقروء في Word
                # بدل أن تظهر رموز التنسيق حرفياً كنص غير مفهوم داخل المستند. ===
                if line.startswith(">"):
                    clean_line = line.lstrip(">").strip()
                    clean_line = clean_line.replace("**", "")
                    _add_bold_line(doc, clean_line)
                    continue

                if line.startswith("### "):
                    h = doc.add_heading(line[4:].strip(), level=3)
                    _set_paragraph_rtl(h)
                    continue
                if line.startswith("## "):
                    h = doc.add_heading(line[3:].strip(), level=2)
                    _set_paragraph_rtl(h)
                    continue
                if line.startswith("# "):
                    h = doc.add_heading(line[2:].strip(), level=1)
                    _set_paragraph_rtl(h)
                    continue

                if line.startswith("**") and line.endswith("**") and len(line) > 4:
                    _add_bold_line(doc, line)
                    continue

                if line.startswith("*") and line.endswith("*") and not line.startswith("**"):
                    p = doc.add_paragraph()
                    run = p.add_run(line.strip("* "))
                    run.italic = True
                    _set_paragraph_rtl(p)
                    continue

                p = doc.add_paragraph(line)
                _set_paragraph_rtl(p)

            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            return bio

        # =====================================================================================
        # === ميزة "استمع للورقة" — تحويل نص الورقة المكيّفة لملف صوتي (MP3) بصوت عربي عبر
        # gTTS، لدعم الطلاب ذوي الإعاقة البصرية أو صعوبات القراءة الذين يحتاجون سماع محتوى
        # الورقة بدل قراءته. نفس مبدأ الحماية المستخدم بباقي دوال التوليد أعلاه: أي خطأ
        # (فشل اتصال بخدمة جوجل الصوتية، أو أي استثناء آخر) يُسجَّل بسجلات الخادم فقط
        # ويُعاد None + رسالة عامة بدل انهيار الصفحة. ===
        # =====================================================================================
        MAX_TTS_CHARS = 4000

        def create_worksheet_audio(text):
            """
            يعيد tuple: (BytesIO صوت MP3 أو None، رسالة تحذير/معلومة أو None).
            """
            if not GTTS_AVAILABLE:
                return None, "⚠️ ميزة الاستماع الصوتي غير مفعّلة على الخادم حالياً. أضف السطر 'gTTS' إلى requirements.txt ثم أعد تشغيل التطبيق (Reboot app)."
            clean_text = text.replace("#", "").replace("*", "").replace(">", "").strip()
            if not clean_text:
                return None, "لا يوجد نص كافٍ لتحويله إلى صوت."
            truncated = False
            if len(clean_text) > MAX_TTS_CHARS:
                clean_text = clean_text[:MAX_TTS_CHARS]
                truncated = True
            try:
                tts = gTTS(text=clean_text, lang="ar")
                bio = io.BytesIO()
                tts.write_to_fp(bio)
                bio.seek(0)
                warning = (
                    f"⚠️ تم تحويل أول {MAX_TTS_CHARS:,} حرف فقط من الورقة إلى صوت بسبب حدود المدة الحالية."
                    if truncated else None
                )
                return bio, warning
            except Exception as e:
                log_internal_error("فشل توليد الملف الصوتي لورقة العمل", e)
                return None, "⚠️ تعذّر توليد الملف الصوتي حالياً (قد تكون خدمة الصوت مؤقتاً غير متاحة). يرجى المحاولة لاحقاً."

        # =====================================================================================
        # === تصميم بصري احترافي لملف PowerPoint (بديل محلي لا يحتاج إنترنت أو حساب Canva) ===
        # =====================================================================================

        # لوحة ألوان احترافية متناسقة مع هوية النظام (كحلي داكن + ذهبي + أبيض + أزرق فاتح)
        PPTX_THEME = {
            "dark_navy": RGBColor(0x10, 0x1B, 0x2D),
            "gold": RGBColor(0xF1, 0xC4, 0x0F),
            "teal": RGBColor(0x2E, 0x6F, 0xBB),
            "white": RGBColor(0xFF, 0xFF, 0xFF),
            "light_bg": RGBColor(0xEA, 0xF1, 0xFB),
            "text_dark": RGBColor(0x10, 0x1B, 0x2D),
        }

        def _draw_icon(kind, size=200, fg=(241, 196, 15, 255), bg=(16, 27, 45, 255)):
            """
            يرسم أيقونة بسيطة (شرح صوري/بصري) باستخدام PIL بدون أي اتصال بالإنترنت،
            وتُستخدم كصور توضيحية داخل شرائح PowerPoint وبطاقات PECS — بديل فوري
            وسريع جداً بدل الاعتماد الحصري على توليد صور الذكاء الاصطناعي.
            """
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            pad = size * 0.12
            d.ellipse([pad*0.4, pad*0.4, size-pad*0.4, size-pad*0.4], fill=bg)

            if kind == "check":
                d.line([(size*0.28, size*0.52), (size*0.44, size*0.68), (size*0.74, size*0.32)],
                       fill=fg, width=int(size*0.09), joint="curve")
            elif kind == "star":
                import math
                cx, cy, r_out, r_in = size/2, size/2, size*0.34, size*0.15
                pts = []
                for i in range(10):
                    ang = math.pi/2 + i * math.pi/5
                    r = r_out if i % 2 == 0 else r_in
                    pts.append((cx + r*math.cos(ang), cy - r*math.sin(ang)))
                d.polygon(pts, fill=fg)
            elif kind == "idea":
                d.ellipse([size*0.32, size*0.20, size*0.68, size*0.56], fill=fg)
                d.rectangle([size*0.42, size*0.55, size*0.58, size*0.68], fill=fg)
                d.rectangle([size*0.40, size*0.70, size*0.60, size*0.76], fill=fg)
            elif kind == "book":
                d.rectangle([size*0.24, size*0.30, size*0.76, size*0.72], outline=fg, width=int(size*0.05))
                d.line([(size*0.5, size*0.30), (size*0.5, size*0.72)], fill=fg, width=int(size*0.04))
            elif kind == "target":
                d.ellipse([size*0.24, size*0.24, size*0.76, size*0.76], outline=fg, width=int(size*0.05))
                d.ellipse([size*0.38, size*0.38, size*0.62, size*0.62], fill=fg)
            else:  # pencil (افتراضي)
                d.polygon([(size*0.30, size*0.72), (size*0.62, size*0.30), (size*0.72, size*0.40), (size*0.40, size*0.82)], fill=fg)

            bio = io.BytesIO()
            img.save(bio, format="PNG")
            bio.seek(0)
            return bio

        def _set_slide_background(slide, rgb):
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = rgb

        # === إصلاح أداء: توليد صور الذكاء الاصطناعي أصبح مقيّداً بمفتاح enable_ai_images
        # (من قسم "إعدادات الأداء والسرعة" أعلى الصفحة). إن كان مطفأً، تُتخطى استدعاءات
        # الشبكة نهائياً بدل محاولتها ثم الفشل أو الانتظار — وهذا هو السبب الرئيسي الثاني
        # في بطء توليد ملفات PowerPoint وبطاقات PECS في النسخة القديمة.
        _ai_images_state = {"available": bool(enable_ai_images)}

        # نماذج الصور الحالية الفعّالة (نموذج Imagen القديم توقف رسمياً من جوجل بتاريخ ١٧/٨/٢٠٢٦)
        IMAGE_MODELS_TO_TRY = ["gemini-3.1-flash-image", "gemini-3.1-flash-lite-image", "gemini-2.5-flash-image"]

        def _generate_ai_illustration(prompt_text):
            if not _ai_images_state["available"]:
                return None

            full_prompt = (
                "رسمة تعليمية بسيطة بأسلوب Flat Design نظيف وواضح، بدون أي كتابة أو حروف "
                "أو أرقام داخل الصورة إطلاقاً، بألوان هادئة تتناسق مع الذهبي (#F1C40F) "
                f"والكحلي الداكن (#101B2D)، توضّح بصرياً الفكرة التالية: {prompt_text}"
            )

            for model_name in IMAGE_MODELS_TO_TRY:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=full_prompt,
                    )
                    for part in response.candidates[0].content.parts:
                        if getattr(part, "inline_data", None) is not None:
                            img_bytes = part.inline_data.data
                            bio = io.BytesIO(img_bytes)
                            bio.seek(0)
                            return bio
                except Exception:
                    continue

            # لو فشلت كل النماذج الحالية، نعطّل محاولات توليد الصور لبقية الجلسة
            # (توفيراً للوقت) وترجع الشرائح للاعتماد على الأيقونات المرسومة بديلاً
            _ai_images_state["available"] = False
            return None

        def _add_footer(slide, prs, page_num):
            left = Inches(0.3)
            top = prs.slide_height - Inches(0.42)
            width = prs.slide_width - Inches(0.6)
            box = slide.shapes.add_textbox(left, top, width, Inches(0.3))
            tf = box.text_frame
            p = tf.paragraphs[0]
            p.text = f"Edu Worksheet Adapt  •  {selected_subject.split(' / ')[0]}  •  {page_num}"
            p.font.size = Pt(10)
            p.font.color.rgb = PPTX_THEME["dark_navy"]
            p.alignment = PP_ALIGN.RIGHT

        # =====================================================================================
        # === تقسيم أذكى لمحتوى الشرائح — يجمع حسب الفقرات الطبيعية (فواصل الأسطر الفارغة)
        # بدل تقسيم كل 5 أسطر بشكل عشوائي قد يقطع سؤالاً أو فكرة في المنتصف ===
        # =====================================================================================
        def split_into_slide_blocks(text, max_chars_per_slide=420):
            raw_blocks = [b.strip() for b in text.split('\n\n') if b.strip()]
            if not raw_blocks:
                raw_blocks = [line.strip() for line in text.split('\n') if line.strip()]

            slides = []
            current_lines = []
            current_len = 0
            for block in raw_blocks:
                block_len = len(block)
                if block_len > max_chars_per_slide:
                    # الكتلة نفسها طويلة جداً — نقسمها على أسطرها الداخلية بدل تركها تفيض من الشريحة
                    if current_lines:
                        slides.append(current_lines)
                        current_lines, current_len = [], 0
                    sub_lines = [l.strip() for l in block.split('\n') if l.strip()]
                    for l in sub_lines:
                        if current_len + len(l) > max_chars_per_slide and current_lines:
                            slides.append(current_lines)
                            current_lines, current_len = [], 0
                        current_lines.append(l)
                        current_len += len(l)
                    continue

                if current_len + block_len > max_chars_per_slide and current_lines:
                    slides.append(current_lines)
                    current_lines, current_len = [], 0

                for l in block.split('\n'):
                    if l.strip():
                        current_lines.append(l.strip())
                current_len += block_len

            if current_lines:
                slides.append(current_lines)

            return slides if slides else [[text]]

        def create_ppt_file(text):
            """
            غلاف آمن: يلتقط أي استثناء غير متوقع أثناء بناء ملف PowerPoint (مثلاً فشل
            تضمين صورة ذكاء اصطناعي بصيغة غير مدعومة، أو أي خطأ آخر في مكتبة python-pptx)
            بدل ترك التطبيق ينهار بالكامل. === إصلاح مهم: قبل هذا التعديل لم يكن هناك أي
            حماية هنا على الإطلاق (خلافاً لملف PDF الذي كان محمياً بالفعل عبر
            _create_pdf_file_impl)، فأي استثناء أثناء توليد PowerPoint كان يصعد دون
            معالجة إلى future_ppt.result() في خطوة التوليد المتوازي أدناه، فيتسبب في
            انهيار الصفحة بالكامل ويسحب معه أزرار تحميل Word وPDF التي تُبنى بالتوازي
            في نفس الخطوة — وهذا هو سبب تعطّل "كل" أزرار التحميل أحياناً وليس فقط زر
            PowerPoint. ===
            """
            if not PPTX_AVAILABLE:
                return None
            try:
                return _create_ppt_file_impl(text)
            except Exception as e:
                log_internal_error("فشل توليد ملف PowerPoint", e)
                return None

        def _create_ppt_file_impl(text):
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            icon_cycle = ["idea", "check", "star", "book", "target", "pencil"]

            slide = prs.slides.add_slide(prs.slide_layouts[6])
            _set_slide_background(slide, PPTX_THEME["dark_navy"])

            band = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(-1), Inches(-0.6), Inches(9), Inches(2.2))
            band.fill.solid()
            band.fill.fore_color.rgb = PPTX_THEME["gold"]
            band.line.fill.background()
            band.shadow.inherit = False

            band2 = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(9.5), Inches(5.2), Inches(6), Inches(2.6))
            band2.fill.solid()
            band2.fill.fore_color.rgb = PPTX_THEME["teal"]
            band2.line.fill.background()
            band2.shadow.inherit = False

            for filename in ["new_logo.png", "Educ_Worksheet_Adapt_Icon_(Square).png", "logo.png", "logo.jpg"]:
                if os.path.exists(filename):
                    try:
                        slide.shapes.add_picture(filename, Inches(0.7), Inches(0.25), height=Inches(1.9))
                    except Exception as e:
                        log_internal_error(f"فشل تضمين ملف الشعار {filename} في PowerPoint", e)
                    break

            cover_illustration = _generate_ai_illustration(
                f"موضوع مادة {selected_subject} لطلاب {selected_grade}"
            )
            cover_embedded = False
            if cover_illustration:
                # === إصلاح: صورة الذكاء الاصطناعي قد تصل بصيغة لا تدعمها python-pptx
                # (مثل WEBP)، وهذا كان يُسقط الملف بالكامل قبل الإصلاح. الآن نرجع فوراً
                # للأيقونة المرسومة محلياً بدل انهيار كل الشرائح. ===
                try:
                    slide.shapes.add_picture(cover_illustration, Inches(8.6), Inches(1.9), height=Inches(3.4))
                    cover_embedded = True
                except Exception as e:
                    log_internal_error("فشل تضمين صورة الذكاء الاصطناعي في غلاف PowerPoint", e)
            if not cover_embedded:
                icon_bio = _draw_icon("idea", size=400, fg=(0x10, 0x1B, 0x2D, 255), bg=(0xF1, 0xC4, 0x0F, 255))
                slide.shapes.add_picture(icon_bio, Inches(9.3), Inches(2.4), height=Inches(2.6))

            title_box = slide.shapes.add_textbox(Inches(0.6), Inches(2.7), Inches(7.5), Inches(1.7))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "ورقة العمل المكيّفة"
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = PPTX_THEME["gold"]
            p.alignment = PP_ALIGN.RIGHT
            p_en = tf.add_paragraph()
            p_en.text = "Adapted Worksheet"
            p_en.font.size = Pt(20)
            p_en.font.color.rgb = PPTX_THEME["white"]
            p_en.alignment = PP_ALIGN.RIGHT

            subtitle_box = slide.shapes.add_textbox(Inches(0.6), Inches(4.6), Inches(7.5), Inches(1.6))
            tf2 = subtitle_box.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = f"{selected_grade}"
            p2.font.size = Pt(18)
            p2.font.color.rgb = PPTX_THEME["white"]
            p2.alignment = PP_ALIGN.RIGHT
            p2b = tf2.add_paragraph()
            p2b.text = f"{selected_subject}"
            p2b.font.size = Pt(18)
            p2b.font.color.rgb = PPTX_THEME["white"]
            p2b.alignment = PP_ALIGN.RIGHT
            p3 = tf2.add_paragraph()
            p3.text = f"{selected_condition.split(' / ')[0]}   |   {selected_level.split('(')[0]}"
            p3.font.size = Pt(14)
            p3.font.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
            p3.alignment = PP_ALIGN.RIGHT

            slide_blocks = split_into_slide_blocks(text)
            page_num = 1
            for chunk in slide_blocks:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                _set_slide_background(slide, PPTX_THEME["light_bg"])

                header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.15))
                header.fill.solid()
                header.fill.fore_color.rgb = PPTX_THEME["dark_navy"]
                header.line.fill.background()
                header.shadow.inherit = False

                accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.15), prs.slide_width, Inches(0.08))
                accent.fill.solid()
                accent.fill.fore_color.rgb = PPTX_THEME["gold"]
                accent.line.fill.background()
                accent.shadow.inherit = False

                header_tf = slide.shapes.add_textbox(Inches(0.6), Inches(0.22), Inches(10.5), Inches(0.8)).text_frame
                hp = header_tf.paragraphs[0]
                hp.text = f"محطة العرض التفاعلي {page_num} / Interactive Station {page_num}"
                hp.font.size = Pt(24)
                hp.font.bold = True
                hp.font.color.rgb = PPTX_THEME["white"]
                hp.alignment = PP_ALIGN.RIGHT

                slide_illustration = None
                if page_num <= 3:
                    slide_illustration = _generate_ai_illustration(chunk[0][:120])

                slide_illustration_embedded = False
                if slide_illustration:
                    try:
                        slide.shapes.add_picture(slide_illustration, Inches(10.0), Inches(1.35), height=Inches(2.2))
                        slide_illustration_embedded = True
                        content_width = Inches(9.1)
                    except Exception as e:
                        log_internal_error("فشل تضمين صورة الذكاء الاصطناعي في شريحة PowerPoint", e)
                if not slide_illustration_embedded:
                    icon_kind = icon_cycle[(page_num - 1) % len(icon_cycle)]
                    icon_bio = _draw_icon(icon_kind, size=220,
                                           fg=(0x10, 0x1B, 0x2D, 255), bg=(0xF1, 0xC4, 0x0F, 255))
                    slide.shapes.add_picture(icon_bio, Inches(11.3), Inches(0.15), height=Inches(0.9))
                    content_width = Inches(11.9)

                body_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), content_width, Inches(5.5))
                body_tf = body_box.text_frame
                body_tf.word_wrap = True

                for idx, line in enumerate(chunk):
                    p = body_tf.paragraphs[0] if idx == 0 else body_tf.add_paragraph()
                    p.text = f"◆  {line}"
                    p.font.size = Pt(19)
                    p.font.color.rgb = PPTX_THEME["text_dark"]
                    p.alignment = PP_ALIGN.RIGHT
                    p.space_after = Pt(14)

                _add_footer(slide, prs, page_num)
                page_num += 1

            closing = prs.slides.add_slide(prs.slide_layouts[6])
            _set_slide_background(closing, PPTX_THEME["dark_navy"])
            cbox = closing.shapes.add_textbox(Inches(1), Inches(3.1), Inches(11.3), Inches(1.3))
            ctf = cbox.text_frame
            cp = ctf.paragraphs[0]
            cp.text = "شكراً لاستخدامك Edu Worksheet Adapt"
            cp.font.size = Pt(30)
            cp.font.bold = True
            cp.font.color.rgb = PPTX_THEME["gold"]
            cp.alignment = PP_ALIGN.CENTER

            bio = io.BytesIO()
            prs.save(bio)
            bio.seek(0)
            return bio

        def create_canva_design(text, title_text):
            if not CANVA_INTEGRATION_ENABLED:
                return None, "لم يتم ضبط بيانات اعتماد Canva (CANVA_API_TOKEN / CANVA_BRAND_TEMPLATE_ID)."

            headers = {
                "Authorization": f"Bearer {CANVA_API_TOKEN}",
                "Content-Type": "application/json",
            }
            payload = {
                "brand_template_id": CANVA_BRAND_TEMPLATE_ID,
                "data": {
                    "content": {"type": "text", "text": text[:2000]},
                    "title": {"type": "text", "text": title_text},
                },
            }
            try:
                create_resp = requests.post(
                    "https://api.canva.com/rest/v1/autofills", headers=headers, json=payload, timeout=30
                )
                create_resp.raise_for_status()
                job_id = create_resp.json()["job"]["id"]

                for _ in range(20):
                    time.sleep(2)
                    status_resp = requests.get(
                        f"https://api.canva.com/rest/v1/autofills/{job_id}", headers=headers, timeout=30
                    )
                    status_resp.raise_for_status()
                    job = status_resp.json()["job"]
                    if job["status"] == "success":
                        design_url = job["result"]["design"]["url"]
                        return design_url, None
                    if job["status"] == "failed":
                        return None, f"فشل إنشاء التصميم عبر Canva: {job.get('error')}"

                return None, "استغرق إنشاء تصميم Canva وقتاً أطول من المتوقع، يرجى المحاولة لاحقاً."
            except Exception as e:
                # === إصلاح أمني: لا نعرض نص الاستثناء الخام (قد يحوي أجزاء من رابط الطلب
                # أو استجابة الخادم) للمعلم — رسالة عامة فقط، والتفصيل الكامل بسجلات الخادم. ===
                log_internal_error("فشل الاتصال بواجهة Canva", e)
                return None, "تعذّر الاتصال بواجهة Canva حالياً. يرجى المحاولة لاحقاً."

        # =====================================================================================
        # === دعم كامل للنص العربي داخل PDF عبر تشكيل الحروف (arabic_reshaper) وترتيب
        # الاتجاه (python-bidi) بالإضافة إلى خط Unicode عربي حقيقي بدل latin-1 القديم الذي
        # كان يحذف كل الحروف العربية بصمت. إن لم يتوفر خط عربي في مجلد المشروع، يُنبَّه
        # المستخدم بوضوح بدل إخراج ملف فارغ من المحتوى العربي دون علمه.
        #
        # === ملاحظة أداء مهمة ===
        # هذه الدالة مخزّنة عبر st.cache_resource فتُنفَّذ مرة واحدة فقط طوال عمر التطبيق.
        # لكن أول تشغيل بعد كل إعادة نشر (redeploy) سيبقى بطيئاً لأنه يثبّت مكتبات وينزّل
        # خطاً من الإنترنت وقت التشغيل. للحصول على أسرع أداء ممكن من أول ثانية، يُنصح
        # بإضافة 'arabic-reshaper' و 'python-bidi' إلى requirements.txt، ورفع ملف خط عربي
        # (مثل Amiri-Regular.ttf) مباشرة داخل مجلد المشروع بدل الاعتماد على هذا التنزيل
        # التلقائي — عندها ستُستخدم النسخة المحلية فوراً بدون أي تأخير شبكي.
        # =====================================================================================
        ARABIC_FONT_CANDIDATES = [
            "Amiri-Regular.ttf",
            "NotoNaskhArabic-Regular.ttf",
            "Cairo-Regular.ttf",
            "Tajawal-Regular.ttf",
        ]

        # رابط خط عربي عام (مفتوح المصدر) من مستودع Google Fonts الرسمي، يُستخدم للتنزيل
        # التلقائي وقت التشغيل فقط إذا لم يوجد أي خط عربي محلي في مجلد المشروع.
        _FALLBACK_ARABIC_FONT_URL = "https://raw.githubusercontent.com/google/fonts/main/ofl/amiri/Amiri-Regular.ttf"

        def _find_local_arabic_font():
            for fname in ARABIC_FONT_CANDIDATES:
                if os.path.exists(fname):
                    return fname
            return None

        @st.cache_resource(show_spinner=False)
        def _ensure_arabic_pdf_support():
            """
            يضمن توفر (1) مكتبتي تشكيل النص العربي و(2) خط عربي صالح، تلقائياً وقت التشغيل،
            حتى لو نسي المستخدم إضافتهما إلى requirements.txt أو رفع ملف خط. يُنفَّذ مرة واحدة
            فقط طوال عمر التطبيق بفضل st.cache_resource (لا يتكرر التنزيل/التثبيت في كل rerun).
            يعيد: (shaping_ready: bool, font_path: str|None, reshape_func, display_func, logs: list[str])
            """
            logs = []
            shaping_ready = ARABIC_SHAPING_AVAILABLE
            reshape_func = None
            display_func = None

            if shaping_ready:
                reshape_func = arabic_reshaper.reshape
                display_func = get_display
            else:
                # === إصلاح أمني/استقرار: كان الكود القديم يشغّل "pip install" تلقائياً
                # استجابةً لتفاعل المستخدم (توليد PDF) — تثبيت حزم وقت التشغيل داخل عملية
                # ستريمليت الحيّة ممارسة غير مستقرة وغير آمنة على الإنتاج (قد تفشل بصمت،
                # تبطئ الطلب بشكل غير متوقع، أو تتعارض مع بيئة الاستضافة). الحل الصحيح هو
                # إضافة المكتبتين إلى requirements.txt مسبقاً بدل تثبيتهما ديناميكياً هنا. ===
                logs.append(
                    "⚠️ مكتبتا دعم العربية (arabic-reshaper و python-bidi) غير مثبّتتين على الخادم. "
                    "الرجاء إضافة السطرين 'arabic-reshaper' و 'python-bidi' إلى requirements.txt "
                    "ثم إعادة تشغيل التطبيق (Reboot app) — سيتم توليد PDF مبسّط بالإنجليزية والأرقام فقط حتى ذلك الحين."
                )

            font_path = _find_local_arabic_font()
            if not font_path:
                try:
                    cache_path = os.path.join(tempfile.gettempdir(), "AutoArabicFont.ttf")
                    if not os.path.exists(cache_path):
                        if REQUESTS_AVAILABLE:
                            resp = requests.get(_FALLBACK_ARABIC_FONT_URL, timeout=15)
                            resp.raise_for_status()
                            with open(cache_path, "wb") as f:
                                f.write(resp.content)
                        else:
                            raise RuntimeError("مكتبة requests غير متوفرة لتنزيل الخط تلقائياً.")
                    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
                        font_path = cache_path
                        logs.append("تم تنزيل خط عربي (Amiri) تلقائياً من الإنترنت لدعم PDF بالعربية.")
                except Exception as e:
                    log_internal_error("تعذّر تنزيل خط عربي تلقائياً", e)
                    logs.append("تعذّر تنزيل خط عربي تلقائياً. سيتم توليد PDF مبسّط بالإنجليزية والأرقام فقط.")

            return shaping_ready, font_path, reshape_func, display_func, logs

        def create_pdf_file(text):
            """
            غلاف آمن: يستدعي _create_pdf_file_impl ويلتقط أي استثناء غير متوقع بدل ترك
            التطبيق ينهار أو يختفي زر التحميل بصمت — يعيد دائماً رسالة سبب واضحة عند الفشل.
            """
            if not PDF_AVAILABLE:
                return None, "⚠️ مكتبة PDF (fpdf2) غير مثبتة على الخادم. أضف السطر 'fpdf2' إلى requirements.txt ثم أعد تشغيل التطبيق (Reboot app)."
            try:
                return _create_pdf_file_impl(text)
            except Exception as e:
                log_internal_error("فشل توليد ملف PDF", e)
                return None, "⚠️ حدث خطأ أثناء توليد ملف PDF. يرجى المحاولة مرة أخرى أو التواصل مع الدعم إن تكرر الخطأ."

        def _create_pdf_file_impl(text):
            """
            يعيد tuple: (BytesIO أو None, رسالة تحذير/معلومة أو None).
            يحاول أولاً ضمان دعم العربية تلقائياً (تثبيت مكتبات + تنزيل خط) قبل التوليد.
            - إن نجح الإصلاح التلقائي أو كان الدعم متوفراً أصلاً: PDF عربي كامل وصحيح بصرياً.
            - إن فشل الإصلاح التلقائي: PDF مبسّط بالإنجليزية/الأرقام فقط مع تنبيه صريح بالسبب.
            - كل سطر يُعالَج بشكل مستقل: لو سطر معين تسبب بخطأ داخلي في fpdf2 (مشكلة معروفة
              مع بعض حروف التشكيل العربية غير المعرَّفة في الخط)، نحاول بدائل أبسط لذلك السطر
              تحديداً بدل أن ينهار الملف بالكامل ويخسر كل المحتوى السابق الذي نجح.
            """
            shaping_ready, font_path, reshape_func, display_func, setup_logs = _ensure_arabic_pdf_support()

            def _shape(line):
                return display_func(reshape_func(line))

            def _write_line_safely(pdf_obj, raw_line, is_title=False):
                """
                يحاول كتابة السطر بثلاث محاولات متدرجة: (1) نص عربي مُشكَّل بالكامل،
                (2) نص عربي خام بدون تشكيل (حروف منفصلة لكن مقروءة)، (3) نص مبسّط بالحروف
                اللاتينية فقط كحل أخير. يعيد True لو نجحت أي محاولة، وإلا False.
                قبل كل محاولة نُعيد ضبط المؤشر الأفقي لبداية الهامش الأيسر صراحة، لأن فشل
                محاولة سابقة قد يترك مؤشر fpdf2 في موضع غير صالح يُفسد كل الأسطر التالية.
                """
                height = 10 if is_title else 8
                align = "C" if is_title else "R"

                def _attempt(content):
                    pdf_obj.set_x(pdf_obj.l_margin)
                    pdf_obj.multi_cell(0, height, txt=content, align=align)

                attempts = [
                    lambda: _attempt(_shape(raw_line)),
                    lambda: _attempt(raw_line),
                    lambda: _attempt(raw_line.encode('latin-1', 'ignore').decode('latin-1') or "-"),
                ]
                for attempt in attempts:
                    try:
                        attempt()
                        return True
                    except Exception:
                        pdf_obj.set_x(pdf_obj.l_margin)
                        continue
                return False

            pdf = FPDF()
            pdf.add_page()

            if font_path and shaping_ready:
                pdf.add_font("ArabicFont", "", font_path, uni=True)
                pdf.set_font("ArabicFont", size=13)

                _write_line_safely(pdf, "ورقة العمل المكيّفة - نظام Edu Worksheet Adapt", is_title=True)
                pdf.ln(4)

                skipped_lines = 0
                for line in text.split('\n'):
                    if line.strip():
                        ok = _write_line_safely(pdf, line.strip(), is_title=False)
                        if not ok:
                            skipped_lines += 1
                    else:
                        pdf.ln(4)

                pdf_output = pdf.output()
                if isinstance(pdf_output, str):
                    pdf_output = pdf_output.encode('latin-1')
                if skipped_lines > 0:
                    return io.BytesIO(pdf_output), f"⚠️ تم إنشاء PDF بنجاح، لكن {skipped_lines} سطر تعذّر عرضه بسبب مشكلة توافق في خط العرض العربي وتم تخطيه."
                return io.BytesIO(pdf_output), None

            else:
                missing_parts = []
                if not font_path:
                    missing_parts.append("خط عربي (حاولنا تنزيله تلقائياً من الإنترنت ولم ننجح)")
                if not shaping_ready:
                    missing_parts.append("مكتبات تشكيل النص العربي (حاولنا تثبيتها تلقائياً ولم ننجح)")
                warning = (
                    "⚠️ تعذّر إنتاج PDF بالعربية رغم محاولة الإصلاح التلقائي، بسبب: "
                    + " و".join(missing_parts)
                    + ". التفاصيل: " + " | ".join(setup_logs) if setup_logs else
                    "⚠️ تعذّر إنتاج PDF بالعربية رغم محاولة الإصلاح التلقائي."
                )
                pdf.set_font("Arial", size=11)
                pdf.multi_cell(0, 10, txt="Adapted Educational Worksheet - Special Ed System", align="C")
                pdf.ln(5)
                for line in text.split('\n'):
                    clean_line = line.encode('latin-1', 'ignore').decode('latin-1')
                    if clean_line.strip():
                        pdf.multi_cell(0, 8, txt=clean_line)
                    else:
                        pdf.ln(4)
                pdf_output = pdf.output()
                if isinstance(pdf_output, str):
                    pdf_output = pdf_output.encode('latin-1')
                return io.BytesIO(pdf_output), warning

        # =====================================================================================
        # === تصدير نموذج تصحيح تلقائي كملف Excel — يحتوي على السؤال والإجابة النموذجية
        # المستخرجة من نفس استجابة الذكاء الاصطناعي، لتوفير وقت التصحيح. ===
        # =====================================================================================
        def create_answer_key_excel(answer_key_list):
            """غلاف آمن: نفس مبدأ الحماية المطبّق على PDF/Word/PowerPoint."""
            if not OPENPYXL_AVAILABLE or not answer_key_list:
                return None
            try:
                return _create_answer_key_excel_impl(answer_key_list)
            except Exception as e:
                log_internal_error("فشل توليد ملف Excel لنموذج التصحيح", e)
                return None

        def _create_answer_key_excel_impl(answer_key_list):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "نموذج التصحيح"
            try:
                ws.sheet_view.rightToLeft = True
            except Exception:
                pass
            ws["A1"] = "#"
            ws["B1"] = "السؤال"
            ws["C1"] = "الإجابة الصحيحة"
            for cell_ref in ["A1", "B1", "C1"]:
                ws[cell_ref].font = Font(bold=True)
                ws[cell_ref].alignment = Alignment(horizontal="center")
            for i, item in enumerate(answer_key_list, start=1):
                ws.cell(row=i + 1, column=1, value=i)
                c_q = ws.cell(row=i + 1, column=2, value=str(item.get("q", "")))
                c_q.alignment = Alignment(horizontal="right", wrap_text=True)
                c_a = ws.cell(row=i + 1, column=3, value=str(item.get("a", "")))
                c_a.alignment = Alignment(horizontal="right", wrap_text=True)
            ws.column_dimensions["A"].width = 6
            ws.column_dimensions["B"].width = 60
            ws.column_dimensions["C"].width = 40
            bio = io.BytesIO()
            wb.save(bio)
            bio.seek(0)
            return bio

        # =====================================================================================
        # === بطاقات PECS بصرية (صورة + كلمة) لأهم مفردات ورقة العمل، تُستخدم كدعم تواصل
        # بصري إضافي حقيقي لطلاب اضطراب طيف التوحد وصعوبات التواصل، بالاعتماد على نفس بنية
        # توليد الصور (_generate_ai_illustration) المستخدمة أصلاً في تصميم شرائح PowerPoint،
        # مع رجوع فوري لأيقونة مرسومة محلياً إذا كان توليد صور الذكاء الاصطناعي مطفأً. ===
        # =====================================================================================
        def create_pecs_cards_pdf(vocab_words):
            """غلاف آمن: نفس مبدأ الحماية المطبّق على PDF/Word/PowerPoint/Excel."""
            if not PDF_AVAILABLE or not vocab_words:
                return None, "لا توجد مفردات كافية أو مكتبة PDF غير متوفرة لإنشاء البطاقات."
            try:
                return _create_pecs_cards_pdf_impl(vocab_words)
            except Exception as e:
                log_internal_error("فشل توليد بطاقات PECS", e)
                return None, "⚠️ حدث خطأ أثناء توليد بطاقات PECS. يرجى المحاولة مرة أخرى."

        def _create_pecs_cards_pdf_impl(vocab_words):
            shaping_ready, font_path, reshape_func, display_func, _logs = _ensure_arabic_pdf_support()

            pdf = FPDF(orientation="P", unit="mm", format="A4")
            pdf.set_auto_page_break(auto=False)
            if font_path and shaping_ready:
                pdf.add_font("ArabicFont", "", font_path, uni=True)

            card_w, card_h = 85, 85
            margin_x, margin_y = 15, 15
            gap = 10
            positions = [
                (margin_x, margin_y), (margin_x + card_w + gap, margin_y),
                (margin_x, margin_y + card_h + gap), (margin_x + card_w + gap, margin_y + card_h + gap),
            ]

            icon_cycle = ["idea", "check", "star", "book", "target", "pencil"]
            words_to_use = vocab_words[:6]
            for i, word in enumerate(words_to_use):
                pos_idx = i % 4
                if pos_idx == 0:
                    pdf.add_page()
                x, y = positions[pos_idx]
                pdf.rect(x, y, card_w, card_h)

                img_bio = _generate_ai_illustration(word)
                if not img_bio:
                    # بديل فوري: أيقونة مرسومة محلياً بدل صورة الذكاء الاصطناعي (أسرع بكثير)
                    icon_kind = icon_cycle[i % len(icon_cycle)]
                    img_bio = _draw_icon(icon_kind, size=300,
                                          fg=(0x10, 0x1B, 0x2D, 255), bg=(0xF1, 0xC4, 0x0F, 255))

                tmp_img_path = os.path.join(tempfile.gettempdir(), f"pecs_{i}.png")
                with open(tmp_img_path, "wb") as f:
                    f.write(img_bio.getvalue())
                try:
                    pdf.image(tmp_img_path, x=x + 5, y=y + 5, w=card_w - 10, h=card_h - 25)
                except Exception:
                    pass

                pdf.set_xy(x, y + card_h - 18)
                if font_path and shaping_ready:
                    pdf.set_font("ArabicFont", size=13)
                    label = display_func(reshape_func(word))
                else:
                    pdf.set_font("Arial", size=11)
                    label = word.encode('latin-1', 'ignore').decode('latin-1')
                try:
                    pdf.multi_cell(card_w, 8, txt=label, align="C")
                except Exception:
                    pass

            pdf_output = pdf.output()
            if isinstance(pdf_output, str):
                pdf_output = pdf_output.encode('latin-1')
            return io.BytesIO(pdf_output), None

        # =====================================================================================
        # === توليد تقرير متابعة PDF من صفحة واحدة لطالب معيّن، يلخّص تاريخ أوراق العمل التي
        # كُيّفت له (يُستخدم من قسم "تقرير متابعة الطالب" أعلاه). ===
        # =====================================================================================
        def create_progress_report_pdf(student_name, history_rows_list):
            if not PDF_AVAILABLE:
                return None, "مكتبة PDF غير متوفرة."
            shaping_ready, font_path, reshape_func, display_func, _logs = _ensure_arabic_pdf_support()

            pdf = FPDF()
            pdf.add_page()
            use_arabic = bool(font_path and shaping_ready)
            if use_arabic:
                pdf.add_font("ArabicFont", "", font_path, uni=True)
                pdf.set_font("ArabicFont", size=16)
                title = display_func(reshape_func(f"تقرير متابعة الطالب: {student_name}"))
            else:
                pdf.set_font("Arial", size=14)
                title = f"Progress Report: {student_name}"
            pdf.multi_cell(0, 10, txt=title, align="C")
            pdf.ln(4)

            for row in history_rows_list:
                line = f"{row['created_at'][:10]}  |  {row['subject']}  |  {row['adaptation_level']}  |  {row['mode']}"
                pdf.set_font("ArabicFont" if use_arabic else "Arial", size=11)
                content = display_func(reshape_func(line)) if use_arabic else line.encode('latin-1', 'ignore').decode('latin-1')
                try:
                    pdf.multi_cell(0, 8, txt=content, align="R")
                except Exception:
                    pass
                pdf.ln(1)

            pdf_output = pdf.output()
            if isinstance(pdf_output, str):
                pdf_output = pdf_output.encode('latin-1')
            return io.BytesIO(pdf_output), None

        # ملاحظة: parse_ai_sections أصبحت دالة على مستوى الملف (أعلاه) بدل تعريفها هنا،
        # ليمكن اختبارها آلياً — راجع tests/test_worksheet_helpers.py.

        # =====================================================================================
        # === إصلاح (٣): معالجة قوية لخطأ "503 UNAVAILABLE / High Demand" عند توليد الامتحان
        # التقييمي وبنك الإجابات الاحتياطي. كان الكود القديم يستدعي نموذجاً واحداً فقط
        # (gemini-3.1-flash-lite) بدون أي إعادة محاولة، فأي ضغط مؤقت على خادم جوجل كان
        # يُفشل العملية فوراً برسالة خطأ تقنية غير مفهومة للمعلم ("This model is currently
        # experiencing high demand..."). الآن نجرّب عدة نماذج بديلة بالتتابع، ولكل نموذج
        # نعيد المحاولة عدة مرات مع فترة انتظار متصاعدة (Exponential Backoff) عند رصد
        # خطأ 503/UNAVAILABLE/high demand تحديداً — بدل الاستسلام من أول محاولة واحدة. ===
        # =====================================================================================
        EXAM_MODELS_TO_TRY = ["gemini-3.1-flash-lite", "gemini-3.5-flash-lite", "gemini-2.5-flash"]

        def _call_model_with_retry(prompt_text, models_list, temperature=0.6, max_output_tokens=3000, max_retries_per_model=2):
            """
            يحاول توليد المحتوى عبر أكثر من نموذج بالتتابع، ولكل نموذج يُعاد المحاولة عدة
            مرات مع فترة انتظار متصاعدة عند مواجهة خطأ 503/ضغط مرتفع على الخادم تحديداً،
            بدل الفشل من أول محاولة أو الاعتماد على نموذج واحد فقط.
            يعيد tuple: (النص الناتج أو None, آخر خطأ حدث أو None).
            """
            last_error = None
            for model_name in models_list:
                for attempt in range(max_retries_per_model):
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=prompt_text,
                            config=types.GenerateContentConfig(
                                temperature=temperature,
                                max_output_tokens=max_output_tokens,
                            ),
                        )
                        if response and response.text and response.text.strip():
                            return response.text.strip(), None
                    except Exception as e:
                        last_error = e
                        err_str = str(e)
                        if ("503" in err_str or "UNAVAILABLE" in err_str.upper()
                                or "overloaded" in err_str.lower() or "high demand" in err_str.lower()):
                            time.sleep((2 * (attempt + 1)) + random.uniform(0, 1))
                            continue
                        else:
                            # خطأ غير مرتبط بالضغط (مثل مفتاح API غير صالح) — لا داعي لإعادة نفس النموذج
                            break
            return None, last_error

        # ملاحظة: _format_exam_text_for_display أصبحت دالة على مستوى الملف (أعلاه)
        # بدل تعريفها هنا، ليمكن اختبارها آلياً — راجع tests/test_worksheet_helpers.py.

        # =====================================================================================
        # === إصلاح: استدعاء احتياطي منفصل لاستخراج بنك الإجابات والمفردات ===
        # المشكلة التي كانت تظهر ("نموذج التصحيح غير متوفر" و"بطاقات PECS غير متوفرة")
        # سببها أن قسمي ### ANSWER_KEY_JSON ### و### KEY_VOCAB ### كانا يُطلبان في
        # نهاية نفس الاستدعاء الرئيسي الذي يولّد ورقة العمل كاملة. عندما تكون ورقة العمل
        # طويلة، كان النموذج يستهلك حد الأسطر المسموح (max_output_tokens) بالكامل في
        # كتابة الورقة نفسها، فينقطع الرد قبل أن يصل لكتابة هذين القسمين إطلاقاً — والنتيجة
        # بنك إجابات وقائمة مفردات فارغة دائماً مع الأوراق الطويلة.
        # الحل: إن جاء الرد الرئيسي بدون هذين القسمين (أو أحدهما)، نطلبهما الآن باستدعاء
        # ثانٍ صغير ومستقل مخصص لهذه المهمة فقط، بحد أسطر كافٍ خاص بها لا يتأثر بطول
        # ورقة العمل نفسها، ويستخدم الآن نفس آلية إعادة المحاولة وتبديل النماذج أعلاه. ===
        # =====================================================================================
        def _generate_structured_extras(main_text_for_extraction):
            """
            استدعاء احتياطي خفيف يطلب فقط بنك الإجابات وقائمة المفردات من نص ورقة عمل
            جاهز بالفعل. يعيد tuple: (answer_key: list, vocab_words: list) — قوائم فارغة
            عند أي فشل بدل رفع استثناء يوقف بقية سير العمل.
            """
            extras_prompt = f"""
                بناءً على ورقة العمل التالية، نفّذ المطلوبين التاليين فقط بدون أي نص إضافي
                قبلهما أو بعدهما أو بينهما، والتزم بالتنسيق حرفياً:

                ### ANSWER_KEY_JSON ###
                [{{"q": "نص مختصر للسؤال", "a": "الإجابة النموذجية الصحيحة"}}]
                (عنصر واحد لكل سؤال تقييمي فعلي ورد في ورقة العمل أدناه، وبدون أي ```)

                ### KEY_VOCAB ###
                من ٤ إلى ٦ كلمات مفتاحية أساسية من محتوى ورقة العمل، مفصولة بفواصل فقط.

                ورقة العمل:
                {main_text_for_extraction[:6000]}
            """
            extras_text, _extras_error = _call_model_with_retry(
                extras_prompt, EXAM_MODELS_TO_TRY, temperature=0.3, max_output_tokens=2000
            )
            if extras_text:
                _, extracted_answer_key, extracted_vocab_words = parse_ai_sections(extras_text)
                return extracted_answer_key, extracted_vocab_words
            return [], []

        # =====================================================================================
        # === توليد ملفات التحميل (Word/PPT/PDF) مرة واحدة فقط لكل نص مُكيَّف، بدل إعادة
        # توليدها في كل rerun من ستريمليت. ===
        # =====================================================================================
        if "adapted_text" not in st.session_state:
            st.session_state.adapted_text = None
        if "adapted_text_draft" not in st.session_state:
            st.session_state.adapted_text_draft = None
        if "answer_key" not in st.session_state:
            st.session_state.answer_key = []
        if "vocab_words" not in st.session_state:
            st.session_state.vocab_words = []
        if "just_generated" not in st.session_state:
            st.session_state.just_generated = False
        if "generated_files" not in st.session_state:
            st.session_state.generated_files = {}
        if "generated_for_text" not in st.session_state:
            st.session_state.generated_for_text = None
        if "history_saved_for" not in st.session_state:
            st.session_state.history_saved_for = None
        if "audio_generated_for_text" not in st.session_state:
            st.session_state.audio_generated_for_text = None
        if "generated_audio" not in st.session_state:
            st.session_state.generated_audio = None
        if "audio_warning" not in st.session_state:
            st.session_state.audio_warning = None

        try:
            start_btn_container = st.container(key="start-ai-button")
        except TypeError:
            start_btn_container = st.container()
        with start_btn_container:
            start_clicked = st.button("🚀 Start", use_container_width=True)

        # === سقف الاستخدام الشهري المرتبط بالاشتراك على Whop: يُطبّق فقط على الحسابات
        # المرتبطة بمفتاح ترخيص (الحسابات القديمة بدون مفتاح تبقى بلا قيود). ===
        _quota_blocked_now = False
        if start_clicked:
            _current_teacher_for_quota = st.session_state.get("teacher_name", "")
            if _current_teacher_for_quota:
                _current_month_usage = get_monthly_usage_count(_current_teacher_for_quota)
                if _current_month_usage >= MONTHLY_WORKSHEET_LIMIT:
                    _quota_blocked_now = True
                    st.error(
                        f"⚠️ لقد استخدمت الحد الأقصى المسموح به هذا الشهر "
                        f"({MONTHLY_WORKSHEET_LIMIT} ورقة عمل) ضمن اشتراكك الحالي. "
                        "سيُعاد تعيين الحد أول الشهر القادم، أو تواصل معنا لترقية الباقة."
                    )

        if start_clicked and not _quota_blocked_now:
            # إعادة تصفير دورة العمل بالكامل عند بدء تكييف جديد
            st.session_state.adapted_text = None
            st.session_state.adapted_text_draft = None
            st.session_state.answer_key = []
            st.session_state.vocab_words = []
            st.session_state.generated_files = {}
            st.session_state.generated_for_text = None
            st.session_state.history_saved_for = None
            st.session_state.audio_generated_for_text = None
            st.session_state.generated_audio = None
            st.session_state.audio_warning = None

            if not extracted_content.strip():
                extracted_content = f"ورقة عمل عامة لمبحث {selected_subject} للصف {selected_grade} وفق النظام {selected_system}."

            trimmed_content = extracted_content[:MAX_INPUT_CHARS] if len(extracted_content) > MAX_INPUT_CHARS else extracted_content
            template_hint = ADAPTATION_TEMPLATE_HINTS.get(selected_category, "")
            formatting_instructions = FORMATTING_INSTRUCTIONS

            structured_output_instructions = """
                بعد الانتهاء من كتابة ورقة العمل كاملة، أضف بالضبط القسمين التاليين في النهاية
                (لا تكتب أي نص بعدهما، والتزم بالتنسيق حرفياً، ولا تضع علامات ```
                حول الـ JSON إطلاقاً — اكتبه كسطر عادي فقط):

                ### ANSWER_KEY_JSON ###
                [{"q": "نص مختصر للسؤال", "a": "الإجابة النموذجية الصحيحة"}]
                (اكتب عنصراً واحداً داخل القائمة لكل سؤال تقييمي فعلي ورد في الورقة، وبدون أي ```)

                ### KEY_VOCAB ###
                اكتب هنا فقط ٤ إلى ٦ كلمات مفتاحية أساسية من محتوى الورقة، مفصولة بفواصل، بدون أي شرح إضافي.
            """

            if generate_alternative:
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن.
                مطلوب تصميم ورقة عمل بديلة مقترحة بالكامل مع **بنك أسئلة تقييمي تشخيصي مفصل يتضمن الأسئلة والحلول النموذجية** يناسب الحالة الخاصة ({selected_condition}) ومستوى التكييف ({selected_level}).

                إرشاد تكييف معتمد لهذه الفئة (استخدمه كأساس أسلوبي): {template_hint}

                البيانات الأساسية:
                - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن

                محتوى الملف المرفق:
                {trimmed_content}

                اكتب ورقة العمل والأسئلة والتمارين والحلول بخطوات تفصيلية كاملة وواضحة باللغة العربية.

                {formatting_instructions}

                {structured_output_instructions}
                """
            else:
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن.
                مطلوب تنفيذ **تكييف وتطوير شامل ودقيق** لورقة العمل التالية لمبحث ({selected_subject}) بناءً على مستوى التكييف ({selected_level}) والحالة الخاصة ({selected_condition}).

                إرشاد تكييف معتمد لهذه الفئة (استخدمه كأساس أسلوبي): {template_hint}

                البيانات الأساسية:
                - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن

                محتوى الملف المرفق:
                {trimmed_content}

                قم بإعادة صياغة ورقة العمل وكتابة الأسئلة المعدلة، التمارين التدريبية، والحلول بشكل كامل ووافٍ دون أي نقصان وبأسلوب تربوي متميز.

                {formatting_instructions}

                {structured_output_instructions}
                """

            # =================================================================================
            # === إصلاح أداء (سرعة مُدركة): توليد الرد بشكل متدفّق (Streaming) بدل انتظار
            # النص كاملاً خلف سبينر صامت. النص يبدأ بالظهور على الشاشة فور وصول أول جزء منه
            # من الذكاء الاصطناعي، فيشعر المستخدم أن التطبيق يعمل ويستجيب فوراً بدل الشك
            # بأنه "علّق" لثوانٍ طويلة — وهذا فرق حقيقي في تجربة الاستخدام حتى لو ظل زمن
            # التوليد الكلي نفسه تقريباً. ===
            # =================================================================================
            adapted_text_raw = None
            models_to_try = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash"]
            last_error = None
            stream_placeholder = st.empty()
            for model_name in models_to_try:
                try:
                    full_text = ""
                    stream = client.models.generate_content_stream(
                        model=model_name,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            temperature=0.7,
                            max_output_tokens=8000,
                        ),
                    )
                    for chunk in stream:
                        chunk_text = getattr(chunk, "text", None)
                        if chunk_text:
                            full_text += chunk_text
                            # نعرض النص أولاً بأول أثناء وصوله، مع إخفاء أي جزء من قسمي
                            # الإجابات/المفردات إن بدأ الظهور قبل انتهاء البث بالكامل
                            visible_text = full_text.split("### ANSWER_KEY_JSON ###")[0]
                            stream_placeholder.markdown(visible_text)
                    if full_text.strip():
                        adapted_text_raw = full_text
                        break
                except Exception as e:
                    last_error = e
                    time.sleep(1)
                    continue
            stream_placeholder.empty()

            if adapted_text_raw:
                main_text, answer_key, vocab_words = parse_ai_sections(adapted_text_raw)

                # --- إصلاح: لو الرد الرئيسي انقطع قبل أن يكتب بنك الإجابات أو المفردات
                # (شائع في الأوراق الطويلة بسبب حد max_output_tokens)، نطلبهما الآن
                # باستدعاء ثانٍ صغير مخصص بدل ترك الأدوات الإضافية فارغة دائماً. ---
                if not answer_key or not vocab_words:
                    with st.spinner("جاري استكمال بنك الإجابات والمفردات..."):
                        fallback_answer_key, fallback_vocab_words = _generate_structured_extras(main_text)
                    if not answer_key:
                        answer_key = fallback_answer_key
                    if not vocab_words:
                        vocab_words = fallback_vocab_words

                st.session_state.adapted_text_draft = main_text
                st.session_state.answer_key = answer_key
                st.session_state.vocab_words = vocab_words
                st.session_state.just_generated = True
                st.success("تم تكييف ورقة العمل بنجاح تام / Adapted Successfully! راجعها أدناه قبل التصدير.")
            else:
                st.error("عذراً، تعذّر الاتصال بخدمة الذكاء الاصطناعي حالياً. يرجى المحاولة لاحقاً، أو التأكد من صلاحية مفتاح GOOGLE_API_KEY.")
                if last_error:
                    # === إصلاح أمني: التفصيل التقني الكامل (قد يحوي أجزاء من الطلب المرسل
                    # لجوجل) يُطبع بسجلات الخادم فقط، ولا يُعرض للمعلم كما كان سابقاً. ===
                    log_internal_error("فشل توليد ورقة العمل عبر كل النماذج المتاحة", last_error)

        # =====================================================================================
        # === خطوة مراجعة وتعديل يدوي قبل التصدير النهائي — النص لا يذهب مباشرة لتوليد
        # الملفات، بل يُعرض كمسودة قابلة للتعديل، ولا تُنشأ ملفات Word/PPT/PDF إلا بعد
        # ضغط المعلم على "اعتماد وتصدير". ===
        # =====================================================================================
        if st.session_state.adapted_text_draft and not st.session_state.adapted_text:
            st.markdown("### ✏️ مراجعة وتعديل الورقة قبل التصدير النهائي / Review & Edit Before Export")
            st.info("عدّل النص مباشرة إذا رغبت، ثم اضغط 'اعتماد وتصدير' لإنشاء ملفات التحميل النهائية.")
            edited_text = st.text_area(
                "محتوى ورقة العمل / Worksheet Content:",
                value=st.session_state.adapted_text_draft,
                height=420,
                key="review_textarea"
            )
            col_approve, col_discard = st.columns(2)
            with col_approve:
                if st.button("✅ اعتماد وتصدير / Approve & Export", use_container_width=True, key="approve_export_btn"):
                    st.session_state.adapted_text = edited_text
                    st.rerun()
            with col_discard:
                if st.button("🔄 إلغاء والبدء من جديد / Discard & Restart", use_container_width=True, key="discard_btn"):
                    st.session_state.adapted_text_draft = None
                    st.session_state.answer_key = []
                    st.session_state.vocab_words = []
                    st.rerun()

        if st.session_state.adapted_text:

            if st.session_state.just_generated:
                if enable_ding:
                    play_ready_ding()
                st.session_state.just_generated = False

            st.markdown("### ورقة العمل المطورة والمكيفة / Adapted Worksheet Output:")
            # === جديد: عرض الورقة داخل بطاقة منسّقة (تباعد أسطر مريح، عناوين وخط عريض
            # وفواصل مميّزة بصرياً) بدل نص عادٍ متلاصق — التنسيق الفعلي (# و** و---) يأتي
            # من تعليمات الذكاء الاصطناعي أعلاه؛ هذه البطاقة فقط تُخرجه بشكل مريح للعين. ===
            try:
                worksheet_output_card = st.container(key="worksheet-output-card")
            except TypeError:
                worksheet_output_card = st.container()
            with worksheet_output_card:
                st.markdown(st.session_state.adapted_text)

            current_text = st.session_state.adapted_text

            # =====================================================================================
            # === "🔊 استمع للورقة" — يدعم الطلاب ذوي الإعاقة البصرية أو صعوبات القراءة الذين
            # يحتاجون سماع محتوى الورقة بدل قراءته. يُولَّد الصوت مرة واحدة فقط لكل نص معتمد
            # (مخزَّن بالجلسة) بدل إعادة الاتصال بخدمة الصوت في كل rerun. ===
            # =====================================================================================
            with st.expander("🔊 استمع للورقة / Listen to the Worksheet (لذوي الإعاقة البصرية)"):
                if not GTTS_AVAILABLE:
                    st.info("⚠️ ميزة الاستماع الصوتي غير مفعّلة على الخادم حالياً. أضف السطر 'gTTS' إلى requirements.txt ثم أعد تشغيل التطبيق (Reboot app).")
                else:
                    if st.button("🎧 تحويل الورقة إلى صوت / Generate Audio", key="generate_audio_btn"):
                        with st.spinner("جاري تحويل الورقة إلى صوت..."):
                            audio_bio, audio_warning = create_worksheet_audio(current_text)
                        st.session_state.generated_audio = audio_bio.getvalue() if audio_bio else None
                        st.session_state.audio_warning = audio_warning
                        st.session_state.audio_generated_for_text = current_text

                    if st.session_state.audio_generated_for_text == current_text and st.session_state.generated_audio:
                        st.audio(st.session_state.generated_audio, format="audio/mp3")
                        st.download_button(
                            label="⬇️ تحميل الملف الصوتي (.mp3)",
                            data=st.session_state.generated_audio,
                            file_name="Adapted_Worksheet_Audio.mp3",
                            mime="audio/mp3",
                        )
                    if st.session_state.audio_warning:
                        st.info(st.session_state.audio_warning)

            # --- حفظ نسخة من هذه الورقة في سجل الطالب مرة واحدة فقط لكل نص معتمد ---
            if st.session_state.history_saved_for != current_text:
                _history_student_id = selected_student_record["id"] if selected_student_record else None
                _history_student_name = (
                    selected_student_record["full_name"] if selected_student_record
                    else (new_student_name.strip() if new_student_name.strip() else "بدون ربط بطالب")
                )
                save_worksheet_history(
                    current_teacher, _history_student_id, _history_student_name,
                    selected_subject, selected_grade, selected_level, mode_items[mode_idx][1],
                    current_text, json.dumps(st.session_state.answer_key, ensure_ascii=False)
                )
                st.session_state.history_saved_for = current_text

            st.markdown("---")

            # === جديد: معاينة نهائية للورقة تظهر مباشرة قبل أزرار التحميل، بنفس التنسيق
            # المنسّق أعلاه، حتى يراجعها المعلم آخر مرة قبل أن يحفظها بأي صيغة. ===
            st.markdown("### 📄 معاينة نهائية للورقة قبل التحميل / Final Preview Before Download")
            try:
                worksheet_preview_before_download = st.container(key="worksheet-output-card-preview")
            except TypeError:
                worksheet_preview_before_download = st.container()
            with worksheet_preview_before_download:
                st.markdown(current_text)

            st.markdown("---")
            st.subheader("📥 تحميل الملفات المطورة / Download Adapted Files:")

            # =====================================================================================
            # === إصلاح أداء: توليد الملفات الخمسة (Word/PPT/PDF/Excel/PECS) بالتوازي بدل
            # التسلسل. كانت تُبنى الواحد تلو الآخر فيتراكم الوقت (مجموع كل الملفات)؛ الآن
            # تُبنى معاً في خيوط منفصلة فيصبح الوقت الكلي مساوياً تقريباً لأبطأ ملف واحد
            # فقط بدل مجموعها كلها. الفرق يصبح كبيراً جداً خصوصاً مع تفعيل صور الذكاء
            # الاصطناعي أو عند توليد PDF لأول مرة بعد إعادة النشر. ===
            # =====================================================================================
            if st.session_state.generated_for_text != current_text:
                with st.spinner("جاري تجهيز ملفات Word و PowerPoint و PDF للتحميل (مرة واحدة فقط)..."):
                    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                        future_word = executor.submit(create_word_file, current_text)
                        future_ppt = executor.submit(create_ppt_file, current_text)
                        future_pdf = executor.submit(create_pdf_file, current_text)
                        future_excel = executor.submit(create_answer_key_excel, st.session_state.answer_key)
                        future_pecs = executor.submit(create_pecs_cards_pdf, st.session_state.vocab_words)

                        # === إصلاح جوهري: كل دالة توليد أعلاه أصبحت الآن محمية داخلياً بـ
                        # try/except خاص بها، لكن هذا الحاجز الإضافي هنا يضمن أنه حتى لو
                        # حدث خطأ غير متوقع تماماً (مثلاً خطأ داخل الحماية نفسها)، فإن
                        # future.result() لن يرفع الاستثناء ليُسقط الصفحة كاملة ويُفشل
                        # عرض بقية أزرار التحميل معه. سابقاً: أي استثناء من أي ملف واحد
                        # كان يُسقط الخمسة ملفات كلها ويُظهر خطأ عام للمعلم. ===
                        def _safe_result(future, label, default):
                            try:
                                return future.result()
                            except Exception as e:
                                log_internal_error(f"فشل غير متوقع أثناء توليد {label}", e)
                                return default

                        word_bio = _safe_result(future_word, "ملف Word", None)
                        ppt_bio = _safe_result(future_ppt, "ملف PowerPoint", None)
                        pdf_bio, pdf_warning = _safe_result(
                            future_pdf, "ملف PDF",
                            (None, "⚠️ حدث خطأ غير متوقع أثناء توليد ملف PDF.")
                        )
                        answer_key_excel_bio = _safe_result(future_excel, "ملف Excel لنموذج التصحيح", None)
                        pecs_pdf_bio, pecs_warning = _safe_result(
                            future_pecs, "بطاقات PECS",
                            (None, "⚠️ حدث خطأ غير متوقع أثناء توليد بطاقات PECS.")
                        )

                    st.session_state.generated_files = {
                        "word": word_bio.getvalue() if word_bio else None,
                        "ppt": ppt_bio.getvalue() if ppt_bio else None,
                        "pdf": pdf_bio.getvalue() if pdf_bio else None,
                        "pdf_warning": pdf_warning,
                        "answer_key_excel": answer_key_excel_bio.getvalue() if answer_key_excel_bio else None,
                        "pecs_pdf": pecs_pdf_bio.getvalue() if pecs_pdf_bio else None,
                        "pecs_warning": pecs_warning,
                    }
                    st.session_state.generated_for_text = current_text

            files = st.session_state.generated_files

            col1, col2, col3 = st.columns(3)

            with col1:
                if files.get("word") and DOCX_AVAILABLE:
                    st.download_button(
                        label="تحميل Word (.docx)",
                        data=files["word"],
                        file_name="Adapted_Worksheet.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.info("تصدير Word غير متوفر حالياً.")

            with col2:
                if files.get("ppt") and PPTX_AVAILABLE:
                    st.download_button(
                        label="تحميل PowerPoint (.pptx) 🎨",
                        data=files["ppt"],
                        file_name="Interactive_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                else:
                    st.info("تصدير PowerPoint غير متوفر حالياً.")

                if CANVA_INTEGRATION_ENABLED:
                    if st.button("🎨 إنشاء نسخة مصمّمة عبر Canva"):
                        with st.spinner("جاري إنشاء التصميم عبر Canva..."):
                            design_url, err = create_canva_design(
                                current_text,
                                f"{selected_grade} - {selected_subject}"
                            )
                        if design_url:
                            st.success("تم إنشاء التصميم بنجاح عبر Canva!")
                            st.markdown(f"[فتح التصميم في Canva]({design_url})")
                        else:
                            st.warning(err)
                else:
                    st.caption("ℹ️ لتفعيل التصميم عبر حساب Canva فعلياً، أضف CANVA_API_TOKEN و CANVA_BRAND_TEMPLATE_ID في Secrets.")

            with col3:
                if files.get("pdf") and PDF_AVAILABLE:
                    st.download_button(
                        label="تحميل PDF (.pdf)",
                        data=files["pdf"],
                        file_name="Adapted_Worksheet.pdf",
                        mime="application/pdf"
                    )
                    if files.get("pdf_warning"):
                        st.warning(files["pdf_warning"])
                else:
                    st.error(files.get("pdf_warning") or "تصدير PDF غير متوفر حالياً لسبب غير معروف.")

            st.markdown("#### 🧩 أدوات إضافية / Extra Tools:")
            col4, col5 = st.columns(2)
            with col4:
                if files.get("answer_key_excel") and OPENPYXL_AVAILABLE:
                    st.download_button(
                        label="📊 تحميل نموذج التصحيح (Excel)",
                        data=files["answer_key_excel"],
                        file_name="Answer_Key.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("نموذج التصحيح غير متوفر لهذه الورقة (لم يُستخرج بنك إجابات صالح).")
            with col5:
                if files.get("pecs_pdf"):
                    st.download_button(
                        label="🖼️ تحميل بطاقات PECS البصرية (PDF)",
                        data=files["pecs_pdf"],
                        file_name="PECS_Cards.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.info(files.get("pecs_warning") or "بطاقات PECS غير متوفرة لهذه الورقة.")

            st.markdown("---")

            # =====================================================================================
            # === جديد: تحويل الورقة إلى امتحان تقييمي — أيقونة واضحة ومكتوب جنبها، تظهر بعد
            # التكييف وقبل صندوق الشكر مباشرة، لتساعد المعلم يتأكد إن الطالب فهم المادة فعلاً
            # ووصلته المعلومة، عبر توليد امتحان قصير مبني على نفس محتوى الورقة المكيّفة. ===
            # =====================================================================================
            if "exam_text" not in st.session_state:
                st.session_state.exam_text = None
            if "exam_for_text" not in st.session_state:
                st.session_state.exam_for_text = None

            try:
                exam_section_card = st.container(key="exam-conversion-card")
            except TypeError:
                exam_section_card = st.container()

            with exam_section_card:
                st.markdown("#### 🧪📋 تحويل الورقة إلى امتحان تقييمي / Convert to Assessment Exam")
                st.caption(
                    "استخدم هذا الخيار للتأكد أن الطالب فهم المادة فعلاً ووصلته المعلومة — يولّد "
                    "الذكاء الاصطناعي امتحاناً تقييمياً قصيراً (٥-٨ أسئلة) مبنياً على نفس محتوى "
                    "الورقة أعلاه، مع نموذج إجابة لكل سؤال."
                )
                convert_to_exam_clicked = st.button(
                    "🧪 حوّل هذه الورقة إلى امتحان تقييمي / Convert to Exam",
                    key="convert_to_exam_btn",
                    use_container_width=True
                )

            if convert_to_exam_clicked:
                with st.spinner("جاري توليد الامتحان التقييمي..."):
                    exam_prompt = f"""
                        أنت خبير تربوي متخصص في التقييم التربوي والتربية الخاصة.
                        بناءً على ورقة العمل التالية (الصف: {selected_grade} | المادة: {selected_subject} |
                        الحالة الخاصة: {selected_condition} | مستوى التكييف: {selected_level}),
                        صمم امتحاناً تقييمياً قصيراً من ٥ إلى ٨ أسئلة هدفه التأكد الفعلي من أن الطالب
                        فهم المفاهيم واستوعب المعلومة، وليس تكراراً حرفياً لنفس أسئلة الورقة. نوّع بين
                        أسئلة اختيار من متعدد وأسئلة قصيرة مباشرة تناسب حالة الطالب ومستوى التكييف.

                        {FORMATTING_INSTRUCTIONS}

                        التزم أيضاً بما يلي خاص بالامتحان تحديداً:
                        - ابدأ بعنوان رئيسي واحد فقط بصيغة Markdown من المستوى الأول، مثل:
                          # امتحان تقييمي: {selected_subject}
                        - رقّم الأسئلة بالتسلسل بدءاً من "السؤال ١" ثم "السؤال ٢" وهكذا دون تخطي رقم.
                        - بعد كل سؤال وخياراته (إن وجدت)، أضف نموذج الإجابة الصحيحة في سطر مستقل
                          يبدأ حرفياً بالعبارة "الإجابة النموذجية:" ثم نص الإجابة مباشرة بعدها في
                          نفس السطر (بدون فقرة إضافية بعده).
                        - لا تكرر عبارة "الإجابة النموذجية:" أكثر من مرة لكل سؤال.

                        محتوى ورقة العمل الأصلية:
                        {current_text[:6000]}

                        اكتب الامتحان مباشرة بدون أي مقدمات أو تعليق قبله.
                    """
                    exam_text_result, exam_error = _call_model_with_retry(
                        exam_prompt, EXAM_MODELS_TO_TRY, temperature=0.6, max_output_tokens=3000
                    )
                    if exam_text_result:
                        st.session_state.exam_text = _format_exam_text_for_display(exam_text_result)
                        st.session_state.exam_for_text = current_text
                    else:
                        st.warning(
                            "تعذّر توليد الامتحان التقييمي حالياً بسبب ضغط مرتفع مؤقت على خدمة الذكاء "
                            "الاصطناعي (Google). تم تجربة عدة نماذج بديلة تلقائياً ولم تنجح أي منها. "
                            "يرجى الانتظار دقيقة واحدة ثم الضغط على الزر مرة أخرى."
                        )
                        if exam_error:
                            st.caption(f"تفاصيل تقنية: {exam_error}")

            if st.session_state.exam_text and st.session_state.exam_for_text == current_text:
                st.markdown("##### 📋 الامتحان التقييمي الناتج / Generated Assessment Exam:")
                try:
                    exam_output_card = st.container(key="exam-output-card")
                except TypeError:
                    exam_output_card = st.container()
                with exam_output_card:
                    st.markdown(st.session_state.exam_text)

                if DOCX_AVAILABLE:
                    exam_word_bio = create_word_file(st.session_state.exam_text)
                    if exam_word_bio:
                        st.download_button(
                            label="تحميل الامتحان كملف Word (.docx)",
                            data=exam_word_bio.getvalue(),
                            file_name="Assessment_Exam.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_exam_word_btn"
                        )

            st.markdown("---")
            # =====================================================================================
            # === إصلاح: الحجم السابق (بعد الضغط الشديد في المرة الماضية) أصبح صغيراً جداً
            # وكل النصوص متلاصقة في سطر واحد. الآن حجم متوسط متوازن: كل سطر (العنوان،
            # الجملة العربية، الفاصل، الجملة الإنجليزية) في سطر مستقل وواضح، بخط مقروء،
            # مع مساحة داخلية كافية — بدون الرجوع لحجم الصندوق الضخم الأصلي. ===
            # =====================================================================================
            st.markdown(f"""
                <div class="animated-box" style="
                    background: linear-gradient(135deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%);
                    border: none;
                    padding: 24px 20px;
                    border-radius: 20px;
                    text-align: center;
                    margin-top: 18px;
                    box-shadow: 0px 10px 26px rgba(16,27,45,0.30);
                ">
                    <div style="font-size: 30px; line-height: 1; margin-bottom: 8px;">🎓✨</div>
                    <div style="font-weight: 900; color: {GOLD}; font-size: 19px; line-height: 1.5;">
                        شكراً لاستخدامك Edu Worksheet Adapt
                    </div>
                    <div style="height: 1px; background: rgba(255,255,255,0.25); margin: 14px auto; width: 50%;"></div>
                    <div style="color: {WHITE}; font-size: 14px; font-weight: 700; line-height: 1.7;">
                        نحو تعليم أكثر شمولاً يليق بكل طالب 💙
                    </div>
                    <div style="margin-top: 6px; color: rgba(255,255,255,0.8); font-size: 12.5px; font-weight: 500;">
                        Thank you for using Edu Worksheet Adapt
                    </div>
                </div>
            """, unsafe_allow_html=True)


# =========================================================================================
# === اختبارات آلية (Automated Tests) — بنفس ملف app.py حتى يبقى التطبيق ملفاً واحداً
# فقط لتسهيل النشر (بدل ملف worksheet_helpers.py منفصل).
#
# للتشغيل: pytest app.py -v
#
# لماذا هذا آمن ولا يؤثر على التطبيق الفعلي إطلاقاً:
# 1) عند تشغيل "streamlit run app.py"، تُنفَّذ واجهة التطبيق كاملة لأن __name__ يساوي
#    "__main__"، لكن هذه الكتلة تبدأ بعد نهاية تلك الكتلة، وpytest غير مستورد أصلاً في
#    بيئة الإنتاج عادة — وحتى لو كان مثبتاً، دوال pytest.fixture/test_* هنا لا تُستدعى
#    من أي مكان بالتطبيق نفسه، فهي ببساطة تعريفات خاملة لا تُنفَّذ إلا عبر أمر pytest.
# 2) عند تشغيل "pytest app.py"، يستورد pytest هذا الملف كوحدة عادية (__name__ يساوي
#    اسم الوحدة وليس "__main__")، فكتلة واجهة ستريمليت أعلاه لا تُنفَّذ إطلاقاً، ويبقى
#    فقط تنفيذ طبقة قاعدة البيانات وتعريف الدوال — تماماً كما يحتاج الاستيراد للاختبار.
# =========================================================================================
try:
    import pytest as _pytest_for_tests
except ImportError:
    _pytest_for_tests = None

if _pytest_for_tests is not None:
    pytest = _pytest_for_tests

    @pytest.fixture(autouse=True)
    def _isolated_test_db(tmp_path, monkeypatch):
        """
        كل اختبار يحصل على قاعدة بيانات SQLite مؤقتة ونظيفة خاصة به (بتغيير مجلد
        العمل الحالي إلى مجلد مؤقت فريد، ثم إعادة إنشاء الجداول فيه)، حتى لا تتداخل
        الاختبارات مع بعضها أو مع أي قاعدة بيانات حقيقية للتطبيق.

        كذلك نستبدل (monkeypatch) دالة التحقق من مفتاح ترخيص Whop بنسخة وهمية تنجح
        فوراً دون أي اتصال شبكة فعلي — الاختبارات تتحقق من منطق الحسابات محلياً
        (تسجيل/دخول/قفل) لا من تكامل Whop نفسه، الذي لا يمكن اختباره بدون شبكة حقيقية.
        """
        monkeypatch.chdir(tmp_path)
        import sys as _sys_for_patch
        monkeypatch.setattr(
            _sys_for_patch.modules[__name__], "_whop_check_license",
            lambda license_key: (True, "active", None),
        )
        init_db()
        yield

    # =====================================================================================
    # === كلمات المرور: التجزئة والتحقق ===
    # =====================================================================================

    def test_password_hash_is_salted_and_not_reversible_plaintext():
        hash1 = _hash_password("Teacher123")
        hash2 = _hash_password("Teacher123")
        assert hash1 != hash2, "نفس كلمة المرور يجب أن تنتج تجزئتين مختلفتين بفضل الملح العشوائي"
        assert "Teacher123" not in hash1

    def test_password_hash_verifies_correctly():
        stored = _hash_password("Teacher123")
        is_valid, needs_upgrade = _verify_password("Teacher123", stored)
        assert is_valid is True
        assert needs_upgrade is False

    def test_password_hash_rejects_wrong_password():
        stored = _hash_password("Teacher123")
        is_valid, _ = _verify_password("WrongPass1", stored)
        assert is_valid is False

    def test_legacy_plain_sha256_hash_still_verifies():
        legacy_hash = hashlib.sha256("1234".encode("utf-8")).hexdigest()
        is_valid, needs_upgrade = _verify_password("1234", legacy_hash)
        assert is_valid is True
        assert needs_upgrade is True, "يجب رفع علم الترقية التلقائية لحساب بالتنسيق القديم"

    def test_legacy_hash_rejects_wrong_password():
        legacy_hash = hashlib.sha256("1234".encode("utf-8")).hexdigest()
        is_valid, needs_upgrade = _verify_password("9999", legacy_hash)
        assert is_valid is False
        assert needs_upgrade is False

    def test_is_valid_password_rejects_four_digit_pin():
        assert _is_valid_password("1234") is False, "الأرقام الأربعة وحدها لم تعد كافية"

    def test_is_valid_password_rejects_short_password():
        assert _is_valid_password("ab1") is False

    def test_is_valid_password_rejects_letters_only():
        assert _is_valid_password("abcdef") is False

    def test_is_valid_password_rejects_digits_only():
        assert _is_valid_password("123456") is False

    def test_is_valid_password_accepts_six_char_alnum():
        assert _is_valid_password("abc123") is True

    # =====================================================================================
    # === التسجيل وتسجيل الدخول عبر قاعدة البيانات الفعلية (SQLite مؤقتة) ===
    # =====================================================================================

    _DUMMY_LICENSE_KEY = "DUMMY-LICENSE-KEY-0001"

    def test_register_rejects_weak_password():
        success, message = register_teacher("teacher_a", "1234", _DUMMY_LICENSE_KEY)
        assert success is False
        assert "٦ خانات" in message

    def test_register_then_login_succeeds():
        success, _ = register_teacher("teacher_b", "Secret12", _DUMMY_LICENSE_KEY)
        assert success is True
        success, message = login_teacher("teacher_b", "Secret12")
        assert success is True

    def test_register_duplicate_username_rejected():
        register_teacher("teacher_c", "Secret12", _DUMMY_LICENSE_KEY)
        success, message = register_teacher("teacher_c", "Another99", _DUMMY_LICENSE_KEY)
        assert success is False
        assert "محجوز" in message

    def test_login_wrong_password_rejected():
        register_teacher("teacher_d", "Secret12", _DUMMY_LICENSE_KEY)
        success, message = login_teacher("teacher_d", "WrongPass")
        assert success is False

    def test_login_unknown_username_rejected():
        success, message = login_teacher("no_such_user", "Whatever1")
        assert success is False
        assert "لا يوجد حساب" in message

    def test_account_locks_after_max_failed_attempts():
        register_teacher("teacher_e", "Secret12", _DUMMY_LICENSE_KEY)
        for _ in range(MAX_FAILED_LOGIN_ATTEMPTS - 1):
            success, _ = login_teacher("teacher_e", "WrongPass")
            assert success is False
        # المحاولة الخاطئة الأخيرة يجب أن تُفعّل القفل المؤقت
        success, message = login_teacher("teacher_e", "WrongPass")
        assert success is False
        assert "قفل" in message

        # حتى بكلمة المرور الصحيحة، الحساب مقفل مؤقتاً
        success, message = login_teacher("teacher_e", "Secret12")
        assert success is False
        assert "قفل" in message

    def test_successful_login_resets_failed_attempts():
        register_teacher("teacher_f", "Secret12", _DUMMY_LICENSE_KEY)
        login_teacher("teacher_f", "WrongPass")
        login_teacher("teacher_f", "WrongPass")
        success, _ = login_teacher("teacher_f", "Secret12")
        assert success is True
        # بعد نجاح الدخول، عداد المحاولات الخاطئة يجب أن يُصفَّر بحيث لا يُقفل الحساب بسرعة لاحقاً
        for _ in range(MAX_FAILED_LOGIN_ATTEMPTS - 1):
            success, _ = login_teacher("teacher_f", "WrongPass")
            assert success is False
        success, message = login_teacher("teacher_f", "Secret12")
        assert success is True, "لأن العداد أُعيد تصفيره بعد الدخول الناجح، لم يصل بعد لحد القفل"

    def test_legacy_account_upgraded_after_successful_login():
        conn = get_db()
        cur = conn.cursor()
        cur.execute(
            _q("INSERT INTO teachers (teacher_name, pin_hash, created_at, failed_attempts) VALUES (?, ?, ?, 0)"),
            ("legacy_teacher", hashlib.sha256("1234".encode("utf-8")).hexdigest(), "2024-01-01T00:00:00")
        )
        conn.commit()
        release_db(conn)

        success, _ = login_teacher("legacy_teacher", "1234")
        assert success is True

        conn = get_db()
        cur = conn.cursor()
        cur.execute(_q("SELECT pin_hash FROM teachers WHERE teacher_name = ?"), ("legacy_teacher",))
        row = _row_to_dict(cur, cur.fetchone())
        release_db(conn)
        assert row["pin_hash"].startswith("pbkdf2_sha256$"), "الحساب يجب أن يُرقَّى تلقائياً للتجزئة الجديدة"

    # =====================================================================================
    # === معالجة نص الذكاء الاصطناعي (parse_ai_sections) ===
    # =====================================================================================

    def test_parse_ai_sections_full_response():
        full_text = (
            "نص ورقة العمل هنا.\n"
            "### ANSWER_KEY_JSON ###\n"
            '[{"q": "ما هو ٢+٢؟", "a": "٤"}]\n'
            "### KEY_VOCAB ###\n"
            "جمع, طرح, عدد"
        )
        main_text, answer_key, vocab_words = parse_ai_sections(full_text)
        assert main_text == "نص ورقة العمل هنا."
        assert answer_key == [{"q": "ما هو ٢+٢؟", "a": "٤"}]
        assert vocab_words == ["جمع", "طرح", "عدد"]

    def test_parse_ai_sections_truncated_response_loses_nothing_it_has():
        """== الخطأ الحقيقي الذي اكتُشف بالاختبار: عندما ينقطع رد الذكاء الاصطناعي قبل
        كتابة قسم الإجابات (شائع بالأوراق الطويلة بسبب حد max_output_tokens)، يجب أن
        يبقى نص ورقة العمل الرئيسي كاملاً على الأقل، مع قوائم فارغة (وليس استثناء) لبقية
        الأقسام، بحيث يستطيع الاستدعاء الاحتياطي (_generate_structured_extras) استكمالها
        لاحقاً بدل أن تُفقد ورقة العمل بأكملها. =="""
        truncated_text = "نص ورقة العمل الطويلة جداً... وينقطع الرد هنا فجأة"
        main_text, answer_key, vocab_words = parse_ai_sections(truncated_text)
        assert main_text == truncated_text
        assert answer_key == []
        assert vocab_words == []

    def test_parse_ai_sections_malformed_json_falls_back_to_bracket_extraction():
        full_text = (
            "نص الورقة\n"
            "### ANSWER_KEY_JSON ###\n"
            'مقدمة غير متوقعة من النموذج [{"q": "س", "a": "ج"}] كلام إضافي بعدها\n'
            "### KEY_VOCAB ###\n"
            "كلمة"
        )
        main_text, answer_key, vocab_words = parse_ai_sections(full_text)
        assert answer_key == [{"q": "س", "a": "ج"}]

    def test_parse_ai_sections_code_fenced_json():
        full_text = (
            "نص الورقة\n"
            "### ANSWER_KEY_JSON ###\n"
            '```json\n[{"q": "س", "a": "ج"}]\n```\n'
            "### KEY_VOCAB ###\n"
            "كلمة1, كلمة2"
        )
        main_text, answer_key, vocab_words = parse_ai_sections(full_text)
        assert answer_key == [{"q": "س", "a": "ج"}]
        assert vocab_words == ["كلمة1", "كلمة2"]

    def test_parse_ai_sections_vocab_capped_at_six_words():
        full_text = (
            "نص\n### ANSWER_KEY_JSON ###\n[]\n### KEY_VOCAB ###\n"
            "a, b, c, d, e, f, g, h"
        )
        _, _, vocab_words = parse_ai_sections(full_text)
        assert len(vocab_words) == 6

    def test_parse_ai_sections_no_special_markers_returns_full_text_unchanged():
        plain_text = "مجرد نص عادي بدون أي أقسام خاصة."
        main_text, answer_key, vocab_words = parse_ai_sections(plain_text)
        assert main_text == plain_text
        assert answer_key == []
        assert vocab_words == []

    # =====================================================================================
    # === تنسيق نص الامتحان التقييمي ===
    # =====================================================================================

    def test_format_exam_text_highlights_model_answer_as_blockquote():
        raw = "السؤال: ما هو ٢+٢؟\nالإجابة النموذجية: ٤"
        formatted = _format_exam_text_for_display(raw)
        assert "> ✅ **الإجابة النموذجية:** ٤" in formatted
        assert "السؤال: ما هو ٢+٢؟" in formatted

    def test_format_exam_text_leaves_other_lines_untouched():
        raw = "عنوان الامتحان\nسؤال ١\nسؤال ٢"
        formatted = _format_exam_text_for_display(raw)
        assert formatted == raw
