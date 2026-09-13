import os
import io
import sys
import time
import base64
import random
import subprocess
import tempfile
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
# === هوية الألوان الرسمية للتطبيق: كحلي غامق + كحلي/أزرق فاتح + ذهبي + أبيض ===
# =========================================================================================
NAVY_DARK = "#101B2D"      # الأزرق الكحلي الغامق (خلفية الهيدر والبطاقات المختارة)
NAVY_LIGHT = "#EAF1FB"     # الأزرق الفاتح جداً (خلفية الشبكة العامة)
BLUE_ACCENT = "#2E6FBB"    # أزرق متوسط لدعم التدرجات والحدود
GOLD = "#F1C40F"           # الذهبي (لون التمييز والعناصر النشطة)
WHITE = "#FFFFFF"

st.markdown(f"""
    <style>
    /* إخفاء أيقونة GitHub وحدها من الشريط العلوي */
    .stAppToolbar [data-testid="stToolbarActions"] {{
        display: none !important;
    }}

    /* خلفية عامة فاتحة زرقاء هادئة تماشياً مع هوية التطبيق */
    .stApp {{
        background-color: {NAVY_LIGHT};
    }}

    /* تأثير الحركة الانسيابية (Animation) لصندوق الشكر */
    @keyframes fadeInScale {{
        0% {{ opacity: 0; transform: scale(0.95); }}
        100% {{ opacity: 1; transform: scale(1); }}
    }}
    .animated-box {{
        animation: fadeInScale 0.8s ease-in-out;
    }}

    /* شريط علوي كحلي غامق يشبه هيدر التطبيقات مع بحث وهمي للزينة */
    .app-topbar {{
        background: linear-gradient(135deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%);
        border-radius: 18px;
        padding: 18px 20px;
        margin-bottom: 22px;
        box-shadow: 0px 6px 18px rgba(16,27,45,0.25);
    }}
    .app-topbar .search-fake {{
        background-color: {WHITE};
        border-radius: 12px;
        padding: 10px 16px;
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

    /* صناديق الاختيار (Selectbox) و رافع الملفات بهوية كحلي/ذهبي موحّدة لكل الأوضاع */
    div[data-baseweb="select"] > div, div.stFileUploader > div {{
        background-color: {WHITE} !important;
        border-radius: 12px !important;
        border: 2px solid {GOLD} !important;
    }}
    div[data-baseweb="select"] > div * {{
        color: {NAVY_DARK} !important;
    }}

    /* بطاقات الشبكة الشبيهة بواجهة التطبيق المرفقة */
    div[data-testid="stButton"] > button {{
        width: 100%;
        min-height: 108px;
        height: auto !important;
        padding: 16px 12px !important;
        border-radius: 16px !important;
        border: 2px solid {BLUE_ACCENT}33 !important;
        background-color: {WHITE} !important;
        color: {NAVY_DARK} !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        line-height: 1.6 !important;
        box-shadow: 0px 3px 10px rgba(16,27,45,0.08);
        transition: all 0.15s ease-in-out;
        white-space: pre-line !important;
    }}
    div[data-testid="stButton"] > button:hover {{
        border: 2px solid {GOLD} !important;
        transform: translateY(-2px);
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
        background-color: {WHITE} !important;
        color: {BLUE_ACCENT} !important;
        font-weight: 900 !important;
        font-size: 19px !important;
        line-height: 1.6 !important;
        height: auto !important;
        min-height: 90px !important;
        padding: 18px 14px !important;
        border-radius: 18px !important;
        border: 3px solid {NAVY_DARK} !important;
        box-shadow: 0px 6px 16px rgba(16,27,45,0.18);
        white-space: normal !important;
        overflow: visible !important;
        word-wrap: break-word !important;
    }}
    div[class*="st-key-start-ai-button"] div[data-testid="stButton"] > button:hover {{
        border: 3px solid {GOLD} !important;
        color: {BLUE_ACCENT} !important;
    }}
    div[class*="st-key-start-ai-button"] div[data-testid="stButton"] > button p {{
        color: {BLUE_ACCENT} !important;
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
        background: linear-gradient(160deg, {WHITE} 0%, {NAVY_LIGHT} 100%) !important;
        border: 3px solid {BLUE_ACCENT}55 !important;
        font-size: 22px !important;
        line-height: 1 !important;
        padding: 0 !important;
        margin: 0 auto !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0px 4px 10px rgba(16,27,45,0.12);
    }}
    div[class*="st-key-circlebtn-"] button p {{
        margin: 0 !important;
        line-height: 1 !important;
    }}
    div[class*="st-key-circlebtn-"] button:hover {{
        border: 3px solid {GOLD} !important;
        transform: scale(1.08);
    }}
    div[class*="st-key-circlebtn-"][class*="-on"] button {{
        border: 3px solid {GOLD} !important;
        background: linear-gradient(160deg, {NAVY_DARK} 0%, {BLUE_ACCENT} 100%) !important;
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

    .selection-summary {{
        background-color: {NAVY_DARK};
        color: {GOLD};
        border-radius: 12px;
        padding: 10px 16px;
        text-align: center;
        font-weight: 800;
        margin: 10px 0 18px 0;
    }}
    </style>
""", unsafe_allow_html=True)

# عرض الشعار الجديد (new_logo.png) بجودة عالية وبحجم مناسب في منتصف الصفحة تماماً
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
            image = Image.open(filename)
            st.image(image, use_container_width=True)
            logo_loaded = True
            break
    if not logo_loaded:
        st.warning("الرجاء التأكد من رفع صورة الأيقونة باسم new_logo.png في نفس مجلد المشروع.")

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
# --- تعديل: الموسيقى الخلفية أصبحت اختيارية بالكامل (Opt-in) بدل التشغيل التلقائي ---
# هذا مهم خصوصاً لأن جزءاً كبيراً من مستخدمي هذا النظام هم طلاب لديهم حساسية حسية
# (مثل اضطراب طيف التوحد)، وصوت يعمل من تلقاء نفسه قد يكون مزعجاً أو مربكاً لهم.
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

    # ---------------- شبكة دوائر اختيار المادة الدراسية (٤ مواد بجانب بعض) ----------------
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

    # ---------------- بقية الحقول عبر قوائم منسدلة بنفس الهوية اللونية (كحلي/ذهبي/أبيض) ----------------
    selected_grade = st.selectbox("اختر الصف الدراسي / Select Grade:", grades)
    selected_system = st.selectbox("اختر النظام التعليمي / Select Educational System:", educational_systems)
    selected_language = st.selectbox("اختر لغة التكييف والمخرجات / Select Output Language / Langue:", languages)
    selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن / Select Governorate in Jordan:", jordan_governorates)

    selected_category = st.selectbox("اختر فئة الحالة الخاصة / Select Special Condition Category:", list(special_conditions_categories.keys()))
    selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة / Select Specific Condition:", special_conditions_categories[selected_category])

    uploaded_file = st.file_uploader("قم بتمرير أو رفع ملف ورقة العمل (PDF أو Word أو TXT) / Upload Worksheet File:", type=["pdf", "docx", "txt"])

    MAX_INPUT_CHARS = 20000

    extracted_content = ""
    if uploaded_file is not None:
        file_extension = uploaded_file.name.split(".")[-1].lower()
        try:
            if file_extension == "txt":
                extracted_content = uploaded_file.getvalue().decode("utf-8")
            elif file_extension == "docx" and DOCX_AVAILABLE:
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
                # --- تعديل: إعلام المستخدم بوضوح إذا كان جزء من الملف سيُقص قبل إرساله للنموذج ---
                if len(extracted_content) > MAX_INPUT_CHARS:
                    st.info(
                        f"⚠️ الملف المرفوع يحتوي على {len(extracted_content):,} حرفاً، وسيتم إرسال أول "
                        f"{MAX_INPUT_CHARS:,} حرف فقط منه إلى نموذج الذكاء الاصطناعي بسبب حدود المعالجة الحالية. "
                        "لضمان تغطية الورقة كاملة، يُفضّل تقسيم الملفات الطويلة إلى أجزاء أصغر قبل الرفع."
                    )
            else:
                st.warning("⚠️ الملف المرفوع لا يحتوي على نص قابل للقراءة المباشرة. سيتم الاعتماد على معلومات النظام والعنوان لتوليد ورقة العمل.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    # =====================================================================================
    # === تعديل: دعم اتجاه RTL الصحيح في مستندات Word (محاذاة يمين + خاصية bidi فعلية) ===
    # =====================================================================================
    def _set_paragraph_rtl(paragraph):
        """يضبط الفقرة لتكون بمحاذاة اليمين وباتجاه RTL فعلي (وليس فقط محاذاة بصرية)."""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pPr = paragraph._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), "1")
        pPr.append(bidi)

    def create_word_file(text):
        if not DOCX_AVAILABLE:
            return None
        doc = Document()
        heading = doc.add_heading('ورقة العمل المطورة (التربية الخاصة) / Adapted Worksheet', 0)
        _set_paragraph_rtl(heading)
        for line in text.split('\n'):
            p = doc.add_paragraph(line)
            _set_paragraph_rtl(p)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

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
        وتُستخدم كصور توضيحية داخل شرائح PowerPoint (بديل ذاتي التوليد بدل صور خارجية).
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

    _ai_images_state = {"available": True}

    def _generate_ai_illustration(prompt_text):
        if not _ai_images_state["available"]:
            return None
        try:
            response = client.models.generate_images(
                model="imagen-4.0-generate-001",
                prompt=(
                    "رسمة تعليمية بسيطة بأسلوب Flat Design نظيف وواضح، بدون أي كتابة أو حروف "
                    "أو أرقام داخل الصورة إطلاقاً، بألوان هادئة تتناسق مع الذهبي (#F1C40F) "
                    f"والكحلي الداكن (#101B2D)، توضّح بصرياً الفكرة التالية: {prompt_text}"
                ),
                config=types.GenerateImagesConfig(number_of_images=1, aspect_ratio="1:1"),
            )
            img_bytes = response.generated_images[0].image.image_bytes
            bio = io.BytesIO(img_bytes)
            bio.seek(0)
            return bio
        except Exception:
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
    # === تعديل: تقسيم أذكى لمحتوى الشرائح — يجمع حسب الفقرات الطبيعية (فواصل الأسطر
    # الفارغة) بدل تقسيم كل 5 أسطر بشكل عشوائي قد يقطع سؤالاً أو فكرة في المنتصف ===
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
        if not PPTX_AVAILABLE:
            return None

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
                slide.shapes.add_picture(filename, Inches(0.7), Inches(0.25), height=Inches(1.9))
                break

        cover_illustration = _generate_ai_illustration(
            f"موضوع مادة {selected_subject} لطلاب {selected_grade}"
        )
        if cover_illustration:
            slide.shapes.add_picture(cover_illustration, Inches(8.6), Inches(1.9), height=Inches(3.4))
        else:
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

            if slide_illustration:
                slide.shapes.add_picture(slide_illustration, Inches(10.0), Inches(1.35), height=Inches(2.2))
                content_width = Inches(9.1)
            else:
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
            return None, f"تعذّر الاتصال بواجهة Canva: {e}"

    # =====================================================================================
    # === تعديل جوهري: دعم كامل للنص العربي داخل PDF عبر تشكيل الحروف (arabic_reshaper)
    # وترتيب الاتجاه (python-bidi) بالإضافة إلى خط Unicode عربي حقيقي بدل latin-1 القديم
    # الذي كان يحذف كل الحروف العربية بصمت. إن لم يتوفر خط عربي في مجلد المشروع، يُنبَّه
    # المستخدم بوضوح بدل إخراج ملف فارغ من المحتوى العربي دون علمه. ===
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
            try:
                subprocess.check_call([
                    sys.executable, "-m", "pip", "install", "--quiet",
                    "arabic-reshaper", "python-bidi"
                ])
                import arabic_reshaper as _ar_runtime
                from bidi.algorithm import get_display as _gd_runtime
                reshape_func = _ar_runtime.reshape
                display_func = _gd_runtime
                shaping_ready = True
                logs.append("تم تثبيت مكتبات دعم العربية (arabic-reshaper, python-bidi) تلقائياً وقت التشغيل.")
            except Exception as e:
                logs.append(f"تعذّر تثبيت مكتبات دعم العربية تلقائياً: {e}")

        font_path = _find_local_arabic_font()
        if not font_path:
            try:
                cache_path = os.path.join(tempfile.gettempdir(), "AutoArabicFont.ttf")
                if not os.path.exists(cache_path):
                    if REQUESTS_AVAILABLE:
                        resp = requests.get(_FALLBACK_ARABIC_FONT_URL, timeout=25)
                        resp.raise_for_status()
                        with open(cache_path, "wb") as f:
                            f.write(resp.content)
                    else:
                        raise RuntimeError("مكتبة requests غير متوفرة لتنزيل الخط تلقائياً.")
                if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
                    font_path = cache_path
                    logs.append("تم تنزيل خط عربي (Amiri) تلقائياً من الإنترنت لدعم PDF بالعربية.")
            except Exception as e:
                logs.append(f"تعذّر تنزيل خط عربي تلقائياً: {e}")

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
            return None, f"⚠️ حدث خطأ أثناء توليد ملف PDF: {e}"

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
            """
            height = 10 if is_title else 8
            align = "C" if is_title else "R"
            attempts = [
                lambda: pdf_obj.multi_cell(0, height, txt=_shape(raw_line), align=align),
                lambda: pdf_obj.multi_cell(0, height, txt=raw_line, align=align),
                lambda: pdf_obj.multi_cell(0, height, txt=raw_line.encode('latin-1', 'ignore').decode('latin-1') or "-", align=align),
            ]
            for attempt in attempts:
                try:
                    attempt()
                    return True
                except Exception:
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
    # === تعديل جوهري: توليد ملفات التحميل (Word/PPT/PDF) مرة واحدة فقط لكل نص مُكيَّف،
    # بدل إعادة توليدها (بما فيها صور الذكاء الاصطناعي المكلفة) في كل rerun من ستريمليت ===
    # =====================================================================================
    if "adapted_text" not in st.session_state:
        st.session_state.adapted_text = None
    if "just_generated" not in st.session_state:
        st.session_state.just_generated = False
    if "generated_files" not in st.session_state:
        st.session_state.generated_files = {}
    if "generated_for_text" not in st.session_state:
        st.session_state.generated_for_text = None

    try:
        start_btn_container = st.container(key="start-ai-button")
    except TypeError:
        start_btn_container = st.container()
    with start_btn_container:
        start_clicked = st.button("🚀 Start", use_container_width=True)

    if start_clicked:
        if not extracted_content.strip():
            extracted_content = f"ورقة عمل عامة لمبحث {selected_subject} للصف {selected_grade} وفق النظام {selected_system}."

        with st.spinner(" "):

            trimmed_content = extracted_content[:MAX_INPUT_CHARS] if len(extracted_content) > MAX_INPUT_CHARS else extracted_content

            if generate_alternative:
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن.
                مطلوب تصميم ورقة عمل بديلة مقترحة بالكامل مع **بنك أسئلة تقييمي تشخيصي مفصل يتضمن الأسئلة والحلول النموذجية** يناسب الحالة الخاصة ({selected_condition}) ومستوى التكييف ({selected_level}).

                البيانات الأساسية:
                - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن

                محتوى الملف المرفق:
                {trimmed_content}

                اكتب ورقة العمل والأسئلة والتمارين والحلول بخطوات تفصيلية كاملة وواضحة باللغة العربية.
                """
            else:
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن.
                مطلوب تنفيذ **تكييف وتطوير شامل ودقيق** لورقة العمل التالية لمبحث ({selected_subject}) بناءً على مستوى التكييف ({selected_level}) والحالة الخاصة ({selected_condition}).

                البيانات الأساسية:
                - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن

                محتوى الملف المرفق:
                {trimmed_content}

                قم بإعادة صياغة ورقة العمل وكتابة الأسئلة المعدلة، التمارين التدريبية، والحلول بشكل كامل ووافٍ دون أي نقصان وبأسلوب تربوي متميز.
                """

            adapted_text = None
            models_to_try = ["gemini-3.5-flash-lite", "gemini-2.5-flash"]
            last_error = None
            for model_name in models_to_try:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            temperature=0.7,
                            max_output_tokens=8000,
                        ),
                    )
                    if response and response.text:
                        adapted_text = response.text
                        break
                except Exception as e:
                    last_error = e
                    time.sleep(1)
                    continue

            if adapted_text:
                st.session_state.adapted_text = adapted_text
                st.session_state.just_generated = True
                st.success("تم تكييف ورقة العمل بنجاح تام / Adapted Successfully!")
            else:
                st.error("عذراً، تعذّر الاتصال بخدمة الذكاء الاصطناعي حالياً. يرجى المحاولة لاحقاً، أو التأكد من صلاحية مفتاح GOOGLE_API_KEY.")
                if last_error:
                    st.caption(f"تفاصيل تقنية: {last_error}")

    if st.session_state.adapted_text:

        if st.session_state.just_generated:
            if enable_ding:
                play_ready_ding()
            st.session_state.just_generated = False

        st.markdown("### ورقة العمل المطورة والمكيفة / Adapted Worksheet Output:")
        st.markdown(st.session_state.adapted_text)

        st.markdown("---")
        st.subheader("📥 تحميل الملفات المطورة / Download Adapted Files:")

        # --- تعديل: توليد الملفات مرة واحدة فقط لكل نص مُكيَّف جديد، وتخزينها كبايتات جاهزة ---
        current_text = st.session_state.adapted_text
        if st.session_state.generated_for_text != current_text:
            with st.spinner("جاري تجهيز ملفات Word و PowerPoint و PDF للتحميل (مرة واحدة فقط)..."):
                word_bio = create_word_file(current_text)
                ppt_bio = create_ppt_file(current_text)
                pdf_bio, pdf_warning = create_pdf_file(current_text)

                st.session_state.generated_files = {
                    "word": word_bio.getvalue() if word_bio else None,
                    "ppt": ppt_bio.getvalue() if ppt_bio else None,
                    "pdf": pdf_bio.getvalue() if pdf_bio else None,
                    "pdf_warning": pdf_warning,
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
                # نعرض سبب الفشل الحقيقي بدل عبارة عامة، حتى يعرف المستخدم بالضبط ما الناقص
                st.error(files.get("pdf_warning") or "تصدير PDF غير متوفر حالياً لسبب غير معروف.")

        st.markdown("---")
        st.markdown(f"""
            <div class="animated-box" style="background-color: rgba(241, 196, 15, 0.15); border: 2px solid {GOLD}; padding: 20px; border-radius: 12px; text-align: center; margin-top: 20px; box-shadow: 0px 4px 15px rgba(241, 196, 15, 0.2);">
                <h3 style="margin: 0; font-weight: 900; line-height: 1.6; color: {NAVY_DARK};">شكراً لاستخدامك برنامج Edu Worksheet Adapt</h3>
                <h4 style="margin: 8px 0 0 0; font-weight: 900; line-height: 1.6; color: {NAVY_DARK};">Thank you for using Edu Worksheet Adapt</h4>
            </div>
        """, unsafe_allow_html=True)
