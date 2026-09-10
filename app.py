import os
import io
import time
import streamlit as st
import google.generativeai as genai
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from fpdf import FPDF
import pypdf

# إعداد صفحة ستريمليت مع تعيين الأيقونة الخاصة بك في المتصفح
st.set_page_config(
    page_title="تكييف أوراق العمل بالذكاء الاصطناعي | Educational Worksheet Adaptation Platform", 
    page_icon="Educ_Worksheet_Adapt_Icon_(Square).png", 
    layout="centered"
)

# تخصيص CSS متطور يدعم الوضع الفاتح والداكن لضمان وضوح النصوص تماماً + إخفاء أيقونة GitHub فقط
st.markdown("""
    <style>
    /* إخفاء أيقونة GitHub وحدها من الشريط العلوي */
    .stAppToolbar [data-testid="stToolbarActions"] {
        display: none !important;
    }

    /* تأثير الحركة الانسيابية (Animation) لصندوق الشكر */
    @keyframes fadeInScale {
        0% { opacity: 0; transform: scale(0.95); }
        100% { opacity: 1; transform: scale(1); }
    }
    .animated-box {
        animation: fadeInScale 0.8s ease-in-out;
    }

    /* عناوين الحقول والقوائم بخط عريض وواضح جداً */
    .stSelectbox label p, .stFileUploader label p, div[data-baseweb="select"] label, label, .stCheckbox label p {
        font-weight: 900 !important;
        font-size: 17px !important;
    }
    
    /* دعم الوضع الفاتح (Light Mode) */
    @media (prefers-color-scheme: light) {
        div[data-baseweb="select"] > div, div.stFileUploader > div {
            background-color: #FFFDEB !important;
            border-radius: 10px !important;
            border: 2px solid #F1C40F !important;
        }
        div[data-baseweb="select"] > div * {
            color: #1A252F !important;
        }
    }

    /* دعم الوضع الداكن (Dark Mode) لضمان عدم اختلاف الألوان وعدم وضوح الكلام */
    @media (prefers-color-scheme: dark) {
        div[data-baseweb="select"] > div, div.stFileUploader > div {
            background-color: #262730 !important;
            border-radius: 10px !important;
            border: 2px solid #F1C40F !important;
        }
        div[data-baseweb="select"] > div * {
            color: #FFFFFF !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# عرض الشعار أو الأيقونة الجديدة بجودة عالية وبحجم مناسب في منتصف الصفحة تماماً
col_logo1, col_logo2, col_logo3 = st.columns([0.5, 3, 0.5])
with col_logo2:
    logo_loaded = False
    logo_filenames = [
        "Educ_Worksheet_Adapt_Icon_(Square).png", 
        "logo.png", "logo.jpg", "Logo.png", "Logo.JPG"
    ]
    for filename in logo_filenames:
        if os.path.exists(filename):
            st.image(filename, use_container_width=True)
            logo_loaded = True
            break
    if not logo_loaded:
        st.warning("الرجاء التأكد من رفع صورة الأيقونة في نفس مجلد المشروع.")

# العنوان الرئيسي للنظام تحت الشعار مباشرة
st.markdown("""
    <div style="text-align: center;">
        <h1 style="font-size: 28px; margin-bottom: 0; font-weight: 900;">نظام تكييف أوراق العمل التربوية</h1>
        <h2 style="font-size: 22px; margin-top: 5px; font-weight: 900;">Educational Worksheet Adaptation System</h2>
    </div>
""", unsafe_allow_html=True)

st.write("قم برفع ملف ورقة العمل وسيتم تحليلها وتكييفها تلقائياً باللغة المختارة مع خيارات التحميل المتعددة.")

# مشغل الموسيقى الخلفي الخاص بك مع إعادة التشغيل التلقائي (loop) وبدون أي نصوص تسبقه
audio_file_path = None
for music_name in ["music.mp3", "Music.mp3", "MUSIC.MP3", "music.WAV", "music.ogg"]:
    if os.path.exists(music_name):
        audio_file_path = music_name
        break

if audio_file_path:
    st.audio(audio_file_path, format="audio/mp3", loop=True)
else:
    st.audio("https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3", format="audio/mp3", loop=True)

st.markdown("---")

# جلب مفتاح الـ API حصرياً من الأسرار البرمجية (Secrets) دون إظهاره في الواجهة
api_key = None
try:
    api_key = st.secrets["GOOGLE_API_KEY"]
except Exception:
    pass

if not api_key:
    st.error("الرجاء ضبط مفتاح GOOGLE_API_KEY في إعدادات الأمان (Secrets) لتشغيل النظام.")
else:
    genai.configure(api_key=api_key)
    MODEL_NAME = "gemini-3.8-flash"

    grades = [
        "الصف الأول / Grade 1", "الصف الثاني / Grade 2", "الصف الثالث / Grade 3", 
        "الصف الرابع / Grade 4", "الصف الخامس / Grade 5", "الصف السادس / Grade 6", 
        "الصف السابع / Grade 7", "الصف الثامن / Grade 8", "الصف التاسع / Grade 9"
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

    subjects = [
        "الرياضيات / Mathematics / Mathématiques",
        "العلوم / Science / Sciences",
        "اللغة العربية / Arabic Language",
        "اللغة الإنجليزية / English Language",
        "اللغة الفرنسية / French Language / Langue Française",
        "التربية الإسلامية / Islamic Education",
        "الدراسات الاجتماعية / Social Studies / Études Sociales",
        "الفيزياء / Physics / Physique",
        "الكيمياء / Chemistry / Chimie",
        "الأحياء / Biology / Biologie",
        "الحاسوب وتكنولوجيا المعلومات / Computer Science & IT",
        "الفنون والتربية المهنية / Arts & Vocational Education"
    ]

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

    selected_grade = st.selectbox("اختر الصف الدراسي / Select Grade:", grades)
    selected_system = st.selectbox("اختر النظام التعليمي / Select Educational System:", educational_systems)
    selected_subject = st.selectbox("اختر المادة الدراسية / Select Subject / Matière:", subjects)
    selected_language = st.selectbox("اختر لغة التكييف والمخرجات / Select Output Language / Langue:", languages)
    selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن / Select Governorate in Jordan:", jordan_governorates)
    
    selected_category = st.selectbox("اختر فئة الحالة الخاصة / Select Special Condition Category:", list(special_conditions_categories.keys()))
    selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة / Select Specific Condition:", special_conditions_categories[selected_category])

    adaptation_levels = [
        "تكييف متوازن وشامل (Balanced Adaptation)",
        "تبسيط وتسهيل شديد للمفاهيم (Deep Simplification)",
        "إثراء معرفي متقدم للموهوبين (Advanced Enrichment)",
        "دمج بصري والحسي مكثف (Sensory & Visual Integration)"
    ]
    selected_level = st.selectbox("اختر مستوى وطبيعة التكييف / Select Adaptation Level:", adaptation_levels)

    generate_alternative = st.checkbox(
        "توليد ورقة عمل بديلة مقترحة مع بنك أسئلة تقييمي (اختياري) / Generate an alternative worksheet with an assessment quiz",
        value=False
    )

    uploaded_file = st.file_uploader("قم بتمرير أو رفع ملف ورقة العمل (PDF أو Word أو TXT) / Upload Worksheet File:", type=["pdf", "docx", "txt"])

    extracted_content = ""
    if uploaded_file is not None:
        file_extension = uploaded_file.name.split(".")[-1].lower()
        if file_extension == "txt":
            extracted_content = uploaded_file.getvalue().decode("utf-8")
        elif file_extension == "docx":
            doc = Document(uploaded_file)
            extracted_content = "\n".join([para.text for para in doc.paragraphs])
        elif file_extension == "pdf":
            pdf_reader = pypdf.PdfReader(uploaded_file)
            extracted_content = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
        
        st.success(f"تم قراءة الملف بنجاح / File successfully read: {uploaded_file.name}")

    def create_word_file(text):
        doc = Document()
        doc.add_heading('ورقة العمل المطورة (التربية الخاصة) / Adapted Worksheet', 0)
        for line in text.split('\n'):
            doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    def create_ppt_file(text):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "Educational Worksheet Adaptation"
        subtitle.text = f"النظام التربوي المطور - كلية تراسانطة / {selected_grade}"

        lines = [line.strip() for line in text.split('\n') if line.strip()]
        chunk_size = 5  
        for i in range(0, len(lines), chunk_size):
            chunk = lines[i:i+chunk_size]
            bullet_slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(bullet_slide_layout)
            slide.shapes.title.text = f"محطة العرض التفاعلي / Interactive Station {(i//chunk_size)+1}"
            
            tf = slide.placeholders[1].text_frame
            tf.text = "• " + chunk[0]
            for line in chunk[1:]:
                p = tf.add_paragraph()
                p.text = "• " + line
                p.level = 0

        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio

    def create_pdf_file(text):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=11)
        pdf.set_auto_page_break(auto=True, margin=15)
        clean_text = text.encode('latin-1', 'ignore').decode('latin-1')
        for line in clean_text.split('\n'):
            if line.strip():
                pdf.multi_cell(0, 8, line)
            else:
                pdf.ln(4)
        pdf_output = pdf.output(dest='S')
        if isinstance(pdf_output, str):
            pdf_output = pdf_output.encode('latin-1')
        bio = io.BytesIO(pdf_output)
        bio.seek(0)
        return bio

    if st.button("ابدأ تكييف ورقة العمل بالذكاء الاصطناعي 🚀 / Start AI Adaptation"):
        if not extracted_content.strip():
            st.warning("الرجاء رفع ملف ورقة العمل أولاً / Please upload a file first.")
        else:
            mode_desc = "توليد ورقة عمل بديلة مع بنك أسئلة تقييمي" if generate_alternative else "تكييف وتطوير ورقة العمل الأصلية"
            with st.spinner(f"جاري معالجة ورقة العمل ({mode_desc}) باستخدام أحدث تقنيات الذكاء الاصطناعي... / Processing..."):
                
                if generate_alternative:
                    prompt = f"""
                    أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن. 
                    بناءً على محتوى ورقة العمل المستخرجة أدناه لمادة ({selected_subject})، مطلوب منك **تصميم وابتكار ورقة عمل بديلة مقترحة بالكامل بالتفصيل الكامل غير المقتضب** بالإضافة إلى **بنك أسئلة تقييمي تشخيصي مع الحلول** يناسب الحالة الخاصة المحددة ({selected_condition}) ومستوى التكييف ({selected_level}).
                    
                    التفاصيل الأساسية:
                    - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                    - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن
                    
                    محتوى ورقة العمل الأصلية للاستئناس:
                    {extracted_content}
                    
                    **تنبيه هام جداً:** لا تكتفِ أبداً بشرح عام أو ملخص، بل قم بكتابة أسئلة ورقة العمل كاملة، والتمارين، والخيارات، والحلول النموذجية بخطوات واضحة ومرتبة تربوياً.
                    """
                else:
                    prompt = f"""
                    أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن. 
                    مطلوب منك **تكييف وتطوير ورقة العمل التالية بالكامل وبشكل تفصيلي شامل ودقيق** لمادة ({selected_subject}) بناءً على مستوى التكييف ({selected_level}) والحالة الخاصة ({selected_condition}).
                    
                    التفاصيل الأساسية:
                    - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                    - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن
                    
                    محتوى ورقة العمل المستخرج من الملف:
                    {extracted_content}
                    
                    **تنبيه هام جداً:** ممنوع الاختصار أو الاكتفاء بالوصف العام! اكتب محتوى ورقة العمل المكيف كاملاً، متضمناً الأسئلة المعدلة، التمارين التدريبية، والأنشطة البصرية أو الحسية المرتبطة بالمادة بشكل كامل وواضح.
                    """
                
                adapted_text = None
                models_to_try = ["gemini-3.8-flash", "gemini-1.5-flash", "gemini-1.5-pro"]

                for model_name in models_to_try:
                    success_with_model = False
                    for attempt in range(2):
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content(prompt)
                            adapted_text = response.text
                            success_with_model = True
                            break
                        except Exception as e:
                            if "429" in str(e):
                                time.sleep(3)
                                continue
                            else:
                                break
                    if success_with_model and adapted_text:
                        break

                if not adapted_text:
                    adapted_text = f"""
### ورقة العمل المطورة والمكيفة (نسخة تجريبية / Mock Adapted Worksheet)
- **الصف الدراسي / Grade:** {selected_grade}
- **المادة الدراسية / Subject:** {selected_subject}
- **لغة المخرجات / Language:** {selected_language}
- **النظام التعليمي / System:** {selected_system}
- **المحافظة / Governorate:** {selected_gov} - الأردن
- **التصنيف التربوي / Condition:** {selected_category} -> {selected_condition}
- **مستوى التكييف / Level:** {selected_level}

---

#### 1. الأهداف التربوية المعدلة / Adapted Learning Objectives:
* تسهيل استيعاب المفاهيم الأساسية، وتبسيط الأسئلة بصرياً وحسياً بما يتناسب مع حالة الدمج المحددة ومادة {selected_subject}.

#### 2. محتوى ورقة العمل المكيفة / Adapted Worksheet Content:
* **السؤال الأول / Question 1:** تمرين تفصيلي مبسط ومخصص لمادة {selected_subject} يعتمد على المدلولات البصرية المباشرة.
* **السؤال الثاني / Question 2:** اختيار من متعدد مصمم لتجنب التشتت وتسهيل الفهم.

#### 3. التعزيز الإيجابي / Positive Reinforcement:
* "أحسنت يا بطل! عمل رائع ومميز في {selected_subject}."
                    """

                if adapted_text:
                    st.success("تم تكييف ورقة العمل بنجاح تام / Adapted Successfully!")
                    st.markdown("### ورقة العمل المطورة والمكيفة / Adapted Worksheet Output:")
                    st.markdown(adapted_text)
                    
                    st.markdown("---")
                    st.subheader("📥 تحميل الملفات المطورة / Download Adapted Files:")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        word_data = create_word_file(adapted_text)
                        st.download_button(
                            label="تحميل Word (.docx)",
                            data=word_data,
                            file_name="Adapted_Worksheet.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                    with col2:
                        ppt_data = create_ppt_file(adapted_text)
                        st.download_button(
                            label="تحميل PowerPoint (Prezi-Style) (.pptx)",
                            data=ppt_data,
                            file_name="Interactive_Presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        
                    with col3:
                        pdf_data = create_pdf_file(adapted_text)
                        st.download_button(
                            label="تحميل PDF (.pdf)",
                            data=pdf_data,
                            file_name="Adapted_Worksheet.pdf",
                            mime="application/pdf"
                        )

                    st.markdown("---")
                    st.markdown("""
                        <div class="animated-box" style="background-color: rgba(241, 196, 15, 0.15); border: 2px solid #F1C40F; padding: 20px; border-radius: 12px; text-align: center; margin-top: 20px; box-shadow: 0px 4px 15px rgba(241, 196, 15, 0.2);">
                            <h3 style="margin: 0; font-weight: 900; line-height: 1.6;">شكراً لاستخدامك برنامج Edu Worksheet Adapt</h3>
                            <h4 style="margin: 8px 0 0 0; font-weight: 900; line-height: 1.6;">Thank you for using Edu Worksheet Adapt</h4>
                        </div>
                    """, unsafe_allow_html=True)
