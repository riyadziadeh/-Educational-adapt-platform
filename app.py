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

# جلب مفتاح الـ API بمرونة تامة (سواء من الأسرار أو من متغيرات البيئة)
api_key = None
try:
    api_key = st.secrets["GOOGLE_API_KEY"]
except Exception:
    api_key = os.getenv("GOOGLE_API_KEY")

if not api_key:
    st.error("الرجاء ضبط مفتاح GOOGLE_API_KEY في إعدادات الأمان (Secrets) أو متغيرات البيئة لتشغيل النظام.")
else:
    genai.configure(api_key=api_key)

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
            with st.spinner(f"جاري معالجة ورقة العمل ({mode_desc}) بالتفصيل الكامل... / Processing..."):
                
                if generate_alternative:
                    prompt = f"""
                    أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن. 
                    مطلوب منك كتابة محتوى **كامل ومتشعب وشامل** وغير مقتضب إطلاقاً.
                    بناءً على محتوى ورقة العمل المستخرجة أدناه لمادة ({selected_subject}), صمم ورقة عمل بديلة مقترحة بالكامل مع **بنك أسئلة تقييمي تشخيصي مفصل يتضمن الأسئلة كاملة والحلول النموذجية** يناسب الحالة الخاصة ({selected_condition}) ومستوى التكييف ({selected_level}).
                    
                    التفاصيل:
                    - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                    - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن
                    
                    محتوى ورقة العمل الأصلية للاستئناس:
                    {extracted_content}
                    
                    تعليمات صارمة جداً: ممنوع الاختصار أو الاكتفاء بالعناوين أو الملخصات. اكتب ورقة العمل والأسئلة والتمارين والحلول بخطوات تفصيلية كاملة وواضحة للنهاية.
                    """
                else:
                    prompt = f"""
                    أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن. 
                    مطلوب منك تنفيذ **تكييف وتطوير شامل وكامل ودقيق** لورقة العمل التالية لمادة ({selected_subject}) بناءً على مستوى التكييف ({selected_level}) والحالة الخاصة ({selected_condition}).
                    
                    التفاصيل:
                    - الصف: {selected_grade} | النظام: {selected_system} | المادة: {selected_subject}
                    - لغة المخرجات: {selected_language} | المحافظة: {selected_gov} - الأردن
                    
                    محتوى ورقة العمل المستخرج من الملف:
                    {extracted_content}
                    
                    تعليمات صارمة جداً: ممنوع الاختصار أو الاكتفاء بالوصف العام أو المقدمات. قم بإعادة صياغة ورقة العمل الأصلية وكتابة كافة الأسئلة المعدلة، التمارين التدريبية، الأنشطة، والحلول بشكل كامل ووافٍ دون أي نقصان حتى النهاية.
                    """
                
                adapted_text = None
                # النماذج المدعومة والمستقرة في المكتبة لتجنب أي أخطاء
                models_to_try = ["gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-1.5-pro"]

                for model_name in models_to_try:
                    success_with_model = False
                    for attempt in range(2):
                        try:
                            model = genai.GenerativeModel(model_name)
                            # ضبط الـ Generation Config لضمان السماح بردود طويلة ومفصلة دون تقطيع
                            response = model.generate_content(
                                prompt, 
                                generation_config=genai.types.GenerationConfig(
                                    temperature=0.7,
                                    max_output_tokens=8192
                                )
                            )
                            if response and response.text:
                                adapted_text = response.text
                                success_with_model = True
                                break
                        except Exception as e:
                            time.sleep(2)
                            continue
                    if success_with_model and adapted_text:
                        break

                if not adapted_text:
                    st.warning("عذراً، لم يتم استجابة الخادم بالكامل. الرجاء المحاولة مرة أخرى بعد ثوانٍ.")

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
