import os
import io
import time
import streamlit as st
from google import genai
from google.genai import types
import pypdf
from PIL import Image

# محاولة استيراد مكتبات Word و PowerPoint بأمان تامة لضمان عدم انهيار السيرفر
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# إعداد صفحة ستريمليت مع تعيين الأيقونة الجديدة (new_logo.png) في المتصفح
st.set_page_config(
    page_title="تكييف أوراق العمل بالذكاء الاصطناعي | Educational Worksheet Adaptation Platform", 
    page_icon="new_logo.png", 
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
    client = genai.Client(api_key=api_key)

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
            else:
                st.warning("⚠️ الملف المرفوع لا يحتوي على نص قابل للقراءة المباشرة. سيتم الاعتماد على معلومات النظام والعنوان لتوليد ورقة العمل.")
        except Exception as e:
            st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

    def create_word_file(text):
        if DOCX_AVAILABLE:
            doc = Document()
            doc.add_heading('ورقة العمل المطورة (التربية الخاصة) / Adapted Worksheet', 0)
            for line in text.split('\n'):
                doc.add_paragraph(line)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            return bio
        return None

    def create_ppt_file(text):
        if PPTX_AVAILABLE:
            prs = Presentation()
            
            # الشريحة الأولى: غلاف احترافي بتصميم أنيق
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            
            # تلوين خلفية الغلاف بلون هادئ ومميز
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(245, 247, 250)
            
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = "Educational Worksheet Adaptation"
            subtitle.text = f"النظام التربوي المطور - كلية تراسانطة\n{selected_grade} | {selected_subject}"

            # تنسيق الشرائح اللاحقة للمحطات التفاعلية
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            chunk_size = 5  
            for i in range(0, len(lines), chunk_size):
                chunk = lines[i:i+chunk_size]
                bullet_slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(bullet_slide_layout)
                
                # تلوين خلفية الشرائح التفاعلية
                bg_fill = slide.background.fill
                bg_fill.solid()
                bg_fill.fore_color.rgb = RGBColor(255, 255, 255)
                
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
        return None

    def create_txt_file(text):
        bio = io.BytesIO(text.encode('utf-8'))
        bio.seek(0)
        return bio

    # تهيئة الذاكرة المؤقتة لمنع اختفاء النص عند التحميل
    if "adapted_text" not in st.session_state:
        st.session_state.adapted_text = None

    if st.button("ابدأ تكييف ورقة العمل بالذكاء الاصطناعي 🚀 / Start AI Adaptation"):
        if not extracted_content.strip():
            extracted_content = f"ورقة عمل عامة لمبحث {selected_subject} للصف {selected_grade} وفق النظام {selected_system}."

        mode_desc = "توليد ورقة عمل بديلة مع بنك أسئلة تقييمي" if generate_alternative else "تكييف وتطوير ورقة العمل الأصلية"
        
        with st.spinner(f"جاري معالجة ورقة العمل ({mode_desc}) وتحليلها عبر الذكاء الاصطناعي... يرجى الانتظار قليلاً..."):
            
            trimmed_content = extracted_content[:3500] if len(extracted_content) > 3500 else extracted_content

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
                            max_output_tokens=4000,
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
                st.success("تم تكييف ورقة العمل بنجاح تام / Adapted Successfully!")
            else:
                st.error("عذراً، تعذّر الاتصال بخدمة الذكاء الاصطناعي حالياً. يرجى المحاولة لاحقاً، أو التأكد من صلاحية مفتاح GOOGLE_API_KEY.")
                if last_error:
                    st.caption(f"تفاصيل تقنية: {last_error}")

    # عرض النتيجة وأزرار التحميل طالما أنها مخزنة في الذاكرة (لا تختفي عند الضغط على أي زر تحميل)
    if st.session_state.adapted_text:
        st.markdown("### ورقة العمل المطورة والمكيفة / Adapted Worksheet Output:")
        st.markdown(st.session_state.adapted_text)
        
        st.markdown("---")
        st.subheader("📥 تحميل الملفات المطورة / Download Adapted Files:")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            word_data = create_word_file(st.session_state.adapted_text)
            if word_data and DOCX_AVAILABLE:
                st.download_button(
                    label="تحميل Word (.docx)",
                    data=word_data,
                    file_name="Adapted_Worksheet.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.info("تصدير Word غير متوفر حالياً.")
            
        with col2:
            ppt_data = create_ppt_file(st.session_state.adapted_text)
            if ppt_data and PPTX_AVAILABLE:
                st.download_button(
                    label="تحميل PowerPoint (.pptx)",
                    data=ppt_data,
                    file_name="Interactive_Presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.info("تصدير PowerPoint غير متوفر حالياً.")
            
        with col3:
            txt_data = create_txt_file(st.session_state.adapted_text)
            st.download_button(
                label="تحميل نصي (.txt)",
                data=txt_data,
                file_name="Adapted_Worksheet.txt",
                mime="text/plain"
            )

        st.markdown("---")
        st.markdown("""
            <div class="animated-box" style="background-color: rgba(241, 196, 15, 0.15); border: 2px solid #F1C40F; padding: 20px; border-radius: 12px; text-align: center; margin-top: 20px; box-shadow: 0px 4px 15px rgba(241, 196, 15, 0.2);">
                <h3 style="margin: 0; font-weight: 900; line-height: 1.6;">شكراً لاستخدامك برنامج Edu Worksheet Adapt</h3>
                <h4 style="margin: 8px 0 0 0; font-weight: 900; line-height: 1.6;">Thank you for using Edu Worksheet Adapt</h4>
            </div>
        """, unsafe_allow_html=True)
