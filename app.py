import os
import io
import streamlit as st
import google.generativeai as genai
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from fpdf import FPDF
import pypdf

# إعداد صفحة ستريمليت
st.set_page_config(page_title="تكييف أوراق العمل بالذكاء الاصطناعي | Educational Worksheet Adaptation Platform", layout="centered")

# تخصيص CSS لتلوين الخانات بالأصفر الهادئ وجعل الخطوط عريضة (Bold)
st.markdown("""
    <style>
    /* عناوين الحقول والقوائم بخط عريض وواضح جداً */
    .stSelectbox label p, .stFileUploader label p, div[data-baseweb="select"] label, label {
        font-weight: 900 !important;
        font-size: 17px !important;
        color: #1A252F !important;
    }
    
    /* خلفية خانات القوائم المنسدلة بلون أصفر بارد وهادئ */
    div[data-baseweb="select"] > div {
        background-color: #FFFDEB !important;
        border-radius: 10px !important;
        border: 2px solid #F1C40F !important;
    }

    /* خلفية خانة رفع الملفات بلون أصفر بارد وهادئ */
    div.stFileUploader > div {
        background-color: #FFFDEB !important;
        border-radius: 10px !important;
        border: 2px solid #F1C40F !important;
    }
    </style>
""", unsafe_allow_html=True)

# عرض الشعار الأفقي بجودة عالية وبحجم عريض في منتصف الصفحة تماماً
col_logo1, col_logo2, col_logo3 = st.columns([0.5, 3, 0.5])
with col_logo2:
    logo_loaded = False
    for filename in ["logo.png", "logo.jpg", "Logo.png", "Logo.JPG"]:
        if os.path.exists(filename):
            st.image(filename, use_container_width=True)
            logo_loaded = True
            break
    if not logo_loaded:
        st.warning("الرجاء التأكد من رفع صورة الشعار باسم logo.png في نفس مجلد المشروع.")

# العنوان الرئيسي للنظام تحت الشعار مباشرة
st.markdown("""
    <div style="text-align: center;">
        <h1 style="font-size: 28px; margin-bottom: 0; font-weight: 900; color: #2C3E50;">نظام تكييف أوراق العمل التربوية</h1>
        <h2 style="font-size: 22px; margin-top: 5px; color: #34495E; font-weight: 900;">Educational Worksheet Adaptation System</h2>
    </div>
""", unsafe_allow_html=True)

st.write("قم برفع ملف ورقة العمل وسيتم تحليلها وتكييفها تلقائياً باللغتين مع خيارات التحميل المتعددة.")

# كبسة تشغيل الموسيقى المباشرة (مع بحث تلقائي عن صيغ الملفات الصوتية الممكنة)
audio_file_path = None
for music_name in ["music.mp3", "Music.mp3", "MUSIC.MP3", "music.WAV", "music.ogg"]:
    if os.path.exists(music_name):
        audio_file_path = music_name
        break

if audio_file_path:
    st.audio(audio_file_path, format="audio/mp3")
else:
    # في حال لم يتم العثور على الملف محلياً، يتم عرض مشغل برابط بديل مؤقت لكي يظهر الزر تماماً ولا يتوقف التطبيق
    st.audio("https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3", format="audio/mp3")

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
    MODEL_NAME = "gemini-3.6-flash"

    # القوائم ثنائية اللغة بالكامل
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

    # واجهة الإدخال ثنائية اللغة
    selected_grade = st.selectbox("اختر الصف الدراسي / Select Grade:", grades)
    selected_system = st.selectbox("اختر النظام التعليمي / Select Educational System:", educational_systems)
    selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن / Select Governorate in Jordan:", jordan_governorates)
    
    selected_category = st.selectbox("اختر فئة الحالة الخاصة / Select Special Condition Category:", list(special_conditions_categories.keys()))
    selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة / Select Specific Condition:", special_conditions_categories[selected_category])

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
        bio = io.BytesIO(pdf.output(dest='S'))
        bio.seek(0)
        return bio

    if st.button("ابدأ تكييف ورقة العمل بالذكاء الاصطناعي 🚀 / Start AI Adaptation"):
        if not extracted_content.strip():
            st.warning("الرجاء رفع ملف ورقة العمل أولاً / Please upload a file first.")
        else:
            with st.spinner("جاري معالجة ورقة العمل وتكييفها باللغتين... / Processing..."):
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن (كلية دي لاسال / تراسنطة). 
                يرجى تكييف وتطوير ورقة العمل التالية بدقة فائقة مع توفير المصطلحات باللغتين العربية والإنجليزية:
                - الصف الدراسي / Grade: {selected_grade}
                - النظام التعليمي / System: {selected_system}
                - موقع المدرسة (المحافظة) / Governorate: {selected_gov} - الأردن
                - التصنيف والحالة الخاصة / Condition: {selected_category} -> {selected_condition}
                
                محتوى ورقة العمل المستخرج من الملف:
                {extracted_content}
                
                يرجى إعادة صياغة ورقة العمل وتنظيمها بطريقة تربوية احترافية ثنائية اللغة (عربي/إنجليزي) تراعي الفروق الفردية وإرشادات الدمج الشامل.
                """
                try:
                    model = genai.GenerativeModel(MODEL_NAME)
                    response = model.generate_content(prompt)
                    adapted_text = response.text
                    
                    st.success("تم تكييف ورقة العمل بنجاح تام / Adapted Successfully!")
                    st.markdown("### ورقة العمل المطورة ثنائية اللغة / Bilingual Adapted Worksheet:")
                    st.markdown(adapted_text)
                    
                    st.markdown("---")
                    st.subheader("📥 تحميل الملفات المطورة / Download Adapted Files:")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        word_data = create_word_file(adapted_text)
                        st.download_button(
                            label="تحميل Word (.docx)",
                            data=word_data,
                            file_name="Bilingual_Adapted_Worksheet.docx",
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
                            file_name="Bilingual_Adapted_Worksheet.pdf",
                            mime="application/pdf"
                        )

                except Exception as e:
                    st.error(f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي / Error: {str(e)}")
