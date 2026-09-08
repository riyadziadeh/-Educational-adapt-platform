import os
import io
import streamlit as st
import google.generativeai as genai
from docx import Document
from pptx import Presentation
from fpdf import FPDF
import pypdf

# إعداد صفحة ستريمليت
st.set_page_config(page_title="تكييف أوراق العمل بالذكاء الاصطناعي", layout="centered")

st.title("📚 نظام تكييف أوراق العمل التربوية")
st.write("قم برفع ملف ورقة العمل وسيتم تحليلها وتكييفها تلقائياً مع خيارات التحميل المتعددة.")

# جلب مفتاح الـ API بأمان
api_key = None
try:
    api_key = st.secrets["GOOGLE_API_KEY"]
except:
    pass

if not api_key:
    api_key = st.text_input("أدخل مفتاح Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    MODEL_NAME = "gemini-3.6-flash"

    grades = [
        "الصف الأول", "الصف الثاني", "الصف الثالث", "الصف الرابع", 
        "الصف الخامس", "الصف السادس", "الصف السابع", "الصف الثامن", "الصف التاسع"
    ]

    educational_systems = [
        "وطني (National)", 
        "دولي (International)"
    ]

    jordan_governorates = [
        "العاصمة (عمان)", "إربد", "الزرقاء", "البلقاء", "المفرق", 
        "الكرك", "مادبا", "جرش", "عجلون", "معان", "الطفيلة", "العقبة"
    ]

    special_conditions_categories = {
        "1. الإعاقات الحسية والجسدية": [
            "الإعاقة البصرية (كف تام أو ضعف بصر شديد / بريل ومطبوعات كبيرة)",
            "الإعاقة السمعية (صمم تام أو ضعف سمعي بحاجة لمعينات/إشارة)",
            "الإعاقة الحركية أو الجسدية (شلل، ضمور عضلات، بتر أطراف، تشوهات)",
            "الإعاقة الحسية المزدوجة (الصم-المكفوفين)"
        ],
        "2. الاضطرابات النمائية وصعوبات التعلم": [
            "صعوبات التعلم المحددة (ديسليكسيا، عسر كتابة، صعوبة حساب)",
            "اضطراب طيف التوحد (ASD)",
            "اضطراب فرط الحركة ونقص الانتباه (ADHD)",
            "اضطرابات النطق واللغة والتواصل (تأتأة، عيوب نطق)"
        ],
        "3. الإعاقات الذهنية والسلوكية": [
            "الإعاقة الذهنية / العقلية (بسيطة، متوسطة، شديدة)",
            "الاضطرابات الانفعالية والسلوكية (قلق شديد، اكتئاب، مخاوف مدرسية)",
            "الإعاقات المتعددة (أكثر من إعاقة معاً)"
        ],
        "4. الإعاقات والحالات الصحية المزمنة": [
            "الأمراض المزمنة المحتاجة لمتابعة (سكري، ربو شديد، صرع، أمراض قلب)",
            "مرضى السرطان (برامج استكمال وعلاجات مستمرة)",
            "حالات الفشل الكلوي (غسيل دوري)"
        ],
        "5. فئة الموهبة والتفوق": [
            "الطلبة الموهوبون والمتفوقون (برامج إثراء معرفي وتسريع أكاديمي)"
        ]
    }

    selected_grade = st.selectbox("اختر الصف الدراسي:", grades)
    selected_system = st.selectbox("اختر النظام التعليمي:", educational_systems)
    selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن:", jordan_governorates)
    
    selected_category = st.selectbox("اختر فئة الحالة الخاصة:", list(special_conditions_categories.keys()))
    selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة:", special_conditions_categories[selected_category])

    uploaded_file = st.file_uploader("قم بتمرير أو رفع ملف ورقة العمل (PDF أو Word أو TXT):", type=["pdf", "docx", "txt"])

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
        
        st.success(f"تم قراءة الملف بنجاح: {uploaded_file.name}")

    # دوال توليد الملفات المتوافقة مع اللغة العربية والنصوص الدقيقة
    def create_word_file(text):
        doc = Document()
        doc.add_heading('ورقة العمل المطورة (التربية الخاصة)', 0)
        for line in text.split('\n'):
            doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    def create_ppt_file(text):
        prs = Presentation()
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "ورقة العمل المطورة"
        subtitle.text = text[:300] + "..." if len(text) > 300 else text
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio

    def create_pdf_file(text):
        pdf = FPDF()
        pdf.add_page()
        # استخدام خط عام متوافق وتجنب التشفير الخاطئ
        pdf.set_font("Arial", size=11)
        pdf.set_auto_page_break(auto=True, margin=15)
        # تنظيف النص لضمان توافق التحميل
        clean_text = text.encode('latin-1', 'ignore').decode('latin-1')
        for line in clean_text.split('\n'):
            if line.strip():
                pdf.multi_cell(0, 8, line)
            else:
                pdf.ln(4)
        bio = io.BytesIO(pdf.output(dest='S'))
        bio.seek(0)
        return bio

    if st.button("ابدأ تكييف ورقة العمل بالذكاء الاصطناعي 🚀"):
        if not extracted_content.strip():
            st.warning("الرجاء رفع ملف ورقة العمل أولاً ليتم استخراج محتواه.")
        else:
            with st.spinner("جاري معالجة ورقة العمل وتكييفها عبر الذكاء الاصطناعي..."):
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج في الأردن (كلية دي لاسال / تراسنطة). يرجى تكييف وتطوير ورقة العمل التالية بدقة فائقة:
                - الصف الدراسي: {selected_grade}
                - النظام التعليمي: {selected_system}
                - موقع المدرسة (المحافظة): {selected_gov} - الأردن
                - التصنيف والحالة الخاصة: {selected_category} -> {selected_condition}
                
                محتوى ورقة العمل المستخرج من الملف:
                {extracted_content}
                
                يرجى إعادة صياغة ورقة العمل وتنظيمها بطريقة تربوية احترافية تراعي الفروق الفردية والخصائص المذكورة بدقة تامة، وإظهار دليل المعلم وإرشادات الدمج بوضوح.
                """
                try:
                    model = genai.GenerativeModel(MODEL_NAME)
                    response = model.generate_content(prompt)
                    adapted_text = response.text
                    
                    st.success("تم تكييف ورقة العمل بنجاح تام!")
                    st.markdown("### ورقة العمل المطورة:")
                    st.markdown(adapted_text)
                    
                    st.markdown("---")
                    st.subheader("📥 تحميل ورقة العمل المطورة:")
                    
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
                            label="تحميل PowerPoint (.pptx)",
                            data=ppt_data,
                            file_name="Adapted_Worksheet.pptx",
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

                except Exception as e:
                    st.error(f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي: {str(e)}")
else:
    st.info("الرجاء إدخال مفتاح الـ API الخاص بك في الأعلى لتشغيل التطبيق.")