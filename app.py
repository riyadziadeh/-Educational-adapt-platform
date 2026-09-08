import os
import io
import streamlit as st
import google.generativeai as genai
from docx import Document
from pptx import Presentation
from fpdf import FPDF

# إعداد صفحة ستريمليت
st.set_page_config(page_title="تكييف أوراق العمل بالذكاء الاصطناعي", layout="centered")

st.title("📚 نظام تكييف أوراق العمل التربوية")
st.write("قم بإدخال تفاصيل ورقة العمل وسيتم تكييفها تلقائياً مع خيارات التحميل المتعددة.")

# إدخال مفتاح الـ API
api_key = st.text_input("أدخل مفتاح Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    MODEL_NAME = "gemini-2.5-flash"

    # القوائم والخيارات الشاملة
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

    # واجهة الإدخال في التطبيق
    selected_grade = st.selectbox("اختر الصف الدراسي:", grades)
    selected_system = st.selectbox("اختر النظام التعليمي:", educational_systems)
    selected_gov = st.selectbox("اختر محافظة المدرسة في الأردن:", jordan_governorates)
    
    selected_category = st.selectbox("اختر فئة الحالة الخاصة:", list(special_conditions_categories.keys()))
    selected_condition = st.selectbox("اختر الحالة التشخيصية المحددة:", special_conditions_categories[selected_category])

    original_content = st.text_area("الصق محتوى ورقة العمل الأصلية هنا:")

    # دوال توليد الملفات للتحميل
    def create_word_file(text):
        doc = Document()
        doc.add_heading('ورقة العمل المطورة (التربية الخاصة)', 0)
        doc.add_paragraph(text)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio

    def create_ppt_file(text):
        prs = Presentation()
        slide_layout = prs.slide_layouts[1] # شريحة عنوان ومحتوى
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "ورقة العمل المطورة"
        # تقصير النص للشريحة إذا كان طويلاً
        subtitle.text = text[:500] + "..." if len(text) > 500 else text
        
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio

    def create_pdf_file(text):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        # معالجة النصوص العربية المبسطة للـ PDF الافتراضي
        safe_text = text.encode('latin-1', 'replace').decode('latin-1')
        for line in safe_text.split('\n'):
            pdf.multi_cell(0, 10, line)
        
        bio = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
        bio.seek(0)
        return bio

    if st.button("ابدأ تكييف ورقة العمل بالذكاء الاصطناعي 🚀"):
        if not original_content.strip():
            st.warning("الرجاء إدخال محتوى ورقة العمل الأصلية أولاً.")
        else:
            with st.spinner("جاري معالجة ورقة العمل وتكييفها..."):
                prompt = f"""
                أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج. يرجى تكييف وتطوير ورقة العمل التالية بدقة عالية:
                - الصف الدراسي: {selected_grade}
                - النظام التعليمي: {selected_system}
                - موقع المدرسة (المحافظة): {selected_gov} - الأردن
                - التصنيف والحالة الخاصة: {selected_category} -> {selected_condition}
                
                محتوى ورقة العمل الأصلية:
                {original_content}
                
                يرجى إعادة صياغة ورقة العمل وتنظيمها بطريقة تربوية احترافية تراعي الفروق الفردية والخصائص المذكورة بدقة تامة.
                """
                try:
                    model = genai.GenerativeModel(MODEL_NAME)
                    response = model.generate_content(prompt)
                    adapted_text = response.text
                    
                    st.success("تم تكييف ورقة العمل بنجاح!")
                    st.markdown("### ورقة العمل المطورة:")
                    st.markdown(adapted_text)
                    
                    # خيارات التحميل المتعددة
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
