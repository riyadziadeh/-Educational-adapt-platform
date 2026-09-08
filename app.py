import io
import os
import streamlit as st
from docx import Document
from fpdf import FPDF
from pptx import Presentation

# إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="نظام تكييف أوراق العمل التربوية / Educational Worksheet Adaptation System",
    page_icon="📚",
    layout="centered",
)

# اختيار لغة الواجهة (العربية / الإنجليزية)
lang = st.sidebar.selectbox(
    "Choose Language / اختر اللغة", ["العربية", "English"]
)

if lang == "العربية":
  t_title = "نظام تكييف أوراق العمل التربوية"
  t_subtitle = "Educational Worksheet Adaptation System"
  t_subject = "المبحث / المادة الدراسية:"
  t_topic = "عنوان الدرس أو المهارة:"
  t_level = "مستوى الطالب / الفئة المستهدفة:"
  t_levels_list = [
      "تكييف لذوي الإعاقة الفكرية البسيطة",
      "تكييف لاضطراب طيف التوحد",
      "تكييف لصعوبات التعلم",
      "تكييف للدمج الشامل",
  ]
  t_canva_label = "أدخل رابط تصميم أو قالب Canva (اختياري):"
  t_prezi_label = "أدخل رابط عرض Prezi (اختياري):"
  t_btn = "تكييف ورقة العمل بالذكاء الاصطناعي"
  t_spinner = "جاري تكييف المادة وتوليد الملفات التربوية..."
  t_success = "تم تكييف ورقة العمل بنجاح!"
  t_downloads = "تحميل الملفات المطورة / Download Adapted Files:"
  t_feedback = "ملاحظات التعزيز والأداء / Feedback & Performance Notes:"
  cb_1 = "أتقن المهارة بمساعدة لمسية / Mastered with tactile support"
  cb_2 = "أتقن المهارة باستقلالية / Mastered independently"
  cb_3 = "يحتاج إلى إعادة توجيه / Needs prompt"
else:
  t_title = "Educational Worksheet Adaptation System"
  t_subtitle = "نظام تكييف أوراق العمل التربوية"
  t_subject = "Subject / Course Material:"
  t_topic = "Lesson Title or Skill:"
  t_level = "Student Level / Target Group:"
  t_levels_list = [
      "Mild Intellectual Disability Adaptation",
      "Autism Spectrum Disorder Adaptation",
      "Learning Difficulties Adaptation",
      "Inclusive Education Adaptation",
  ]
  t_canva_label = "Enter Canva Design/Template URL (Optional):"
  t_prezi_label = "Enter Prezi Presentation URL (Optional):"
  t_btn = "Adapt Worksheet with AI"
  t_spinner = "Adapting material and generating educational files..."
  t_success = "Worksheet adapted successfully!"
  t_downloads = "Download Adapted Files / تحميل الملفات المطورة:"
  t_feedback = "Feedback & Performance Notes / ملاحظات التعزيز والأداء:"
  cb_1 = "Mastered with tactile support / أتقن المهارة بمساعدة لمسية"
  cb_2 = "Mastered independently / أتقن المهارة باستقلالية"
  cb_3 = "Needs prompt / يحتاج إلى إعادة توجيه"

# عرض العناوين
st.markdown(
    f"<h1 style='text-align: center; color: #2563eb;'>{t_title}</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    f"<h3 style='text-align: center; color: #64748b;'>{t_subtitle}</h3>",
    unsafe_allow_html=True,
)
st.write("---")

# المدخلات الأساسية
subject = st.text_input(t_subject)
topic = st.text_input(t_topic)
student_level = st.selectbox(t_level, t_levels_list)

# روابط المنصات الخارجية (Canva & Prezi)
st.write("---")
st.subheader(
    "🎨 التكامل البصري والتفاعلي / Visual & Interactive Integration"
)
canva_url = st.text_input(t_canva_label, placeholder="https://www.canva.com/...")
prezi_url = st.text_input(t_prezi_label, placeholder="https://prezi.com/...")

# زر المعالجة والتكييف
if st.button(t_btn):
  if subject and topic:
    with st.spinner(t_spinner):
      try:
        # توليد المحتوى
        adapted_content = f"""
        Worksheet Adaptation Report / تقرير تكييف ورقة العمل:
        - Subject / المبحث: {subject}
        - Skill / المهارة: {topic}
        - Level / الفئة المستهدفة: {student_level}
        
        1. Adapted Learning Objective / الهدف التعليمي المعدل: Simplify core concepts into smaller steps.
        2. Strategies & Tools / الاستراتيجيات والوسائل: Visual supports, tactile cues, and structured guidance.
        3. Exercises / الأسئلة والتمارين: Custom tailored exercises matching individual performance levels.
        """

        st.success(t_success)
        st.write(adapted_content)

        # عرض روابط Canva و Prezi إذا وجدت
        if canva_url:
          st.info("Canva Integration / تم ربط تصميم Canva:")
          st.markdown(
              f"🔗 [Open Canva Template / فتح قالب Canva]({canva_url})"
          )

        if prezi_url:
          st.info("Prezi Integration / تم ربط عرض Prezi:")
          st.markdown(
              f"🔗 [Open Prezi Presentation / فتح عرض Prezi]({prezi_url})"
          )

        # 1. توليد ملف Word (.docx)
        doc = Document()
        doc.add_heading("Educational Worksheet Adaptation System", 0)
        doc.add_paragraph(adapted_content)
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # 2. توليد ملف PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(
            0,
            10,
            txt=(
                "Worksheet Adaptation Report\n"
                f"Subject: {subject}\n"
                f"Topic: {topic}\n"
                f"Level: {student_level}\n\n"
                "1. Objective: Simplified core concepts.\n"
                "2. Strategy: Visual and structured support.\n"
                "3. Exercises: Tailored to student needs."
            ),
        )
        pdf_bytes = pdf.output(dest="S").encode("latin1")
        pdf_io = io.BytesIO(pdf_bytes)

        # 3. توليد ملف PowerPoint (.pptx)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = topic
        slide.placeholders[1].text = adapted_content
        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        st.write("---")
        st.subheader(t_downloads)

        # أزرار التحميل المباشر لكافة الصيغ
        st.download_button(
            label="📄 تحميل Word (.docx)",
            data=doc_io,
            file_name="Adapted_Worksheet.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

        st.download_button(
            label="📑 تحميل PDF (.pdf)",
            data=pdf_io,
            file_name="Adapted_Worksheet.pdf",
            mime="application/pdf",
        )

        st.download_button(
            label="📊 تحميل PowerPoint (.pptx)",
            data=pptx_io,
            file_name="Adapted_Presentation.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )

      except Exception as e:
        st.error(f"Error occurred / حدث خطأ: {e}")
  else:
    st.warning(
        "يرجى إدخال المبحث وعنوان الدرس أولاً / Please enter subject and lesson"
        " title."
    )

# قسم الملاحظات والتعزيز أسفل الصفحة
st.write("---")
st.markdown(f"### {t_feedback}")
st.checkbox(cb_1)
st.checkbox(cb_2)
st.checkbox(cb_3)
