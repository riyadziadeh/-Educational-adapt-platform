import io
import os
import streamlit as st
from docx import Document
from pptx import Presentation

# إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="نظام تكييف أوراق العمل التربوية", page_icon="📚", layout="centered"
)

# عنوان التطبيق الواجهة
st.markdown(
    "<h1 style='text-align: center; color: #2563eb;'>نظام تكييف أوراق العمل"
    " التربوية</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<h3 style='text-align: center; color: #64748b;'>Educational Worksheet"
    " Adaptation System</h3>",
    unsafe_allow_html=True,
)
st.write("---")

# مدخلات المستخدم الأساسية
subject = st.text_input("المبحث / المادة الدراسية:")
topic = st.text_input("عنوان الدرس أو المهارة:")
student_level = st.selectbox(
    "مستوى الطالب / الفئة المستهدفة:",
    [
        "تكييف لذوي الإعاقة الفكرية البسيطة",
        "تكييف لاضطراب طيف التوحد",
        "تكييف لصعوبات التعلم",
        "تكييف للدمج الشامل",
    ],
)

# قسم خاص بتكامل Canva
st.write("---")
st.subheader("🎨 تصميم وتصدير عبر Canva")
canva_template_url = st.text_input(
    "أدخل رابط قالب Canva التفاعلي (اختياري):",
    placeholder="https://www.canva.com/design/...",
)

# زر المعالجة والتكييف
if st.button("تكييف ورقة العمل بالذكاء الاصطناعي"):
  if subject and topic:
    with st.spinner("جاري تكييف المادة وتوليد الملفات التربوية..."):
      try:
        # محاكاة وتوليد المحتوى المكيّف بشكل آمن نصياً ورياضياً
        adapted_content = f"""
        تقرير تكييف ورقة العمل التربوية:
        - المبحث: {subject}
        - المهارة: {topic}
        - الفئة المستهدفة: {student_level}
        
        1. الهدف التعليمي المعدل: أن يتعرف الطالب على المفاهيم الأساسية بطريقة مبسطة ومجزأة.
        2. الاستراتيجيات والوسائل: استخدام الدعم البصري واللمسي، وتعزيز الاستقلالية.
        3. الأسئلة والتمارين: تم تكييف الأسئلة لتناسب مستوى الأداء الفردي مع توفير مساحات إجابة واضحة.
        """

        st.success("تم تكييف ورقة العمل بنجاح!")
        st.write(adapted_content)

        # إذا قام المستخدم بإدخال رابط Canva، نقوم بعرضه بشكل تفاعلي
        if canva_template_url:
          st.info(
              "تم ربط المحتوى بنجاح مع تصميم Canva الخاص بك لتسهيل العرض"
              " البصري للطلاب."
          )
          st.markdown(
              f"🔗 [اضغط هنا لفتح قالب Canva وتعديله مباشرة]"
              f"({canva_template_url})"
          )

        # 1. توليد ملف Word بشكل آمن (بصيغة بايتات Bytes لتجنب أي أخطاء)
        doc = Document()
        doc.add_heading("نظام تكييف أوراق العمل التربوية", 0)
        doc.add_paragraph(adapted_content)

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # 2. توليد ملف PowerPoint بشكل آمن
        prs = Presentation()
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = topic
        slide.placeholders[1].text = adapted_content

        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)

        st.write("---")
        st.subheader("تحميل الملفات المطورة / Download Adapted Files:")

        # أزرار التحميل الآمنة
        st.download_button(
            label="تحميل Word (.docx)",
            data=doc_io,
            file_name="Adapted_Worksheet.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

        st.download_button(
            label="تحميل PowerPoint (.pptx)",
            data=pptx_io,
            file_name="Adapted_Presentation.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )

      except Exception as e:
        st.error(f"حدث خطأ أثناء الاتصال أو معالجة الملفات: {e}")
  else:
    st.warning("يرجى إدخال المبحث وعنوان الدرس أولاً.")

# قسم الملاحظات والأداء أسفل الصفحة
st.write("---")
st.markdown("### Feedback / ملاحظات التعزيز والأداء:")
st.checkbox("Mastered with tactile support / أتقن المهارة بمساعدة لمسية")
st.checkbox("Mastered independently / أتقن المهارة باستقلالية")
st.checkbox("Needs prompt / يحتاج إلى إعادة توجيه")
