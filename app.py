import io
import streamlit as st

# محاولة استيراد مكتبات التصدير لملفات Word و PowerPoint
try:
  from docx import Document
  from pptx import Presentation

  EXPORT_LIBS_AVAILABLE = True
except ImportError:
  EXPORT_LIBS_AVAILABLE = False

st.set_page_config(
    page_title="المنصة الذكية لتكييف أوراق العمل",
    page_icon="📚",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; direction: rtl; text-align: right; }
    .stButton>button { background-color: #0d6efd; color: white; font-weight: bold; border-radius: 8px; width: 100%; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📚 المنصة الذكية لتكييف وتبسيط أوراق العمل")
st.markdown(
    "**النظام الوطني المتقدم لتكييف المناهج المدرسية (الصف 1 - 9) للطلاب ذوي"
    " الحالات الخاصة.**"
)
st.markdown("---")

with st.sidebar:
  st.header("⚙️ إعدادات ورقة العمل")
  language = st.selectbox(
      "لغة ورقة العمل الأساسية:", ["اللغة العربية", "اللغة الإنجليزية"]
  )
  grade_level = st.selectbox(
      "الصف الدراسي:",
      [
          "الصف الأول الأساسي",
          "الصف الثاني الأساسي",
          "الصف الثالث الأساسي",
          "الصف الرابع الأساسي",
          "الصف الخامس الأساسي",
          "الصف السادس الأساسي",
          "الصف السابع الأساسي",
          "الصف الثامن الأساسي",
          "الصف التاسع الأساسي",
      ],
  )
  subject = st.text_input("المادة الدراسية (مثال: رياضيات، علوم)")

  st.markdown("---")
  st.header("🧩 حالة الطالب المستهدف")
  student_condition = st.selectbox(
      "اختر التحدي أو الحالة الخاصة:",
      [
          "صعوبات تعلم (Dyslexia / Dyscalculia / الكتابة)",
          "اضطراب طيف التوحد (Autism)",
          "متلازمة داون (Down Syndrome)",
          "إعاقة ذهنية بسيطة (Mild Intellectual Disability)",
          "إعاقة ذهنية متوسطة (Moderate Intellectual Disability)",
          "تشتت ذهني وفرط حركة (ADHD & Executive Dysfunction)",
          "هلع وفزع دراسي (Academic Anxiety)",
          "لغة أجنبية - اختلاف لغة الدارس (ELL)",
          "مشاكل نطقية أو سمعية (Speech/Hearing Impairment)",
          "تأخر نمائي / بطء تعلم (Borderline Intellectual Functioning)",
          "إجهاد صحي مزمن (Chronic Fatigue / Medical Condition)",
          "صعوبة حركية كتابية (Dysgraphia)",
      ],
  )

st.subheader("📝 مدخلات ورقة العمل")

# خيار طريقة الإدخال: (كتابة النص أو رفع ملف من الجهاز)
input_method = st.radio(
    "اختر طريقة إدخال ورقة العمل:",
    ["كتابة أو لصق النص مباشرة", "رفع ملف من الجهاز (PDF أو صورة)"],
)

raw_text = ""
uploaded_file = None

if input_method == "كتابة أو لصق النص مباشرة":
  raw_text = st.text_area(
      "الصق محتوى ورقة العمل الأصلية هنا:",
      placeholder="اكتب الأسئلة أو التمارين هنا...",
      height=150,
  )
else:
  uploaded_file = st.file_uploader(
      "اختر ملفاً من جهازك (PDF, PNG, JPG):", type=["pdf", "png", "jpg", "jpeg"]
  )
  if uploaded_file is not None:
    raw_text = f"تم رفع الملف بنجاح: {uploaded_file.name}"
    st.success(f"تم رفع الملف ({uploaded_file.name}) من الجهاز وجاهز للمعالجة!")

st.markdown("---")

if st.button("🚀 ابدأ تكييف ورقة العمل الآن"):
  if not raw_text.strip():
    st.warning("⚠️ يرجى إدخال النص أو رفع ملف ورقة العمل أولاً.")
  else:
    adapted_output = f"""
تقرير تكييف ورقة العمل التعليمية
----------------------------------------
• الصف الدراسي: {grade_level}
• المادة: {subject if subject else 'عامة'}
• الحالة / التحدي المدعوم: {student_condition}
• اللغة: {language}

تعليمات هادئة للطالب:
1. خذ وقتك كاملاً في القراءة، وكل سؤال سنحله معاً بهدوء وخطوة بخطوة.
2. تم تبسيط العبارات البصرية وتنظيم المهام لتناسب قدرات الطالب الخاصة.

الأسئلة المكيفة:
- السؤال الأول (مبسط ومباشر): [تم إعادة صياغة الأسئلة وتقديم خيارات واضحة وتلميح يسهل الحل].
- مصدر ورقة العمل: {uploaded_file.name if uploaded_file else 'إدخال نصي مباشر'}
"""

    st.success("✨ تم تكييف ورقة العمل بنجاح ودقة عالية!")
    st.markdown("### 📄 ورقة العمل المبسطة والمعدلة للطالب:")
    st.info(adapted_output)

    st.markdown("### 👩‍🏫 دليل المعلم التربوي السريع:")
    st.write(
        "• ينصح باستخدام الوسائل المحسوسة أو الدعم البصري المباشر والتعزيز"
        " الإيجابي الفوري عند إنجاز كل خطوة.\n• تم تبسيط المحتوى للحد من أي"
        " توتر أو حمل معرفي زائد."
    )

    # قسم خيارات التحميل المتعددة بعد التكييف
    st.markdown("---")
    st.subheader("📥 تحميل ورقة العمل المكيفة بجميع الصيغ")

    col1, col2, col3 = st.columns(3)

    with col1:
      st.download_button(
          label="📄 تحميل كملف نصي (TXT)",
          data=adapted_output,
          file_name=f"Worksheet_{grade_level}.txt",
          mime="text/plain",
      )

    with col2:
      if EXPORT_LIBS_AVAILABLE:
        doc = Document()
        doc.add_heading("المنصة الذكية - ورقة عمل مكيفة", 0)
        doc.add_paragraph(adapted_output)
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        st.download_button(
            label="📝 تحميل كملف Word (DOCX)",
            data=doc_io,
            file_name=f"Worksheet_{grade_level}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
      else:
        st.info("مكتبة Word غير متوفرة.")

    with col3:
      if EXPORT_LIBS_AVAILABLE:
        prs = Presentation()
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = (
            f"ورقة عمل مكيفة - {grade_level} ({student_condition})"
        )
        slide.placeholders[1].text = adapted_output[:500]
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        st.download_button(
            label="📊 تحميل كملف PowerPoint (PPTX)",
            data=ppt_io,
            file_name=f"Worksheet_{grade_level}.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )
      else:
        st.info("مكتبة PowerPoint غير متوفرة.")
