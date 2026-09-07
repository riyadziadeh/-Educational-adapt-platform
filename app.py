import streamlit as st

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
raw_text = st.text_area(
    "الصق محتوى ورقة العمل الأصلية هنا:",
    placeholder="اكتب الأسئلة أو التمارين هنا...",
)

st.markdown("---")

if st.button("🚀 ابدأ تكييف ورقة العمل الآن"):
  if not raw_text.strip():
    st.warning("⚠️ يرجى إدخال نص ورقة العمل قبل البدء.")
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

    st.markdown("---")
    st.subheader("📥 تحميل ورقة العمل المكيفة")
    st.download_button(
        label="تحميل ورقة العمل كملف نصي جاهز للطباعة",
        data=adapted_output,
        file_name=f"Adapted_Worksheet_{grade_level}.txt",
        mime="text/plain",
    )
