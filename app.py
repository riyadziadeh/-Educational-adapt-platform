import os
import streamlit as st

# إعدادات صفحة التطبيق والواجهة الفاخرة
st.set_page_config(
    page_title="المنصة الذكية لتكييف أوراق العمل",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# تصميم تنسيقات CSS مخصصة لواجهة راقية ونظيفة
st.markdown(
    """
    <style>
    .main {
        background-color: #f8f9fa;
        direction: rtl;
        text-align: right;
    }
    .stButton>button {
        background-color: #0d6efd;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #0b5ed7;
    }
    h1, h2, h3 {
        color: #1e293b;
        font-family: 'Cairo', sans-serif;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# العنوان الرئيسي ووصف المنصة
st.title("📚 المنصة الذكية لتكييف وتبسيط أوراق العمل")
st.markdown(
    "**مرحباً بك في النظام الوطني المتقدم لتكييف المناهج المدرسية (الصف 1 - 9) للطلاب ذوي الحالات الخاصة والحسّاسة.**"
)
st.markdown("---")

# الشريط الجانبي لإدخال بيانات المعلم والورقة
with st.sidebar:
  st.header("⚙️ إعدادات ورقة العمل")

  # اختيار لغة ورقة العمل
  language = st.selectbox(
      "لغة ورقة العمل الأساسية:", ["اللغة العربية", "اللغة الإنجليزية"]
  )

  # اختيار الصف الدراسي (من الأول للتاسع)
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

  # اختيار المادة الدراسية
  subject = st.text_input(
      "المادة الدراسية (مثال: رياضيات، علوم، لغة عربية):"
  )

  st.markdown("---")
  st.header("🧩 حالة الطالب المستهدف")

  # تصنيف الحالات الخاصة الدقيقة التي اعتمدناها
  student_condition = st.selectbox(
      "اختر التحدي أو الحالة الخاصة:",
      [
          "صعوبات تعلم (Dyslexia / Dyscalculia / عسر قراءة وكتابة)",
          "تشتت ذهني وفرط حركة (ADHD & Executive Dysfunction)",
          "هلع وفزع دراسي (Academic Anxiety)",
          "اختلاف لغة الدارس (ELL - لغة أجنبية)",
          "مشاكل نطقية أو سمعية (Speech/Hearing Impairment)",
          "تأخر نمائي / بطء تعلم (Borderline Intellectual Functioning)",
          "إجهاد صحي مزمن (Chronic Fatigue / Medical Condition)",
          "صعوبة حركية كتابية (Dysgraphia)",
      ],
  )

# مساحة المحتوى الرئيسي لرفع أو كتابة ورقة العمل
col1, col2 = st.columns([1, 1])

with col1:
  st.subheader("📝 مدخلات ورقة العمل")
  input_method = st.radio(
      "طريقة الإدخال:", ["كتابة أو لصق النص", "رفع ملف (PDF / صورة)"]
  )

  worksheet_text = ""
  if input_method == "كتابة أو لصق النص":
    worksheet_text = st.text_area(
        "الصق محتوى ورقة العمل هنا:",
        height=250,
        placeholder="اكتب أو الصق الأسئلة والتمارين الأصلية هنا...",
    )
  else:
    uploaded_file = st.file_uploader(
        "ارفع ملف ورقة العمل (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"]
    )
    if uploaded_file is not None:
      st.success("تم رفع الملف بنجاح وجاهز للتحليل!")
      worksheet_text = (
          "[تم إرفاق ملف خارجي يحتوي على التمارين والأسئلة الخاصة بالدرس]"
      )

  process_button = st.button("🚀 تكييف ورقة العمل الآن")

with col2:
  st.subheader("🎯 المخرجات التربوية المخصصة")

  if process_button:
    if not worksheet_text.strip():
      st.warning(
          "الرجاء إدخال نص ورقة العمل أو رفع الملف قبل الضغط على زر التكييف."
      )
    else:
      with st.spinner("جاري تحليل المعايير، مطابقة المنهاج، وتكييف الأسئلة..."):

        # محاكاة عمل محرك الذكاء الاصطناعي بناءً على البروموت الذي وضعناه مسبقاً
        # (في النسخة المكتملة، سيتم إرسال هذا الطلب مباشرة عبر API الخاص بالنموذج الذكي)

        st.success("تم تكييف ورقة العمل بنجاح وفق معايير وزارة التربية والتعليم!")

        st.markdown("### 📄 1. ورقة العمل المبسطة والمعدلة للطالب:")
        st.info(
            f"*(مخصصة لـ: {grade_level} - المادة: {subject} - الحالة: {student_condition})*\n\n"
            "• **تعليمات هادئة للطالب:** خذ وقتك بالحل، كل سؤال سنحله خطوة بخطوة معاً.\n"
            "• **السؤال الأول (مبسط ومباشر):** [تم تعديل شكل السؤال ليكون مناسباً لقدرات الطالب، مع تقسيم الجمل الطويلة وتقليل الحمل البصري].\n"
            "• **المهام المعرفية المقسمة (Checklist):**\n"
            "  ☐ الخطوة الأولى\n"
            "  ☐ الخطوة الثانية"
        )

        st.markdown("### 👩‍🏫 2. دليل المعلم التربوي السريع:")
        st.write(
            "• تم تبسيط الصياغة لتجنب إثارة التوتر أو التشتت البصري.\n"
            "• ينصح بتقديم الدعم اللفظي البسيط أثناء إجابة الطالب على السؤال الأول لكسر حاجز الخوف."
        )

        st.markdown("### 📊 3. مؤشر تقييم الأثر المبدئي:")
        st.write(
            "*(سؤال للمعلم بعد الحصة)*: هل شعرت أن الطالب أنجز المهمة بسرعة أكبر ودون توتر مقارنة بالأوراق العادية؟ (نعم / لا / بحاجة لتعديل إضافي)."
        )

