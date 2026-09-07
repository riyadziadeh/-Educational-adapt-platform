import io
import google.generativeai as genai
import pypdf
from docx import Document
from pptx import Presentation
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

# سحب المفتاح من أسرار المنصة السحابية تلقائياً
api_key = None
try:
  if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
  pass

st.title("📚 المنصة الذكية لتكييف وتبسيط أوراق العمل")
st.markdown(
    "**النظام الوطني المتقدم لتكييف المناهج المدرسية المدعوم بالذكاء الاصطناعي.**"
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
          "صعوبة حركية كتابية (Dysgraphia)",
      ],
  )

st.subheader("📝 مدخلات ورقة العمل")
input_method = st.radio(
    "اختر طريقة إدخال ورقة العمل:",
    ["كتابة أو لصق النص مباشرة", "رفع ملف PDF من الجهاز"],
)

extracted_text = ""

if input_method == "كتابة أو لصق النص مباشرة":
  extracted_text = st.text_area(
      "الصق محتوى ورقة العمل الأصلية هنا:",
      placeholder="اكتب الأسئلة أو التمارين هنا ليتم تكييفها بالذكاء الاصطناعي...",
      height=150,
  )
else:
  uploaded_file = st.file_uploader("اختر ملف PDF من جهازك:", type=["pdf"])
  if uploaded_file is not None:
    try:
      reader = pypdf.PdfReader(uploaded_file)
      for page in reader.pages:
        text = page.extract_text()
        if text:
          extracted_text += text + "\n"
      st.success(f"تم قراءة نص الملف ({uploaded_file.name}) بنجاح!")
    except Exception as e:
      st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

st.markdown("---")

if st.button("🚀 ابدأ تكييف ورقة العمل بالذكاء الاصطناعي"):
  if not api_key:
    st.error(
        "⚠️ يرجى إضافة المفتاح مؤقتاً في إعدادات المنصة (Secrets) ليعمل النظام"
        " بسلاسة."
    )
  elif not extracted_text.strip():
    st.warning("⚠️ يرجى إدخال النص أو رفع ملف ورقة العمل أولاً.")
  else:
    with st.spinner(
        "🤖 يقوم الذكاء الاصطناعي الآن بقراءة وتحليل وتكييف ورقة العمل... الرجاء"
        " الانتظار..."
    ):
      try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.5-flash")

        prompt = f"""
        أنت خبير تربوي متخصص في التربية الخاصة وتكييف المناهج.
        المطلوب منك تكييف وتبسيط ورقة العمل التالية لتناسب قدرات طالب في {grade_level} يعاني من: {student_condition}.
        المادة الدراسية: {subject}.
        لغة الإخراج المطلوبة: {language}.
        
        النص الأصلي لورقة العمل:
        {extracted_text}
        
        التعليمات الصارمة:
        1. أعد صياغة الأسئلة لتكون واضحة، مباشرة، ومناسبة تماماً لحالة الطالب.
        2. استخدم أسلوباً مشجعاً وهادئاً في صياغة التعليمات.
        3. قسم الأسئلة الطويلة أو المعقدة إلى خطوات بسيطة ومحددة.
        4. أعطني فقط ورقة العمل النهائية المكيفة والمنسقة بشكل احترافي وجاهز للطباعة.
        """

        response = model.generate_content(prompt)
        adapted_output = response.text

        st.success("✨ تمت عملية التكييف الذكي بنجاح!")
        st.markdown("### 📄 ورقة العمل المكيفة بالذكاء الاصطناعي:")
        st.info(adapted_output)

        st.markdown("---")
        st.subheader("📥 تحميل ورقة العمل المكيفة بجميع الصيغ")
        col1, col2, col3 = st.columns(3)

        with col1:
          st.download_button(
              label="📄 تحميل نصي (TXT)",
              data=adapted_output,
              file_name=f"AI_Worksheet_{grade_level}.txt",
              mime="text/plain",
          )

        with col2:
          doc = Document()
          doc.add_heading("ورقة عمل مكيفة (بالذكاء الاصطناعي)", 0)
          doc.add_paragraph(adapted_output)
          doc_io = io.BytesIO()
          doc.save(doc_io)
          doc_io.seek(0)
          st.download_button(
              label="📝 تحميل Word",
              data=doc_io,
              file_name=f"AI_Worksheet_{grade_level}.docx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              ),
          )

        with col3:
          prs = Presentation()
          slide_layout = prs.slide_layouts[1]
          slide = prs.slides.add_slide(slide_layout)
          slide.shapes.title.text = f"ورقة عمل مكيفة - {grade_level}"
          slide.placeholders[1].text = (
              adapted_output[:700] + "\n...(اقرأ الباقي في ملف الوورد)"
          )
          ppt_io = io.BytesIO()
          prs.save(ppt_io)
          ppt_io.seek(0)
          st.download_button(
              label="📊 تحميل PowerPoint",
              data=ppt_io,
              file_name=f"AI_Worksheet_{grade_level}.pptx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.presentationml.presentation"
              ),
          )

      except Exception as e:
        st.error(f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي: تفاصيل: {e}")
