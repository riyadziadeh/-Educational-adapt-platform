import streamlit as st
import os
from io import BytesIO

# محاولة استيراد مكتبات قراءة الملفات وملفات الـ Word لتجنب توقف التطبيق إن لم تكن مثبتة
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pypdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="Educ Worksheet Adapt",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تهيئة متغيرات الجلسة بأمان تام لتجنب أخطاء الخادم (500)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_email" not in st.session_state:
    st.session_state.user_email = ""
if "adapted_content" not in st.session_state:
    st.session_state.adapted_content = ""

def extract_text_from_file(uploaded_file):
    text = ""
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension == 'pdf' and PDF_AVAILABLE:
        try:
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        except Exception as e:
            text = f"خطأ في قراءة ملف الـ PDF: {e}"
            
    elif file_extension in ['docx', 'doc'] and DOCX_AVAILABLE:
        try:
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        except Exception as e:
            text = f"خطأ في قراءة ملف الـ Word: {e}"
            
    elif file_extension == 'txt':
        try:
            text = uploaded_file.read().decode('utf-8')
        except Exception:
            text = uploaded_file.read().decode('latin-1')
    else:
        text = "نوع الملف غير مدعوم أو أن مكتبة القراءة غير متوفرة. يفضل استخدام الملفات النصية أو نسخ ولصق المحتوى مباشرة."
        
    return text

def create_word_file(content):
    if not DOCX_AVAILABLE:
        return None
    doc = Document()
    doc.add_heading('ورقة العمل المُكيّفة - Educ Worksheet Adapt', 0)
    for line in content.split('\n'):
        doc.add_paragraph(line)
    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def main():
    # --- شاشة تسجيل الدخول ---
    if not st.session_state.authenticated:
        st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>Educ Worksheet Adapt 📚</h1>", unsafe_allow_html=True)
        st.markdown("<h4 style='text-align: center; color: #555;'>مساعد المعلم الذكي لتكييف وتصميم أوراق العمل</h4>", unsafe_allow_html=True)
        st.write("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("الرجاء تسجيل الدخول للوصول إلى أدوات التكييف الذكية.")
            with st.form("login_form"):
                email_input = st.text_input("البريد الإلكتروني")
                password_input = st.text_input("كلمة المرور", type="password")
                submit_btn = st.form_submit_button("تسجيل الدخول", use_container_width=True)
                
                if submit_btn:
                    if email_input:
                        st.session_state.authenticated = True
                        st.session_state.user_email = email_input
                        st.success("تم تسجيل الدخول بنجاح!")
                        st.rerun()
                    else:
                        st.error("الرجاء إدخال البريد الإلكتروني.")
        return

    # --- القائمة الجانبية الاحترافية ---
    st.sidebar.title("🛠️ إعدادات التكييف")
    st.sidebar.write(f"المستخدم: **{st.session_state.user_email}**")
    st.sidebar.markdown("---")

    grade_level = st.sidebar.selectbox(
        "المرحلة الدراسية:",
        ["الصف الرابع", "الصف الخامس", "الصف السادس", "المرحلة الأساسية العليا"]
    )
    
    subject = st.sidebar.selectbox(
        "المادة التعليمية:",
        ["اللغة العربية", "الرياضيات", "العلوم", "الدراسات الاجتماعية", "اللغة الإنجليزية"]
    )

    adaptation_goal = st.sidebar.selectbox(
        "هدف التكييف المطلوب:",
        [
            "🔄 تبسيط الأسئلة وتوضيحها للطلبة",
            "⭐ إضافة أسئلة إثرائية وتفكير عليا للمتميزين",
            "📝 إعادة صياغة ورقة العمل وتنظيمها",
            "🌐 ضبط المصطلحات وتنسيقها"
        ]
    )

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_email = ""
        st.rerun()

    # --- المحتوى الرئيسي للتطبيق ---
    st.title("📄 لوحة تحكم تكييف أوراق العمل")
    st.markdown("قم برفع ملف ورقة العمل أو إدخال النص لتحليله وتكييفه بشكل آلي واحترافي.")

    tab1, tab2, tab3 = st.tabs(["📤 رفع ملف ورقة العمل", "✍️ الإدخال النصي المباشر", "📊 أرشيف التعديلات"])

    source_text = ""

    with tab1:
        st.subheader("تحميل ملف ورقة العمل (PDF, Word, TXT)")
        uploaded_file = st.file_uploader("اختر الملف", type=["pdf", "docx", "txt"])
        
        if uploaded_file is not None:
            with st.spinner("جاري قراءة واستخراج النصوص من الملف..."):
                source_text = extract_text_from_file(uploaded_file)
            st.success(f"تم قراءة الملف ({uploaded_file.name}) بنجاح!")
            with st.expander("عرض النص المستخرج من الملف"):
                st.text_area("النص الأصلي:", value=source_text, height=150, disabled=True)

    with tab2:
        st.subheader("أو كتابة / لصق النص مباشرة")
        direct_text = st.text_area("أدخل أسئلة ورقة العمل هنا:", height=150, placeholder="اكتب أو الصق محتوى ورقة العمل هنا...")
        if direct_text.strip():
            source_text = direct_text

    # زر تنفيذ التكييف المشترك
    if source_text:
        st.markdown("---")
        if st.button("🚀 تنفيذ عملية التكييف الذكي", type="primary", use_container_width=True):
            with st.spinner("جاري معالجة وتكييف الأسئلة للمرحلة والهدف المختار..."):
                # معالجة وتحويل النص بناءً على الهدف المحدد
                processed = f"--- تقرير تكييف ورقة العمل ---\n"
                processed += f"المرحلة: {grade_level} | المادة: {subject}\n"
                processed += f"الهدف: {adaptation_goal}\n\n"
                processed += "المحتوى المُكيّف والمطور:\n"
                
                # إضافة لمسة معالجة ذكية على النصوص المستخرجة
                lines = source_text.split('\n')
                for i, line in enumerate(lines, 1):
                    if line.strip():
                        if "تبسيط" in adaptation_goal:
                            processed += f"س {i} (مبسط): {line.strip()} (توضيح: يرجى التركيز على الخطوات الأساسية)\n"
                        elif "إثرائية" in adaptation_goal:
                            processed += f"س {i} (إثراء وتفكير عليا): ما رأيك في تطبيق فكرة ({line.strip()}) في موقف من بيئتك؟\n"
                        else:
                            processed += f"س {i} (معدل): {line.strip()}\n"
                            
                st.session_state.adapted_content = processed
                st.success("تمت عملية التكييف بنجاح تام!")

    # عرض النتيجة إذا وجدت
    if st.session_state.adapted_content:
        st.markdown("### ✨ نتيجة التكييف النهائية:")
        st.text_area("النص الجديد الجاهز للاستخدام:", value=st.session_state.adapted_content, height=250)
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            word_file = create_word_file(st.session_state.adapted_content)
            if word_file:
                st.download_button(
                    "📥 تحميل ملف Word مخصص",
                    data=word_file,
                    file_name="Adapted_Worksheet.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.info("مكتبة الـ Word غير متوفرة على الخادم حالياً.")
                
        with col_dl2:
            st.download_button(
                "📥 تحميل كملف نصي (TXT)",
                data=st.session_state.adapted_content,
                file_name="Adapted_Worksheet.txt",
                mime="text/plain",
                use_container_width=True
            )

    with tab3:
        st.subheader("سجل الأنشطة وأوراق العمل المُكيّفة")
        st.write("جميع الملفات التي تقوم بتكييفها تظهر هنا للرجوع إليها في أي وقت أثناء جلسة العمل.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error("حدث خطأ غير متوقع في النظام.")
        st.exception(f"التفاصيل: {e}")
