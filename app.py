import streamlit as st

# إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="Educ Worksheet Adapt",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تهيئة متغيرات الجلسة بأمان مطلق لمنع انهيار الخادم
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_email" not in st.session_state:
    st.session_state.user_email = ""
if "adapted_content" not in st.session_state:
    st.session_state.adapted_content = ""

def main():
    # --- 1. شاشة تسجيل الدخول ---
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

    # --- 2. القائمة الجانبية الاحترافية ---
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

    # --- 3. واجهة لوحة التحكم الرئيسية ---
    st.title("📄 لوحة تحكم تكييف أوراق العمل")
    st.markdown("قم بلصق محتوى ورقة العمل أو الأسئلة هنا، وسيقوم النظام بمعالجتها وتكييفها فوراً.")

    raw_text = st.text_area(
        "نص ورقة العمل / الأسئلة:", 
        height=220, 
        placeholder="الصق الأسئلة هنا (كل سؤال في سطر)..."
    )
    
    if st.button("🚀 تنفيذ التكييف الذكي", type="primary", use_container_width=True):
        if raw_text.strip():
            with st.spinner("جاري المعالجة..."):
                processed = f"=== تقرير تكييف ورقة العمل ===\n"
                processed += f"المرحلة: {grade_level} | المادة: {subject}\n"
                processed += f"الهدف: {adaptation_goal}\n"
                processed += f"{'='*30}\n\n"
                
                lines = raw_text.split('\n')
                counter = 1
                for line in lines:
                    cleaned = line.strip()
                    if cleaned:
                        if "تبسيط" in adaptation_goal:
                            processed += f"س{counter} (مبسط ومباشر): {cleaned}\n   (توجيه: اطلب من الطالب التركيز على المعطيات الأساسية)\n\n"
                        elif "إثرائية" in adaptation_goal:
                            processed += f"س{counter} (إثراء وتفكير عليا): {cleaned}\n   (توجيه: ناقش كيف يمكن تطبيق هذه الفكرة في الحياة العملية)\n\n"
                        else:
                            processed += f"س{counter} (معدل ومُنظم): {cleaned}\n\n"
                        counter += 1
                            
                st.session_state.adapted_content = processed
                st.success("تمت عملية التكييف بنجاح!")
        else:
            st.warning("الرجاء إدخال نص الأسئلة أولاً.")

    # --- 4. عرض النتائج وخيارات التحميل الآمنة ---
    if st.session_state.adapted_content:
        st.markdown("### ✨ المحتوى المُكيف النهائي:")
        st.text_area("النتيجة الجاهزة:", value=st.session_state.adapted_content, height=250)
        
        st.download_button(
            "📥 تحميل النتيجة كملف نصي (TXT)",
            data=st.session_state.adapted_content,
            file_name="Adapted_Worksheet.txt",
            mime="text/plain",
            use_container_width=True
        )

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error("حدث خطأ في النظام الداخلي.")
        st.exception(f"التفاصيل: {e}")
