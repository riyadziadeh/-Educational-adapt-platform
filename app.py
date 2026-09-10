import streamlit as st

# إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="Educ Worksheet Adapt",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="auto"
)

# تهيئة متغيرات الجلسة بأمان لتجنب الأخطاء المفاجئة
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_email" not in st.session_state:
    st.session_state.user_email = ""

def main():
    # شريط التنقل العلوي البسيط
    st.sidebar.title("قائمة التطبيق")
    
    if not st.session_state.authenticated:
        st.markdown("## مرحباً بك في Educ Worksheet Adapt")
        st.info("الرجاء تسجيل الدخول أو إنشاء حساب للمتابعة والاستفادة من ميزات التعديل.")
        
        # نموذج تسجيل دخول تجريبي / أو ربط مع قاعدة البيانات الخاصة بك
        with st.form("login_form"):
            email_input = st.text_input("البريد الإلكتروني")
            password_input = st.text_input("كلمة المرور", type="password")
            submit_btn = st.form_submit_button("تسجيل الدخول")
            
            if submit_btn:
                if email_input:  # ضع شرط التحقق الخاص بك هنا
                    st.session_state.authenticated = True
                    st.session_state.user_email = email_input
                    st.success("تم تسجيل الدخول بنجاح!")
                    st.rerun()
                else:
                    st.error("الرجاء إدخال البريد الإلكتروني بشكل صحيح.")
                    
        # تنبيه للتحقق من أخطاء الوصول
        st.markdown("---")
        st.caption("إذا واجهتك مشكلة في الوصول، تأكد من فتح التطبيق بالحساب المُسجل في قائمة الاختبار الداخلي لمتجر جوجل.")
        
    else:
        st.success(f"مرحباً بك مجدداً ({st.session_state.user_email})")
        st.write("هنا تظهر لوحة التحكم الخاصة بتكييف أوراق العمل وتعديلها.")
        
        if st.button("تسجيل الخروج"):
            st.session_state.authenticated = False
            st.session_state.user_email = ""
            st.rerun()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # التقاط أي خطأ غير معالج ومنع ظهور خطأ 500 الخام للمستخدم
        st.error("حدث خطأ غير متوقع في النظام. يجتمع الفريق التقني لحله.")
        st.exception(f"التفاصيل الفنية: {e}")
