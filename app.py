import streamlit as st

st.title("Educ Worksheet Adapt - Test Mode")
st.success("السيرفر يعمل بشكل ممتاز بدون أي أخطاء!")

user_name = st.text_input("أدخل اسمك:")
if user_name:
    st.write(f"مرحباً بك يا أستاذ {user_name}! النظام جاهز تماماً.")
