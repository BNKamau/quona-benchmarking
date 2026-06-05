import streamlit as st

if not st.user.is_logged_in:
    st.title("Login")
    st.button("Sign in with Google", on_click=st.login)
    st.stop()

st.write(f"Logged in as: {st.user.email}")
st.button("Sign out", on_click=st.logout)
