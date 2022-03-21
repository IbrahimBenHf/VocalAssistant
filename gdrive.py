import streamlit as st
from streamlit_chat import message as st_message

if "history" not in st.session_state:
    st.session_state.history = []

st.title("Hello Chatbot")


#streamlit chat
def generate_answer():
    user_message = st.session_state.input_text
    message_bot = user_message+' bot'
    st.session_state.history.append({"message": user_message, "is_user": True})
    st.session_state.history.append({"message": message_bot, "is_user": False})
    st.session_state["input_text"] = ""
state = False
def change_state(state):
    if state:
        st.write('Why hello there')
    else:
        st.write('Goodbye')


st.text_input("Talk to the bot", key="input_text", on_change=generate_answer)
st.button("micro",on_click=change_state(state))
st.write("Micro On", key='micro')
st.image()




msg_limit = 10000
for chat in st.session_state.history:
    st_message(chat['message'], chat['is_user'],None,None,str(msg_limit))  # unpacking
    msg_limit=msg_limit-1


st.sidebar.markdown("# About the project")
st.sidebar.markdown("An end to end project to recognize facial emotions of each customer. The aim of this project is to analyze data collected from live stream cameras capturing the agency's cutomers.")
st.sidebar.markdown("This app requires three IP cameras connected to the same network, and some url configurations.")
st.sidebar.markdown("# Our approach")
st.sidebar.markdown("The whole application is devided into three main steps : ")
st.sidebar.markdown("1 - Capture the entering customer and his emotion")
st.sidebar.markdown("2 - Save all cutomer's emotions during the processing operation in the agency's window")
st.sidebar.markdown("3 - Capture the exiting customer, face re-identification, exiting emotion and elapsed time")
st.sidebar.markdown("# Choose an App to run")
option = st.sidebar.selectbox("", ('--  No selected app  --','Enter App','Window App','Exit App'))
