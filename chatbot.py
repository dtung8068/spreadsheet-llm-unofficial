import os
import streamlit as st
import transformers

from main import Main
from openai import OpenAI

with st.sidebar:
    FILE_NAME = st.file_uploader("Upload Spreadsheet", type='xls')
    MODEL_NAME = st.selectbox("Model", ('gpt-3.5', 'gpt-4', 'mistral', 'llama-2', 'llama-3', 'phi-3'))
    TASK = st.checkbox('Identify Number of Tables')

st.title("💬 SpreadsheetGPT")
st.caption("🚀 Chatbot for SpreadsheetLLM")
if "messages" not in st.session_state:
    st.session_state["messages"] = [{"role": "assistant", "content": "How can I help you?"}]

for msg in st.session_state.messages:
    st.chat_message(msg["role"]).write(msg["content"])

if prompt := st.chat_input():    
    st.session_state.messages.append({"role": "user", "content": prompt})
    st.chat_message("user").write(prompt)

    #Check for file first
    if not FILE_NAME:
        st.session_state.messages.append({"role": "assistant", "content": "Please upload an Excel file first"})
        st.chat_message("assistant").write("Please upload an Excel file first")
    else:
        #Copy code from SpreadsheetLLM
        client = OpenAI(api_key=os.environ['OPENAI_API_KEY'])
        response = client.chat.completions.create(model="gpt-3.5-turbo", messages=st.session_state.messages)
        msg = response.choices[0].message.content
        st.session_state.messages.append({"role": "assistant", "content": msg})
        st.chat_message("assistant").write(msg)