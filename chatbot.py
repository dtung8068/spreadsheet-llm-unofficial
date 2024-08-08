import os
import streamlit as st
import transformers

from openai import OpenAI

with st.sidebar:
    FILE_NAME = st.file_uploader("Upload Spreadsheet")
    MODEL_NAME = st.selectbox("Model", ('gpt-3.5', 'gpt-4', 'mistral', 'llama-2', 'llama-3', 'phi-3'))
    TASK = st.selectbox("Task", ('Identify Table', 'Question/Answer'))

st.title("💬 SpreadsheetGPT")
st.caption("🚀 Chatbot for SpreadsheetLLM")
if "messages" not in st.session_state:
    st.session_state["messages"] = [{"role": "assistant", "content": "How can I help you?"}]

for msg in st.session_state.messages:
    st.chat_message(msg["role"]).write(msg["content"])

if prompt := st.chat_input():
    

    client = OpenAI(api_key=os.environ['OPENAI_API_KEY'])
    st.session_state.messages.append({"role": "user", "content": prompt})
    st.chat_message("user").write(prompt)
    response = client.chat.completions.create(model="gpt-3.5-turbo", messages=st.session_state.messages)
    msg = response.choices[0].message.content
    st.session_state.messages.append({"role": "assistant", "content": msg})
    st.chat_message("assistant").write(msg)