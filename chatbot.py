import os
import streamlit as st

from main import SpreadsheetLLMWrapper
from tempfile import NamedTemporaryFile

class Arguments:
    def __init__(self, model, table, question):
        self.model = model
        self.table = table
        self.question = question

args = Arguments(None, False, None)

with st.sidebar:
    file = st.file_uploader("Upload Spreadsheet", type='xls')
    model_name = st.selectbox("Model", ('gpt-3.5', 'gpt-4', 'mistral', 'llama-2', 'llama-3', 'phi-3'))
    task = st.checkbox('Identify Number of Tables')

st.title("💬 Spreadsheet Chat")
st.caption("🚀 Chatbot for SpreadsheetLLM")
if "messages" not in st.session_state:
    st.session_state["messages"] = [{"role": "assistant", "content": "How can I help you?"}]

for msg in st.session_state.messages:
    st.chat_message(msg["role"]).write(msg["content"])

if prompt := st.chat_input():    
    args.question = prompt
    args.model = model_name
    st.session_state.messages.append({"role": "user", "content": args.question})
    st.chat_message("user").write(args.question)

    #Check for file first
    if not file:
        st.session_state.messages.append({"role": "assistant", "content": "Please upload an Excel file first"})
        st.chat_message("assistant").write("Please upload an Excel file first")
    else:
        wrapper = SpreadsheetLLMWrapper()
        #Create temporary file to pull from
        with NamedTemporaryFile(dir='.', suffix='.xls', delete=False) as f:
            f.write(file.getbuffer())
            wb = wrapper.read_spreadsheet(f.name)
            areas, compress_dict = wrapper.compress_spreadsheet(wb)   
        os.remove(f.name)  
        if task:
            args.table = True
        output = wrapper.llm(args, areas, compress_dict)
        st.session_state.messages.append({"role": "assistant", "content": output})
        st.chat_message("assistant").write(output)