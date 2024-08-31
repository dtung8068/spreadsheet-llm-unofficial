![image](https://github.com/user-attachments/assets/72cabb2a-10b6-4c43-a265-bb92ec474d54)

# SpreadsheetLLM
 My unofficial implementation of Microsoft's SpreadsheetLLM paper, found here: https://arxiv.org/pdf/2407.09025.

# Requirements
You will need Python 3.8+, Pandas 2.2+, NumPy, xlrd, transformers, huggingface_hub, and openai. 

You will also need to download the VFUSE dataset from TableSense, found here: https://figshare.com/projects/Versioned_Spreadsheet_Corpora/20116  

Environment Variables: OPENAI_API_KEY for GPT 3.5/4, HUGGING_FACE_KEY for Llama-2/3, Phi-3, and Mistral

# Directions
By default, running python main.py will generate the number of tables in 7b5a0a10-e241-4c0d-a896-11c7c9bf2040.xls. Use the command line arguments if you want to compress all files in a given directory, change the model, etc.

To run the chatbot, run `streamlit run chatbot.py`. 

# Limitations
1. Only text was considered for the structural anchor-based extraction, formatting (border, color, etc.) was not considered
2. NFS Identification currently relies on regular expressions and may not be robust
3. Only .xls files will work at this time

# Future Plans
1. Running tests on the LLM
2. Enabling compatibility with other Excel formats (.xlsx, .xlsm, etc.)
