# SpreadsheetLLM
 My unofficial implementation of Microsoft's SpreadsheetLLM paper, found here: https://arxiv.org/pdf/2407.09025. 
 This project has only implemented the SheetCompressor portion so far.

# Requirements
You will need Python 3.8+, Pandas 2.2+, NumPy, and xlrd. 

You will also need to download the VFUSE dataset from TableSense, found here: https://figshare.com/projects/Versioned_Spreadsheet_Corpora/20116  

Environment Variables: OPENAI_API_KEY for GPT 3.5/4, HUGGING_FACE_KEY for Llama-2/3, Phi-3, and Mistral

# Directions
By default, running python main.py will compress all files in VFUSE, generate the number of tables in 7b5a0a10-e241-4c0d-a896-11c7c9bf2040.xls, and create several questions about the file.

# Limitations
1. Only text was considered for the structural anchor-based extraction, formatting (border, color, etc.) was not considered
2. NFS Identification currently relies on regular expressions and may not be robust

# Future Plans
1. Implementing the LLM Portion of the paper
2. Testing the LLM
3. Creating a Streamlit app
