import os
import transformers

from huggingface_hub import InferenceClient
from openai import OpenAI

#If you only want to check how many tables
PROMPT_TABLE = """INSTRUCTION:
Given an input that is a string denoting data of cells in a table.
The input table contains many tuples, describing the cells with content in the spreadsheet.
Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3).
The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss',etc., represents the CELL DATA FORMATS of Excel.
The content in some cells such as 'IntNum'/'DateData'/'EmailData',etc., represents a category of data with the same format and similar semantics.
For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data.
'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B.
Some cells with empty content in the spreadsheet are not entered.
How many tables are there in the spreadsheet?
DON'T ADD OTHER WORDS OR EXPLANATION.
INPUT: """

#Part 1 of CoS
STAGE_1_PROMPT = """INSTRUCTION:
Below is a question about one certain table in this spreadsheet.
I need you to determine in which table the answer to the following question can be found, and return the RANGE of the ONE table you choose, LIKE ['range':'A1:F9'].
DON'T ADD OTHER WORDS OR EXPLANATION.
INPUT: """

#Part 2 of CoS (after filtering out table)
STAGE_2_PROMPT = """INSTRUCTION:
Given an input that is a string denoting data of cells in a table and a question about this table.
The answer to the question can be found in the table.
The input table includes many pairs, and each pair consists of a cell address and the text in that cell with a ',' in between, like 'A1,Year'.
Cells are separated by '|' like 'A1,Year|A2,Profit'.
The text can be empty so the cell data is like 'A1, |A2,Profit'.
The cells are organized in row-major order.
The answer to the input question is contained in the input table and can be represented by cell address.
I need you to find the cell address of the answer in the given table based on the given question description, and return the cell ADDRESS of the answer like '[B3]' or '[SUM(A2:A10)]'.
DON'T ADD ANY OTHER WORDS."
INPUT: """

MODEL_DICT = {'mistral': 'mistralai/Mistral-7B-Instruct-v0.2',
              'llama-2': 'meta-llama/Llama-2-7b-chat-hf',
              'llama-3': 'meta-llama/Meta-Llama-3-8B-Instruct',
              'phi-3': 'microsoft/Phi-3-mini-128k-instruct',
              'gpt-3.5': 'gpt-3.5-turbo',
              'gpt-4': 'gpt-4'}

class SpreadsheetLLM():
    def __init__(self, model):
        self.model = model
    
    def call(self, prompt):
        if self.model == 'gpt-3.5' or self.model == 'gpt-4': #OpenAI API
          completion = OpenAI(api_key=os.environ['OPENAI_API_KEY']).chat.completions.create(
            model=MODEL_DICT[self.model],
            messages=[
              {"role": "user", "content": prompt}
            ]
          )
          return completion.choices[0].message.content
        else: #Transformers InferenceClient
            output = ''
            client = InferenceClient(
              MODEL_DICT[self.model],
              token=os.environ['HUGGING_FACE_KEY'],
            )
            for message in client.chat_completion(
	            messages=[{"role": "user", "content": prompt}],
	            max_tokens=500,
	            stream=True,
            ):
                output += message.choices[0].delta.content
            return output
    
    def identify_table(self, table):
        global PROMPT_TABLE
        return self.call(PROMPT_TABLE + str(table))

    def question_answer(self, table, question):
        global STAGE_1_PROMPT
        global STAGE_2_PROMPT

        table_range = self.call(STAGE_1_PROMPT + str(table) + '\n QUESTION:' + question)
        return self.call(STAGE_2_PROMPT + str(question + table_range + '\n QUESTION:' + question))