import os
import sys

import constants
# print('before text loader \n')
from langchain.document_loaders import TextLoader
# print('after text loader \n')
from langchain.indexes import VectorstoreIndexCreator
from langchain.llms import OpenAI
from langchain.chat_models import ChatOpenAI
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import streamlit as st 

os.environ["OPENAI_API_KEY"] = constants.APIKEY

# Check the current file path 
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle/exe
    current_path = sys._MEIPASS
else:
    current_path = os.path.dirname(os.path.abspath(__file__))

datafile ='SignalMLA.txt'
datafile_path = os.path.join(current_path, datafile)



st.title('Talk to the Signal MLA  ')
query_input = st.text_input('Ask a question about the Signal MLA here')

# Prompt templates
# query_template = PromptTemplate(
#     input_variables =['topic'],
#     template = '{topic} and also please provide the references to the clause numbers.'
# )

# Prompt templates
# def generate_query_from_template(topic):
#     template = {topic} 
#     return template.format(topic=topic)

# # Generate the desired query string
# query = generate_query_from_template(query_input)

query=query_input

# query = sys.argv[1]
print(query)

loader = TextLoader(datafile)
index = VectorstoreIndexCreator().from_loaders([loader])

# print(index.query(query))

if query:   
    response = index.query(query)
    st.write(response)
    
