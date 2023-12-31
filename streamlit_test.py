import streamlit as st
#import docx
import docxtpl
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
from zipfile import ZipFile
from io import BytesIO
import base64

st.write("""
# File Picker
""")

option = st.selectbox('Select letter template',('template1','template2'))

uploaded_file_data = st.file_uploader("Choose a file with the required data")

if uploaded_file_data is not None:
    df = pd.read_excel(uploaded_file_data)
    print(df)

uploaded_file_template = st.file_uploader("Choose the template file")

if uploaded_file_template is not None:
    template = uploaded_file_template

zipObj = ZipFile("sample.zip", "w")

if option =='template1':
    for index, row in df.iterrows():
        doc = DocxTemplate(template) 
        context = {'Name': row['Name'],
                'Age': row['Age'],
                'Designation': row['Designation'],
                }

        print(context)
        print(doc.render(context))
        doc.save(f"generated_doc2_{index}.docx")
        zipObj.write(f"generated_doc2_{index}.docx")

    zipObj.close()

if option =='template2':
    for index, row in df.iterrows():
        doc = DocxTemplate(template) 
        context = {'Name': row['Name'],
                'Age': row['Age'],
                'City': row['City'],
                }

        print(context)
        print(doc.render(context))
        doc.save(f"generated_doc2_{index}.docx")
        zipObj.write(f"generated_doc2_{index}.docx")
    zipObj.close()

ZipfileDotZip = "sample.zip"

with open(ZipfileDotZip, "rb") as f:
    bytes = f.read()
    b64 = base64.b64encode(bytes).decode()
    href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
        Download letters\
    </a>"
    
st.sidebar.markdown(href, unsafe_allow_html=True)

