import streamlit as st
#import docx
import docxtpl
import os.path
import pathlib
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

zipObj = ZipFile("sample_new_options.zip", "w")

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

ZipfileDotZip = "sample_new_options"

with open(ZipfileDotZip, "rb") as f:
    bytes = f.read()
    b64 = base64.b64encode(bytes).decode()
    href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}.zip'>\
        Click last model weights\
    </a>"
st.sidebar.markdown(href, unsafe_allow_html=True)

"""if uploaded_file is not None:
    bytes_data = uploaded_file.getvalue()
    data = uploaded_file.getvalue().decode('utf-8').splitlines()         
    st.session_state["preview"] = ''
    for i in range(0, min(5, len(data))):
        st.session_state["preview"] += data[i]
preview = st.text_area("CSV Preview", "", height=150, key="preview")
upload_state = st.text_area("Upload State", "", key="upload_state")
def upload():
    if uploaded_file is None:
        st.session_state["upload_state"] = "Upload a file first!"
    else:
        data = uploaded_file.getvalue().decode('utf-8')
        parent_path = pathlib.Path(__file__).parent.parent.resolve()           
        save_path = os.path.join(parent_path, "data")
        complete_name = os.path.join(save_path, uploaded_file.name)
        destination_file = open(complete_name, "w")
        destination_file.write(data)
        destination_file.close()
        st.session_state["upload_state"] = "Saved " + complete_name + " successfully!"
st.button("Upload file to Sandbox", on_click=upload)
"""
