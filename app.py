import pandas as pd
from datetime import date
import streamlit as st
from pathlib import Path
from docxtpl import DocxTemplate
from docx import Document
from PIL import Image
import docx2pdf
from PyPDF2 import PdfReader  
import lxml
from run import make_inspection_document
from streamlit_gsheets import GSheetsConnection


url = 'https://docs.google.com/spreadsheets/d/1sgGFI1Iw_-PtX2ZlwEq-aEaGoOw-tqYxzXQ7666Dz0M/edit#gid=0'
conn = st.connection("gsheets", type=GSheetsConnection)

data = conn.read(spreadsheet=url)
data = data.dropna(how='all',axis=1) 
data = data.dropna(how='all',axis=0) 

def set_updatefields_true(docx_path):
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    doc = Document(docx_path)
    # add child to doc.settings element
    element_updatefields = lxml.etree.SubElement(
        doc.settings.element, f"{namespace}updateFields")
    element_updatefields.set(f"{namespace}val", "true")
    doc.save(docx_path)

st.title(":rainbow[Inspection Report Generator]")

# Get input from the user
st.header(':blue[Info]', divider='violet')
#i
# client
st.subheader(':grey[Client]')#, divider='grey')
c1,c2 = st.columns([4,1])   
data['Client'] = data['Client'].str.upper()
clients = data['Client'].unique().tolist()
clients.sort()
clients = tuple(clients)
client_name = c1.selectbox('Client name:',clients)
clientcode = data[data['Client']==client_name]['Code'].reset_index()
clientcode = clientcode['Code'][0]
client_code = c2.text_input('Client code:',clientcode)
c3,c4 = st.columns([4,1])
clientlocation = data[data['Client']==client_name]['Location'].unique().tolist()
clientlocation.sort()
clientlocation = tuple(clientlocation)
client_location = c3.selectbox('Client location:',clientlocation)
unitnumber = data[(data['Client']==client_name)&(data['Location']==client_location)]['Unit'].tolist()
# unitnumber.sort()
unitnumber = tuple(unitnumber)
unit_number = c4.selectbox('Unit number:',unitnumber)
c11, c12 = st.columns([1,1])
fpage_image = c11.file_uploader("Upload First Page Image:", type=['png','jpeg','jpg'], accept_multiple_files=False)
client_logo = c12.file_uploader("Upload Client logo:", type=['png','jpeg','jpg'], accept_multiple_files=False)
if fpage_image is not None:
    c11.image(fpage_image,width=175)


my_file = Path(rf'Files\Client\Logo\{client_name}.png')
if client_logo is not None:
    c12.image(client_logo,width=175)
elif my_file.is_file():
    client_logo = Image.open(my_file)
    c12.image(client_logo,width=175)


# inspection
st.subheader(':grey[Inspection]')
c5,c6 = st.columns([4,1])
equipment_name = c5.text_input('Equipment name:','Shell Plate Tower')
equipment_name = equipment_name.upper()
tag_number = c6.text_input('Tag number:','Tower No. 402')
tag_number = tag_number.upper()
# st.subheader('Equipment')#, divider='grey')
c7,c8 = st.columns([4,1])

inspection_date = c8.date_input('Inspection date:')
inspection_type = c7.text_input('Inspection type:','TOWER INSPECTION BY ROBOTIC CRAWLER')
# Prepared by
st.subheader(':grey[Authors]')
df = pd.DataFrame(
    [
       {"Date": date.today(),'Job':"Prepared by", "Designation": 'NDT Technician', "Name": 'Sakthivel'},
       {"Date": date.today(),'Job': "Reviewed by", "Designation": 'NDT Technician', "Name": 'Kasirajan'},
       {"Date": date.today(),'Job':'Approved By', "Designation": 'Managing Director', "Name": "Dharmaraj"},
   ]
)
edited_df = st.data_editor(df,hide_index=True,use_container_width=True)

# Content
st.header(':blue[Summary]', divider='violet')
st.subheader(':grey[Result and conclusion]')
text_list=[]
st.caption("Add '-' before each point")
result_and_conclusion = st.text_area(f'Result and conclusion:','- add point')
st.subheader(':grey[Site observation]')
st.caption("Add '#' before headings,  '$' before subheadings,  '-' before each point")
site_observation = st.text_area(f'Site observation:',"""# Heading
$ Sub-heading
- add point""")

# Upload files
st.header(':blue[Upload files]', divider='violet')
st.subheader(':grey[Overall Inspection Summary]')
overall_summary = st.file_uploader("Upload Overall Inspection Summary file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Towershell nominal thickness and height details]')
thickness_details = st.file_uploader("Upload Towershell nominal thickness and height details file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Scanning location and orientation details]')
scanning_details = st.file_uploader("Upload Scanning location and orientation details file:", type=['csv','xlsx'], accept_multiple_files=False)
st.subheader(':grey[Shellwise inspection summary]')
shellwise_inspection = st.file_uploader("Upload Shellwise inspection summary files:", type=['csv','xlsx'], accept_multiple_files=True)
st.subheader(':grey[Tower drawings and scanning location]')
tower_drawing = st.file_uploader("Upload Tower drawings and scanning location pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
if tower_drawing is not None:
    st.image(tower_drawing,width=175)

st.subheader(':grey[Shell plate pictures]')#,divider='red')
shell_plate_pics = st.file_uploader("Upload Shell plate pictures:", type=['png','jpeg','jpg'], accept_multiple_files=True)
if shell_plate_pics is not None:
    st.image(shell_plate_pics,width=175)

# st.subheader(':grey[Detailed reports]')
# result = st.number_input('Number of Sections',min_value=1)
# detailed_report = {}
# for i in range(result):
#     c9,c10 = st.columns([3,1])  
#     section_heading = c9.text_input(f'Section Heading {i+1}:')
#     section_title = c10.text_input(f'Section Title {i+1}:', section_heading[:8])
#     section = st.file_uploader(f"Upload detailed report files {i+1}:", type=['xlsx'], accept_multiple_files=False)
#     detailed_report[section_heading] = [section_title,section]
# st.divider()

filename = f"document.docx"
result =  st.button('Generate Report', use_container_width = True, type = 'primary')

if result == True:
    doc = make_inspection_document(client_name, client_location, unit_number, client_code, fpage_image, inspection_date, equipment_name, tag_number, inspection_type, edited_df, result_and_conclusion, site_observation, overall_summary, thickness_details, scanning_details, shellwise_inspection, tower_drawing, shell_plate_pics)
    doc.save(rf'Files\Temp\{filename}')
    set_updatefields_true(rf'Files\Temp\{filename}')
    st.download_button("Download report" , data=open(rf'Files\Temp\{filename}', "rb").read(), file_name=filename, use_container_width = True, mime = "application/octet-stream" )

