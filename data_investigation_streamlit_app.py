#=======================================================================
## 0.1 Importing libraries, data

#Import the necessary packages
import streamlit as st
import openpyxl
import pygwalker as pyg
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import datetime as dt
import sys
import os
from PIL import Image

#Getting path of images
cwd_dir = os.path.dirname(__file__)
rel_path = './images'
images_path = os.path.join(cwd_dir,rel_path)
logo = Image.open(images_path + '/Browser Icon Reverse.png')

#Configuring Streamlit page
st.set_page_config(page_title = 'Exploratory Data Analysis App',page_icon = logo,layout = "wide")

#Creating container with columns for page heading, specific title and logo
with st.container():
    col1, col2, padding = st.columns([10,1,1])

    #Setting page header and subheader
    with col1:
        st.markdown("### Exploratory Data Analysis App")
        
    #Setting logo
    with col2:
        niv_logo = Image.open(images_path + '/Social Logo.png')
        st.image(niv_logo, width=200, output_format='auto')

st.sidebar.write("****A) File upload****")

#User prompt to select file type
ft = st.sidebar.selectbox("*What is the file type?*",["Excel", "csv"])

#Creating dynamic file upload option in sidebar
uploaded_file = st.sidebar.file_uploader("*Upload file here*")

if uploaded_file is not None:
    file_path = uploaded_file

    if ft == 'Excel':
        try:
            #User prompt to select sheet name in uploaded Excel
            sh = st.sidebar.selectbox("*Which sheet name in the file should be read?*",pd.ExcelFile(file_path).sheet_names)
            #User prompt to define row with column names if they aren't in the header row in the uploaded Excel
            h = st.sidebar.number_input("*Which row contains the column names?*",0,100)
        except:
            st.info("File is not recognised as an Excel file")
            sys.exit()
    
    elif ft == 'csv':
        try:
            #No need for sh and h for csv, set them to None
            sh = None
            h = None
        except:
            st.info("File is not recognised as a csv file.")
            sys.exit()

    #Caching function to load data
    @st.cache_data(experimental_allow_widgets=True)
    def load_data(file_path,ft,sh,h):
        
        if ft == 'Excel':
            try:
                #Reading the excel file
                data = pd.read_excel(file_path,header=h,sheet_name=sh,engine='openpyxl')
            except:
                st.info("File is not recognised as an Excel file.")
                sys.exit()
    
        elif ft == 'csv':
            try:
                #Reading the csv file
                data = pd.read_csv(file_path)
            except:
                st.info("File is not recognised as a csv file.")
                sys.exit()
        
        return data

    data = load_data(file_path,ft,sh,h)
#=======================================================================
## 0.2 Pre-processing datasets

    #Replacing underscores in column names with spaces
    data.columns = data.columns.str.replace('_',' ') 

    data = data.reset_index()

    #Converting column names to title case
    data.columns = data.columns.str.title()

    #Horizontal divider
    st.sidebar.divider()
#=====================================================================================================
## 1. Overview of the data
    st.write( '### 1. Dataset Preview ')

    try:
      #View the dataframe in streamlit
      st.dataframe(data, use_container_width=True,hide_index=True)

    except:
      st.info("The file wasn't read properly. Please ensure that the input parameters are correctly defined.")
      sys.exit()

    #Horizontal divider
    st.divider()
#=====================================================================================================
## 2. Understanding the data
    st.write( '### 2. High-Level Overview ')

    #Creating radio button and sidebar simulataneously
    selected = st.sidebar.radio( "**B) What would you like to know about the data?**", 
                                ["Data Dimensions",
                                 "Field Descriptions",
                                "Summary Statistics", 
                                "Value Counts of Fields"])

    #Showing field types
    if selected == 'Field Descriptions':
        fd = data.dtypes.reset_index().rename(columns={'index':'Field Name',0:'Field Type'}).sort_values(by='Field Type',ascending=False).reset_index(drop=True)
        st.dataframe(fd, use_container_width=True,hide_index=True)

    #Showing summary statistics
    elif selected == 'Summary Statistics':
        ss = pd.DataFrame(data.describe(include='all').round(2).fillna(''))
        #Adding null counts to summary statistics
        nc = pd.DataFrame(data.isnull().sum()).rename(columns={0: 'count_null'}).T
        ss = pd.concat([nc,ss]).copy()
        st.dataframe(ss, use_container_width=True)

    #Showing value counts of object fields
    elif selected == 'Value Counts of Fields':
        # creating radio button and sidebar simulataneously if this main selection is made
        sub_selected = st.sidebar.radio( "*Which field should be investigated?*",data.select_dtypes('object').columns)
        vc = data[sub_selected].value_counts().reset_index().rename(columns={'count':'Count'}).reset_index(drop=True)
        st.dataframe(vc, use_container_width=True,hide_index=True)

    #Showing the shape of the dataframe
    else:
        st.write('###### The data has the dimensions :',data.shape)

    #Horizontal divider
    st.divider()

    #Horizontal divider in sidebar
    st.sidebar.divider()
#=====================================================================================================
## 3. Visualisation

    #Selecting whether visualisation is required
    vis_select = st.sidebar.checkbox("**C) Is visualisation required for this dataset (hide sidebar for full view of dashboard) ?**")

    if vis_select:

        st.write( '### 3. Visual Insights ')

        try:
            #Creating a PyGWalker Dashboard
            walker = pyg.walk(data,return_html=True)

            #Render the HTML string in Streamlit
            st.components.v1.html(walker, width=1250, height=800)
            
        except Exception as e:
            st.error(f"Error occured: {e}")
            
else:
    st.info("Please upload a file to proceed.")

#Horizontal divider
st.divider()

#My branding
st.write("**Created by NivAnalytics - https://www.nivanalytics.com**")
