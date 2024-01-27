#=======================================================================
## 0.1 Importing libraries, data

#Import the necessary packages
import streamlit as st
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import datetime as dt

#Setting the web app page name (optional)
st.set_page_config(page_title='Data Investigation | Streamlit App', page_icon=None, layout="wide")

#Setting markdown
st.markdown("<h1 style='text-align: center;'>Data Investigation Web App by NivAnalytics</h1>", unsafe_allow_html=True)

#Creating dynamic file upload option in sidebar
uploaded_file = st.sidebar.file_uploader("*Upload Excel file here*")

if uploaded_file is not None:
    file_path = uploaded_file
  
    #User prompt to select sheet name in uploaded file
    sh = st.sidebar.selectbox("*Which sheet name in the file should be read?*",pd.ExcelFile(file_path).sheet_names)

    #User prompt to define row with column names if they aren't in the header row in the uploaded file
    h = st.sidebar.number_input("*Which row contains the column names?*",0,100)

    #Reading the data 
    data = pd.read_excel(file_path,header=h,sheet_name=sh)
#=======================================================================
## 0.2 Pre-processing datasets

    #Replacing underscores in column names with spaces
    data.columns = data.columns.str.replace('_',' ') 

    data = data.reset_index()

    #Converting column names to title case
    data.columns = data.columns.str.title()

    #Filling nulls in object columns in file selected with blanks
    for col in data.select_dtypes('object'):
        data[col] = data[col].fillna('')

    #Horizontal divider
    st.sidebar.divider()
#=====================================================================================================
## 1. Overview of the data
    st.write( '### 1. Overview of the Data ')

    #View the dataframe in streamlit
    st.dataframe(data, use_container_width=True)

    #Horizontal divider
    st.divider()
#=====================================================================================================
## 2. Understanding the data
    st.write( '### 2. Understanding the Data ')

    #Creating radio button and sidebar simulataneously
    selected = st.sidebar.radio( "**B) What do you want to know about the data?**", 
                                ["Field Descriptions",
                                "Summary Statistics", 
                                "Value Counts of Fields",
                                "Data Shape"])

    #Showing field types
    if selected == 'Field Descriptions':
        fd = data.dtypes.reset_index().rename(columns={'index':'Field Name',0:'Field Type'}).sort_values(by='Field Type',ascending=False).reset_index(drop=True)
        st.dataframe(fd, use_container_width=True)

    #Showing summary statistics
    elif selected == 'Summary Statistics':
        st.dataframe(data.describe(include='all').round(2).fillna(''), use_container_width=True)

    #Showing value counts of object fields
    elif selected == 'Value Counts of Fields':
        # creating radio button and sidebar simulataneously if this main selection is made
        sub_selected = st.sidebar.radio( "*Which field should be investigated?*",data.select_dtypes('object').columns)
        vc = data[sub_selected].value_counts().reset_index().rename(columns={'count':'Count'}).reset_index(drop=True)
        st.dataframe(vc, use_container_width=True)

    #Showing the shape of the dataframe
    else:
        st.write('###### The shape of the data is :',data.shape)

    #Horizontal divider
    st.divider()

    #Horizontal divider in sidebar
    st.sidebar.divider()
#=====================================================================================================
## 3. Visualisation
    st.write( '### 3. Data Visualization ')

    #Selecting whether visualisation is required
    vis_select = st.sidebar.checkbox("**C) Is visualisation required for this dataset?**")

    if vis_select:

        #Creating mutiselect tab in the left sidebar
        graph_options = st.sidebar.selectbox( "*Select chart type*", options=['Bar Chart','Line Chart'])

        #Assigning color code
        clr = 'r'

        #Option for bar chart
        if 'Bar Chart' in graph_options:

            #Defining columns to display in streamlit
            st.write('#### Bar Chart')

            #Selecting dependent variable(s) or field(s) whose values should be evaluated
            x = st.sidebar.selectbox("*Bar chart: Select values to be evaluated*", options=data.columns)

            #Selecting operation to be performed
            var = st.sidebar.selectbox("*Bar chart: Select operation to be performed*", options=['count','sum','mean'])

            #Selecting fields to evaluate the above values against
            features = st.sidebar.multiselect("*Bar chart: Select fields to compare the values against*", options=data.columns, default=[])

            #Looping across all selected fields
            for feature in features:

                #Creating pivot for analysis --> visualisation
                pivot_df = pd.pivot_table(data,index=feature,values=x,aggfunc=var).sort_values(by=x,ascending=False).reset_index().head(20)

                #Creating bar plot + aesthetics
                fig1 = plt.figure()
                ax = sns.barplot(x=pivot_df[x], y=pivot_df[feature], color=clr)
                plt.xlabel(x,fontweight='bold',color=clr)
                plt.ylabel(feature,fontweight='bold',color=clr)
                plt.grid(which='major',axis='x')
                plt.title(f'{var.title()} of {x} by {feature}',fontweight='bold',color=clr)

                #Adding data value labels to each bar
                for index, value in enumerate(pivot_df[x]):
                    ax.text(value, index, f'{value:.0f}', color='white', ha="right", va="center", fontweight='bold')
                
                st.pyplot(fig1)

        #Option for line chart
        if 'Line Chart' in graph_options:

            #Defining columns to display in streamlit
            st.write('#### Line Chart')

            #Selecting dependent variable(s) or field(s) whose values should be evaluated
            evals = st.sidebar.multiselect("*Line chart: Select values to be evaluated*", options=data.columns, default=[])

            #Selecting datetime fields to evaluate the above values against
            feature = st.sidebar.selectbox("*Line chart: Select date/time field to compare the values against*",options=data.columns)

            #Selecting operation to be performed
            var = st.sidebar.selectbox("*Line chart chart: Select operation to be performed*", options=['count','sum','mean'])

            #Selecting if date/time grouping is required
            grouping = st.sidebar.checkbox("*Is grouping of the date/time field required?*")

            if grouping:
                
                #Creating year and month columns for grouping
                grouping_sel = st.sidebar.radio("*How should it be grouped?*",["Year","Month","Year-Month"])
                if grouping_sel == 'Year':
                    data[feature] = data[feature].dt.year
                    rot = 45

                elif grouping_sel == 'Month':
                    data[feature] = data[feature].dt.month
                    rot = 45
                
                elif grouping_sel == 'Year-Month':
                    data[feature] = pd.to_datetime(data[feature].dt.year.astype(str) + '-' + data[feature].dt.month.astype(str) + '-01')
                    rot = 90

            else:
                grouping_sel = feature
                rot = 45

            #Creating pivot for analysis --> visualisation
            pivot_df = pd.pivot_table(data,index=feature,values=evals,aggfunc=var).reset_index()

            #Creating line plot + aesthetics
            fig2 = plt.figure()
            plt.plot(pivot_df[feature],pivot_df[evals],color=clr)
            plt.xlabel(feature,fontweight='bold',color=clr)
            plt.xticks(pivot_df[feature], rotation=rot)
            plt.grid(which='major')
            plt.xlim(pivot_df[feature].min(), pivot_df[feature].max())
            plt.title(f'{var.title()} of {evals} by {feature} in grouping by {grouping_sel}',fontweight='bold',color=clr)
            st.pyplot(fig2)

else:
    st.info("Please upload an Excel file to proceed.")

#Horizontal divider
st.divider()

#My branding
st.write("**Created by NivAnalytics - https://www.linkedin.com/in/nivantha-bandara/**")
