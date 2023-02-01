import pandas as pd
import openpyxl as px
import streamlit as st
import plotly_express as px

st.set_page_config(page_title='Summary Dashboard',
                   page_icon=':bar_chart',
                   layout='wide')
hide_default_format = """
       <style>
       #MainMenu {visibility: hidden; }
       footer {visibility: hidden;}
       </style>
       """
st.markdown(hide_default_format, unsafe_allow_html=True)


df = pd.read_excel(io="C:\Kojo\effective.xlsx",
                 sheet_name='Data Model',
                         skiprows=5,
                engine= "openpyxl")

#st.dataframe(df)

st.sidebar.header('Filter the data here')
RM = st.sidebar.selectbox(
    "Select the RM",
    options= df['RM NAME'].unique(),
)

rm_selection = df.query(
    '`RM NAME` == @RM'
)

'''rm_company=st.sidebar.selectbox(
        'Please Select Company',
        options= rm_selection['COMPANY'].unique()
 )

rm_company_selection = df.query(
    '`RM NAME`==@RM & `COMPANY` == @rm_company'
)
'''
month =st.sidebar.selectbox(
        'Please Select Month',
        options= rm_selection['MONTH'].unique()
 )

rm_month = df.query(
    '`RM NAME`==@RM & `MONTH` == @month'
)

st.dataframe(rm_month)
st.bar_chart(rm_month,x='COMPANY', y='FACTOR 1'
)
st.line_chart(rm_month,
              y='COMPANY', x='40M CAPITAL NEEDED ')
