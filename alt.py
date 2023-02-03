import pandas as pd
#import openpyxl as px
import streamlit as st
#import plotly_express as px
import altair as alt
import plotly.graph_objects as go

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

rm_company=st.sidebar.selectbox(
        'Please Select Company',
        options= rm_selection['COMPANY'].unique()
 )

rm_company_selection = df.query(
    '`RM NAME`==@RM & `COMPANY` == @rm_company'
)

month =st.sidebar.selectbox(
        'Please Select Month',
        options= rm_selection['MONTH'].unique()
 )

rm_month = df.query(
    '`RM NAME`==@RM & `MONTH` == @month'
)


# '''st.bar_chart(rm_month,x='COMPANY', y='FACTOR 1
# )
# st.line_chart(rm_month,
#               y='COMPANY', x='40M CAPITAL NEEDED ')
# '''
#st.dataframe(rm_selection)
#st.dataframe(rm_company_selection)
#st.dataframe(rm_month)

#---------GAUGES-------------
st.title(rm_company + ' Capital Needed')
kpi1, kpi2, kpi3 =st.columns(3)
a=rm_month.set_index('COMPANY')

with kpi1:
    fig = go.Figure(go.Indicator(
        domain={'x': [0, 1], 'y': [0, 1]},
        value=a.loc[rm_company,'25M CAPITAL NEEDED'],
        mode='gauge+number+delta',
        title={'text': '25M CAPITAL NEEDED'},
        delta={'reference': 300},
        gauge={'axis': {'range': [None, a.loc[rm_company,'25M CAPITAL NEEDED']*5]},
               'steps': [
                   {'range': [0, 250], 'color': 'lightgray'},
                   {'range': [250, 400], 'color': 'lightgray'}, ],
               'threshold': {'line': {'color': 'red', 'width': 4}, 'thickness': 1, 'value': 490}

               }
    )
    )
    st.plotly_chart(fig, use_container_width=True)
#st.dataframe(a)
#st.dataframe(rm_month)
with kpi2:
    fig2 = go.Figure(go.Indicator(
        domain={'x': [0, 1], 'y': [0, 1]},
        value=a.loc[rm_company,'30M CAPITAL NEEDED'],
        mode='gauge+number+delta',
        title={'text': '30M Capital Needed'},
        delta={'reference': 300},
        gauge={'axis': {'range': [None, a.loc[rm_company,'25M CAPITAL NEEDED']*5]},
               'steps': [
                   {'range': [0, 250], 'color': 'lightgray'},
                   {'range': [250, 400], 'color': 'lightgray'}, ],
               'threshold': {'line': {'color': 'red', 'width': 4}, 'thickness': 1, 'value': 490}

               }
    )
    )
    st.plotly_chart(fig2, use_container_width=True)
with kpi3:
    fig3 = go.Figure(go.Indicator(
        domain={'x': [0, 1], 'y': [0, 1]},
        value=a.loc[rm_company,'40M CAPITAL NEEDED '],
        mode='gauge+number+delta',
        title={'text': '40M Capital Needed'},
        delta={'reference': 300},
        gauge={'axis': {'range': [None, a.loc[rm_company,'40M CAPITAL NEEDED ']*5]},
               'steps': [
                   {'range': [0, a.loc[rm_company,'40M CAPITAL NEEDED ']* 2], 'color': 'lightgray'},
                   {'range': [ a.loc[rm_company,'40M CAPITAL NEEDED ']*2,  a.loc[rm_company,'40M CAPITAL NEEDED ']*4], 'color': 'lightgray'}, ],
               'threshold': {'line': {'color': 'red', 'width': 4}, 'thickness': 1, 'value': 9000000}

               }
    )
    )
    st.plotly_chart(fig3, use_container_width=True)

#------ BAR CHART-------------
grap= alt.Chart(rm_selection).mark_bar().encode(
    y='MONTH', x='40M CAPITAL NEEDED ',
    color = 'MONTH',
    row = 'COMPANY'
)
st.title(' Company Performance ')
st.altair_chart(grap, use_container_width=True)

#------- CAR GAUGE------------
fig = go.Figure(go.Indicator(
    domain= {'x': [0,1], 'y':[0,1]},
    value= 450,
    mode='gauge+number+delta',
    title= {'text':'CAR'},
    delta={'reference': 300},
    gauge={'axis':{'range':[None,500]},
           'steps':[
               {'range': [0,250], 'color' : 'lightgray'},
               {'range': [250,400], 'color' : 'lightgray'},],
           'threshold':{'line':{'color':'red', 'width':4}, 'thickness':1, 'value': 490}

    }
)
)
st.title(rm_company +' CAR')
st.plotly_chart(fig)


