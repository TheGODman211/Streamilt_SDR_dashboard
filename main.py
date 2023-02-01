import pandas as pd
import openpyxl as px
import streamlit as st

#line 79 needs reevaluation
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

#----- READ EXCEL FILE ----

#function to parse excel workbooks into dataframes
def excel_file(loc):
    return pd.read_excel(io=loc,
                 sheet_name=None,
                         skiprows=5,
                engine= "openpyxl")

#specify the file directories
local=["C:\AllianzLifeInsuranceCompanyGh.Limited__2020_Fourth Quarter.xls.xlsx"]

# add each workbbok to a list
length=len(local)
ls=[]
for i in range(length):
    x=excel_file(local[i])
    ls.append(x)

#-----SIDEBAR-------

#get company names from their file directories
companies=[]
for i in local:
    companies.append(i[3:-5])

#sidebar header
st.sidebar.header('Please Select Company here')

#create dropdown list in sidebar
company_filter = st.sidebar.selectbox("Select the Company",companies)


#function to return index of selected company
def inter():
    count = 0
    for i in companies:
        if company_filter == i:
            count+=0
            x=ls[count]
        else:
            count+=1

    return x


#----- MAIN PAGE-----
st.title(":bar_chart: SDR DASHBOARD")
st.markdown("##")

#-----TABS------
tab1, tab2, tab3, tab4 = st.tabs([ "GENERAL", "lIQUIDITY", "PROFITABILITY",'INVESTMENT'])

with tab1:

    st.markdown('##')

# ------ TOP KPIs-------
    kpi1, kpi2, kpi3 =st.columns(3)

    with kpi1:
           my_currency = inter()['SDR1'].loc[7,'Year']
           desired_representation = "{:,}".format(my_currency)

           st.metric(label="Total Capital Component",
                      value='GHS ' + desired_representation,
                     delta=5
                     )
           with st.expander('Open to see more'):
               st.markdown('This is a brief overview')
               st.table(inter()['SDR1'].loc[1:7])

    with kpi1:

           st.metric(label='Capital Adequacy Ratio',
                     value=round(inter()['SDR1'].loc[92, 'Year']),
                     delta=150
                     )
           with st.expander('Open to see more'):
               st.write('This is a brief overview')
               st.table(inter()['SDR1'].loc[81:92])

    CAR = []
    COMPANY = []

    for i in range(1):
        if float(ls[i]['SDR1'].loc[92, 'Year']) >= 150:
            CAR.append(ls[i]['SDR1'].loc[92, 'Year']),
            COMPANY.append(companies[i])

    data = {
        "CAR Value": CAR,
        "Company": COMPANY
    }
    df = pd.DataFrame(data)
    st.table(df)

with tab2:
   st.header("LIQUIDITY INFO FOR" + " " + company_filter)
   st.markdown('##')

   with st.container():


       # USING TOTAL ASSSETS IN PLACE OF LIQUID ASSETS
       # a= Total technical reserve
       # b= Total assets
       a = inter()['SDR2i'].loc[7, 'Current Year']
       b = inter()['SDR2'].loc[58, 'Annual, 2020']
       st.metric(label='Technical Reserve Cover',
                 value=round(int(a) / float(b),2),
                 delta=150,
                 )
       with st.expander('Open to see more'):
           st.write('This is a brief overview')
           st.table(inter()['SDR1'].loc[81:92])


       st.metric(label='Land & Building As Investment',
                 value=inter()['SDR2'].loc[15, 'Annual, 2020'],
                 delta=None
                 )



with tab3:
#profitability Ratios

       kpi1,kpi2,kpi3 = st.columns(3)

       #Underwriting kpi
       with kpi1:
           my_currency = inter()['SDR3'].loc[28, 'Current Year']
           desired_representation = "{:,}".format(my_currency)
           my_currency1 = inter()['SDR3'].loc[28, 'Prior Year']
           desired_representation1 = "{:,}".format(my_currency1)
           st.metric(label="Underwriting Results",
              value='GHS' + desired_representation,
              delta='GHS' + desired_representation1
                     )
           with st.expander('Open to see more'):
                st.markdown('This is a brief overview')
                st.table(inter()['SDR3'].loc[11:28])

#Management Expenses
           my_currency = inter()['SDR3'].loc[23, 'Current Year']
           desired_representation = "{:,}".format(my_currency)
           my_currency1 = inter()['SDR3'].loc[23, 'Prior Year']
           desired_representation1 = "{:,}".format(my_currency1)
           st.metric(label="Management Expenses",
                     value="GHS "+ desired_representation,
                     delta="GHS "+ desired_representation1
                     )
           with st.expander('Open to see more'):
               st.markdown('This is a brief overview')
               st.table(inter()['SDR3i'].loc[2:23])

       with kpi2:
# Profit After Tax
            my_currency = inter()['SDR3'].loc[46, 'Current Year']
            desired_representation = "{:,}".format(my_currency)
            my_currency1 = inter()['SDR3'].loc[46, 'Prior Year']
            desired_representation1 = "{:,}".format(my_currency1)
            st.metric(label='Profit After Tax',
                      value="GHS " + desired_representation,
                      delta="GHS " + desired_representation1
                      )
            with st.expander('Open to see more'):
                st.markdown('This is a brief overview')
                st.table(inter()['SDR3'].loc[28:46])
       with kpi3:
#Gross Premiums Written
            my_currency = inter()['SDR3'].loc[5, 'Current Year']
            desired_representation = "{:,}".format(my_currency)
            my_currency1 = inter()['SDR3'].loc[5, 'Prior Year']
            desired_representation1 = "{:,}".format(my_currency1)
            st.metric(label='Gross Premiums Written',
                      value="GHS " + desired_representation,
                      delta="GHS " + desired_representation1
                  )
            with st.expander('Open to see more'):
                st.markdown('This is a brief overview')
                st.table(inter()['SDR8i'].loc[0:13])


with tab4:
    # Total Investment kpi
    my_currency = round(inter()['SDR3'].loc[37, 'Current Year'],2)
    desired_representation = "{:,}".format(my_currency)
    my_currency1 = inter()['SDR3'].loc[5, 'Prior Year']
    desired_representation1 = "{:,}".format(my_currency1)
    st.metric(label='Total Investment Income',
              value="GHS " + desired_representation,
              delta=None
              )
    with st.expander('Open to see more'):
            st.write('This is a brief overview of Investment sources')
            st.table(inter()['SDR3'].loc[30:37])

    # Interest Income kpi
    my_currency = inter()['SDR3'].loc[32, 'Current Year']
    desired_representation = "{:,}".format(my_currency)
    st.metric(label='Interest Income',
                  value="GHS " + desired_representation,
                  delta=None
                  )
    with st.expander('Open to see more'):
                st.write('This is a brief overview of Interest Income')
                st.table(inter()['SDR3'].loc[30:37])
#Total Investment
    my_currency = round(inter()['SDR2'].loc[19,'Annual, 2020'],2)
    desired_representation = "{:,}".format(my_currency)
    my_currency1 = inter()['SDR2'].loc[19,'Annual, 2019']
    desired_representation1 = "{:,}".format(my_currency1)
    st.metric(label="Total Investment",
              value="GHS " + desired_representation,
              delta="GHS " + desired_representation1
    )
    with st.expander('Open to see more'):
                st.write('This is a brief overview of Investments')
                st.table(inter()['SDR7'].loc[0:14])
#GOG Securities
    my_currency = inter()['SDR2'].loc[5, 'Annual, 2020']
    desired_representation = "{:,}".format(my_currency)
    my_currency1 = inter()['SDR2'].loc[5, 'Annual, 2019']
    desired_representation1 = "{:,}".format(my_currency1)
    st.metric(label="GOG Securities",
              value="GHS " + desired_representation,
              delta="GHS " + desired_representation1
              )
    with st.expander('Open to see more'):
        st.write('This is a brief overview of Investments')
        st.table(inter()['SDR7'].loc[0:1])


#Investment Yield
cash= inter()['SDR2'].loc[2,'Annual, 2020']
cash1= inter()['SDR2'].loc[2,'Annual, 2019']

with tab4:
            st.metric(label="Investment Yield",
              value=round(inter()['SDR3'].loc[37, 'Current Year']/(float(cash)+float(inter()['SDR2'].loc[19,'Annual, 2020'])),2),
              delta=inter()['SDR2'].loc[5, 'Annual, 2019']
              )
