import xlwings as xw
import pandas as pd
import streamlit as st
import time

products = ['BRN', 'BOX', 'HO', 'RB', 'ZL', 'ZS', 'ZM', 'ZW', 'ZC', 'NGHH']

def get_summary_from_sheet(product_sheet):
    product_df = pd.DataFrame(product_sheet)
    product_df.columns = product_df.iloc[0]
    product_df = product_df[1:]
    product_df.dropna(axis='index', how='all', inplace=True)
    product_df.fillna(0, inplace=True)
    product_contracts_group = product_df.groupby('Contract')


    product_out_df = pd.DataFrame(columns=['contract', 'netOpenPos', 'buyQty', 'sellQty', 'total_pl', 'running_pl', 'booked_pl', 'MidPrice'])
    for contract, group in product_contracts_group:
        try:
            netOpenPos = group['Open Position'].sum()
            total_pl   = group['Total PnL'].sum()
            running_pl = group['Running PnL'].sum()
            booked_pl  = group['Booked PnL'].sum()
            buyQty     = group['BuyTotalQty'].sum()
            sellQty    = group['SellTotalQty'].sum()
            midPrice   = group['MidPrice'].iloc[0]

            product_out_df.loc[len(product_out_df.index)] = [str(contract), str(netOpenPos), str(buyQty), str(sellQty), str(total_pl), str(running_pl), str(booked_pl), str(midPrice)]
        except:
            pass

    return product_out_df



def color_backgroubd_red_column(col):
    returnList = []
    for val in col:
        try:
            if int(val) > 0:
                returnList.append('background-color: green')
            elif int(val) < 0:
                returnList.append('background-color: red')
            else:
                returnList.append('background-color: None')
        except:
            returnList.append('background-color: None')

    return returnList



st.set_page_config(page_title = "Dashboard", layout="wide")

tabs = st.tabs(products + ['Main'])
placeholder = st.empty()

tabs_placeholder = []

for tab in tabs:
    tabs_placeholder.append(tab.empty())

# Set Streamlit option to show full DataFrame
st.set_option('deprecation.showPyplotGlobalUse', False)

while True:
    time.sleep(1)
    try:
        with placeholder.container():
            main_out_df = pd.DataFrame()
            wb = xw.Book('C:/Users/lovekesh.azad/Downloads/ASE_Manage/Autospreader_monitor_non_macro.xlsx')
            
            for sheet in wb.sheets:
                if sheet.name in products:
                    product_sheet = sheet.used_range.value
                    df = get_summary_from_sheet(product_sheet)
                    
                    
                    main_out_df = pd.concat([main_out_df, df], ignore_index=True)
                    
                    with tabs_placeholder[products.index(sheet.name)]:
                        st.dataframe(df)

                    # with tabs_placeholder[-1]:
                    #     st.dataframe(main_out_df)


        main_sheet = wb.sheets['Main']
        main_sheet.range('A1').value = main_out_df
    except Exception as e:
        print(e)
        time.sleep(2)