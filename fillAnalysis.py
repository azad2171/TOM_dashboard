import pandas as pd
import warnings
import matplotlib.pyplot as plt

tickSize = {
    'BRN': 0.01,
    'CL': 0.01,
    'WBS': 0.01,
    'Box': 0.01,
    'ZM': 0.1,
    'ZL': 0.01,
    'ZS': 0.25,
    'ZC': 0.25,
    'ZW': 0.25,
    'HO': 0.0001,
    'RB': 0.0001,
    'NG': 0.001,
    'G': 0.25,
    'NG|HH': 0.00025
}

tickVal = {
    'BRN': 10,
    'Box': 10,
    'CL': 10,
    'WBS': 10,
    'ZM': 10,
    'ZL': 6,
    'ZS': 12.5,
    'ZC': 12.5,
    'ZW': 12.5,
    'HO': 4.2,
    'RB': 4.2,
    'NG': 10,
    'G': 25,
    'NG|HH': 2.5
}

def dashToNumberFormat(dashVal):
    mapping = {'6': '.75',
               '4': '.5',
               '2': '.25',
               '0': '.0'}
    try:
        pre, suf = dashVal.split("'")
        return float(pre + mapping[suf])
    except:
        return dashVal

warnings.simplefilter('ignore')
def get_fill_analysis_df():
    df = pd.read_excel('fillSheet.xlsx', 'fills')
    df['Price'] = df['Price'].apply(lambda x: dashToNumberFormat(x))
    df['qty*price'] = df['FillQty'] * df['Price']
    cont_grp = df.groupby('Contract')
    contracts_df = pd.DataFrame(cont_grp.groups.keys())

    outDf = pd.DataFrame(columns=['contract', 'netPos','buyQty', 'sellQty', 'avgBuyPrice', 'avgSellPrice'])
    for name, group in cont_grp:
        try:

            total_buy_qty   = group.loc[group['B/S'] == 'B', 'FillQty'].sum()
            total_sell_qty  = group.loc[group['B/S'] == 'S', 'FillQty'].sum()

            netPos = total_buy_qty - total_sell_qty

            avg_buy  = (group.loc[group['B/S'] == 'B', 'qty*price'].sum())/total_buy_qty 
            avg_sell = (group.loc[group['B/S'] == 'S', 'qty*price'].sum())/total_sell_qty

            outDf.loc[len(outDf.index)] = [name, netPos, total_buy_qty, total_sell_qty, avg_buy, avg_sell]
        except:
            # print(group)
            pass

    outDf['product'] = outDf['contract'].apply(lambda x: x.split()[0])
    outDf.loc[outDf['contract'].str.contains('HH'), 'product'] = 'NG|HH'
    outDf['tickSize'] = outDf['product'].apply(lambda x: tickSize[x])
    outDf['tickVal'] = outDf['product'].apply(lambda x: tickVal[x])
    outDf['ticksBooked'] = (outDf['avgSellPrice'] - outDf['avgBuyPrice'])/outDf['tickSize']
    outDf['pnl'] = outDf['ticksBooked']*outDf['tickVal']*outDf['buyQty']
    outDf['pnl'] = outDf['pnl'].round(1)
    outDf.sort_values(by='pnl', inplace=True, ignore_index=True)
    
    return outDf
