import pandas as pd
import numpy as np


all_statements_path = 'Classeur3.xlsx'
def load_and_clean_statement_df(statements_path, sheet_name):
    df = pd.read_excel(statements_path, sheet_name=sheet_name,index_col=0)
    df = df.replace('-', np.nan)
    df = df.dropna(how='all')
    df = df.fillna(0) 
    return df
inc_df = load_and_clean_statement_df(all_statements_path, 'Income Statement')
ca_df = load_and_clean_statement_df(all_statements_path, 'Cash Flow')
bs_df = load_and_clean_statement_df(all_statements_path, 'Balance Sheet')



Currents_Assets=""
Currents_Liabilities=""
ocl_list=bs_df.loc['Other Current Liabilities']
if 'Unearned Revenue, Current' in bs_df.index:
    ocl_list += bs_df.loc['Unearned Revenue, Current'] 
Net_current_assets=bs_df.loc['  Total Receivables']+bs_df.loc['Inventory']+bs_df.loc['Prepaid Exp.']+bs_df.loc['Other Current Assets']
Net_Current_Liabilities=bs_df.loc['Accounts Payable']+bs_df.loc['Accrued Exp.']+bs_df.loc['Curr. Income Taxes Payable']+ocl_list
NWC=Net_current_assets - Net_Current_Liabilities
Change_nwc = NWC - NWC.shift(1)
Change_nwc = Change_nwc.fillna(0)

revenue_list=inc_df.loc['Revenue']
ar_list=bs_df.loc['  Total Receivables']
inv_list=bs_df.loc['Inventory']
peoca_list=bs_df.loc['Prepaid Exp.']

ap_list=bs_df.loc['Accounts Payable']
if 'Unearned Revenue, Current' in bs_df.index:
    ocl_list += bs_df.loc['Unearned Revenue, Current']
cost_list=inc_df.loc['Cost Of Goods Sold']
ae_list=bs_df.loc['Accrued Exp.']
oca_list=bs_df.loc['Other Current Assets']
taxe_list=bs_df.loc['Curr. Income Taxes Payable']
dso_list = []
for i in range(len(revenue_list)):
    if revenue_list[i] == 0:
        dso_list.append(0)
    else:
        dso_list.append(365*(ar_list[i]/revenue_list[i]))
        
dih_list = []
for i in range(len(cost_list)):
    if cost_list[i] == 0:
        dih_list.append(0)
    else:
        dih_list.append(365*(inv_list[i]/cost_list[i]))

Other_Current_Assets_list = []
for i in range(len(revenue_list)):
    if revenue_list[i] == 0:
        Other_Current_Assets_list.append(0)
    else:
        Other_Current_Assets_list.append((oca_list[i]/revenue_list[i]))

dpo_list = []
for i in range(len(cost_list)):
    if cost_list[i] == 0:
        dpo_list.append(0)
    else:
        dpo_list.append(365*(ap_list[i]/cost_list[i]))

Accrued_Liabilites_list = []
for i in range(len(revenue_list)):
    if revenue_list[i] == 0:
        Accrued_Liabilites_list.append(0)
    else:
        Accrued_Liabilites_list.append((ae_list[i]/revenue_list[i]))

Other_Current_Liabilities_list = []
for i in range(len(revenue_list)):
    if revenue_list[i] == 0:
        Other_Current_Liabilities_list.append(0)
    else:
        Other_Current_Liabilities_list.append((ocl_list[i]/revenue_list[i]))

Taxes_Payable_list = []
for i in range(len(revenue_list)):
    if revenue_list[i] == 0:
        Taxes_Payable_list.append(0)
    else:
        Taxes_Payable_list.append((taxe_list[i]/revenue_list[i]))

inv_turnover_list = []
for i in range(len(revenue_list)):
    if inv_list[i] == 0:
        inv_turnover_list.append(0)
    else:
        inv_turnover_list.append(inv_list[i]/cost_list[i])

receivable_turnover_list = []
for i in range(len(revenue_list)):
    if ar_list[i] == 0:
        receivable_turnover_list.append(0)
    else:
        receivable_turnover_list.append(ar_list[i]/revenue_list[i])

        
payable_turnover_list = []
for i in range(len(revenue_list)):
    if ap_list[i] == 0:
        payable_turnover_list.append(0)
    else:
        payable_turnover_list.append(ap_list[i]/cost_list[i])
        
prepaid_turnover_list=[]
for i in range(len(revenue_list)):
    if ap_list[i] == 0:
        prepaid_turnover_list.append(0)
    else:
        prepaid_turnover_list.append(peoca_list[i]/revenue_list[i])
        


NWC_df = pd.DataFrame({'Revenue': inc_df.loc['Revenue'],
        'COGS': inc_df.loc['Cost Of Goods Sold'],
        'Currents_Assets':Currents_Assets,
        'Accounts_Receivable': bs_df.loc['  Total Receivables'],
        'Inventory': bs_df.loc['Inventory'],
        'Prepaid_Exp.': bs_df.loc['Prepaid Exp.'],
        'Other_Current_Assets': bs_df.loc['Other Current Assets'],
        'Net_current_assets':Net_current_assets,
        'Currents_Liabilities':Currents_Liabilities,
        'Accounts Payable':bs_df.loc['Accounts Payable'],
        'AccruedExp.':bs_df.loc['Accrued Exp.'],
        'Curr.IncomeTaxesPayable':bs_df.loc['Curr. Income Taxes Payable'],
        'OtherCurrentLiabilities':ocl_list,
        'Net_Current_Liabilities':Net_Current_Liabilities,
        'NWC':NWC,       
        'Change_nwc':Change_nwc,
        '':'',
    'Account receivables in % of Sales': receivable_turnover_list,
    'DSO': dso_list,
    'Inventory turnover': inv_turnover_list,
    'DIH': dih_list,
    'Prepaid Expenses in % of Sales':prepaid_turnover_list,
    'Other Current Assets': Other_Current_Assets_list,
    'Account payable in % of COGS': payable_turnover_list,              
    'DPO': dpo_list,
    'Accrued_Liabilites in % of Sales':Accrued_Liabilites_list,
    'Other Current Liabilities in % of Sales':Other_Current_Liabilities_list,
    'Taxes Payable in % of Sales':Taxes_Payable_list
       })

NWC_df = NWC_df.T.reset_index(drop=True)
NWC_df.columns = inc_df.columns
NWC_df.index=['Revenue','COGS',
                'Currents Assets','Accounts Receivable',
                'Inventory','Prepaid Exp.','Other Current Assets','Net current assets',
                'Currents Liabilities','Accounts Payable','Accrued Exp.','Curr.Income Taxes Payable','Other Current Liabilities','Net Current Liabilities',
                'NWC','Change_nwc','','Account receivables in % of Sales','DSO','Inventory turnover','DIH','Prepaid Expenses in % of Sales','Other Current Assets in % of Sales','Account payable in % of COGS','DPO','Accrued_Liabilites in % of Sales',
                'Other Current Liabilities in % of Sales','Taxes Payable in % of Sales']

        
Gross_profit=(inc_df.loc['Revenue']-inc_df.loc['Cost Of Goods Sold'])
Opex = inc_df.loc['Selling General & Admin Exp.'] + inc_df.loc['R & D Exp.'] + inc_df.loc['Other Operating Expense/(Income)'] 
Ebit_= Gross_profit - Opex
NOPAT= Ebit_-(inc_df.loc['Income Tax Expense'])
CFO=NOPAT - Change_nwc + ca_df.loc['Depreciation & Amort., Total']


if 'Stock-Based Compensation' in ca_df.index:
    CFO += ca_df.loc['Stock-Based Compensation']

if '(Gain) Loss On Sale Of Invest.' in ca_df.index:
    CFO += ca_df.loc['(Gain) Loss On Sale Of Invest.']
    
if 'Asset Writedown & Restructuring Costs' in ca_df.index:
    CFO += ca_df.loc['Asset Writedown & Restructuring Costs']
    
if 'Other Operating Activities' in ca_df.index:
    CFO += ca_df.loc['Other Operating Activities']
FCFF=CFO- ca_df.loc['Capital Expenditure']


FCFF_bis_df = pd.DataFrame({'Revenue': inc_df.loc['Revenue'],
        'COGS': inc_df.loc['Cost Of Goods Sold'],
        'Gross_profit':Gross_profit,
        'SGA': inc_df.loc['Selling General & Admin Exp.'],
        'RD': inc_df.loc['R & D Exp.'],
        'Other_Operating_Expense': inc_df.loc['Other Operating Expense/(Income)'],
        'DA': ca_df.loc['Depreciation & Amort., Total'],
        'Opex':Opex,
        'Ebit_':Ebit_,
        'Income_Tax_Expense': inc_df.loc['Income Tax Expense'],
        'Net_income':NOPAT,
        'Depreciation_and_Amortization': ca_df.loc['Depreciation & Amort., Total'],
        'Stock-BasedCompensation': ca_df.loc['Stock-Based Compensation'],
        'OtherOperatingActivities': ca_df.loc['Other Operating Activities'],
        '(Gain)LossOnSaleOfInvest.': ca_df.get('(Gain) Loss On Sale Of Invest.', 0),
        'Asset_Writedown_Restructuring_Costs': ca_df.get('Asset Writedown & Restructuring Costs', 0),
        'Change_nwc':Change_nwc,
        'CFO': CFO,
        'Capital_Expenditure':ca_df.loc['Capital Expenditure'],
        'FCFF':FCFF
       })
FCFF_bis_df=FCFF_bis_df.T
FCFF_df = FCFF_bis_df.reset_index(drop=True)
FCFF_df.columns = inc_df.columns
FCFF_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','Net Income','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']

list_variables = ['Revenue', 'COGS', 'Gross_profit', 'SGA', 'RD',
       'Other_Operating_Expense', 'DA', 'Opex', 'Ebit_', 'Income_Tax_Expense',
       'Net_income', 'Depreciation_and_Amortization',
       'Stock-BasedCompensation', 'OtherOperatingActivities',
       '(Gain)LossOnSaleOfInvest.', 'Asset_Writedown_Restructuring_Costs',
       'Change_nwc', 'CFO', 'Capital_Expenditure', 'FCFF']
variables = []

for variable in list_variables:
    x = FCFF_bis_df.loc[variable]/FCFF_bis_df.loc['Revenue']
    variables.append(x)
    
    
FCFF_bis_vertical=pd.DataFrame(variables)
FCFF_bis_vertical.index=['Revenue', 'COGS', 'Gross_profit', 'SGA', 'RD',
       'Other_Operating_Expense', 'DA', 'Opex', 'Ebit_', 'Income_Tax_Expense',
       'Net_income', 'Depreciation_and_Amortization',
       'Stock-BasedCompensation', 'OtherOperatingActivities',
       '(Gain)LossOnSaleOfInvest.', 'Asset_Writedown_Restructuring_Costs',
       'Change_nwc', 'CFO', 'Capital_Expenditure', 'FCFF']
FCFF_vertical=FCFF_bis_vertical.reset_index(drop=True)
FCFF_vertical.columns = inc_df.columns
FCFF_vertical.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','Net Income','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']

import statistics as stat
avg_acr = stat.mean(receivable_turnover_list)
avg_dso = stat.mean(dso_list)
avg_inv = stat.mean(inv_turnover_list)
avg_dih = stat.mean(dih_list)
avg_pre = stat.mean(prepaid_turnover_list)
avg_oca = stat.mean(Other_Current_Assets_list)
avg_acp = stat.mean(payable_turnover_list)
avg_dpo = stat.mean(dpo_list)
avg_acl = stat.mean(Accrued_Liabilites_list)
avg_ocl = stat.mean(Other_Current_Liabilities_list)
avg_taxe = stat.mean(Taxes_Payable_list)

median_acr = stat.mean(receivable_turnover_list)
median_dso = stat.median(dso_list)
median_inv = stat.median(inv_turnover_list)
median_dih = stat.median(dih_list)
median_pre = stat.median(prepaid_turnover_list)
median_oca = stat.median(Other_Current_Assets_list)
median_acp = stat.mean(payable_turnover_list)
median_dpo = stat.median(dpo_list)
median_acl = stat.median(Accrued_Liabilites_list)
median_ocl = stat.median(Other_Current_Liabilities_list)
median_taxe = stat.median(Taxes_Payable_list)

last_acr = receivable_turnover_list[-1]
last_dso = dso_list[-1]
last_inv = inv_turnover_list[-1]
last_dih = dih_list[-1]
last_pre = prepaid_turnover_list[-1]
last_oca = Other_Current_Assets_list[-1]
last_acp = payable_turnover_list[-1]
last_dpo = dpo_list[-1]
last_acl = Accrued_Liabilites_list[-1]
last_ocl = Other_Current_Liabilities_list[-1]
last_taxe = Taxes_Payable_list[-1]
stat_NWC_df = pd.DataFrame({
    'Average': [avg_acr,avg_dso, avg_inv, avg_dih,avg_pre, avg_oca, avg_acp, avg_dpo, avg_acl, avg_ocl, avg_taxe],
    'Median': [median_acr, median_dso, median_inv, median_dih, median_pre, median_oca, median_acp, median_dpo, median_acl, median_ocl, median_taxe],
    'Last Value': [last_acr, last_dso, last_inv, last_dih, last_pre, last_oca, last_acr, last_dpo, last_acl, last_ocl, last_taxe]
})
stat_NWC_df.index=['Account receivables in % of Sales','DSO','Inventory Turnover','DIH', 'Prepaid Expenses in % of Sales', 'Other Current Assets in % of Sales','Account payable in % of COGS','DPO','Accrued_Liabilites in % of Sales',
                'Other Current Liabilities in % of Sales','Taxes Payable in % of Sales']



list_variables = ['Revenue', 'COGS', 'Gross_profit', 'SGA', 'RD',
       'Other_Operating_Expense', 'DA', 'Opex', 'Ebit_', 'Income_Tax_Expense',
       'Net_income', 'Depreciation_and_Amortization',
       'Stock-BasedCompensation', 'OtherOperatingActivities',
       '(Gain)LossOnSaleOfInvest.', 'Asset_Writedown_Restructuring_Costs',
       'Change_nwc', 'CFO', 'Capital_Expenditure', 'FCFF']
avg_variable=[]
median_variable=[]
last_variable=[]
for variables in list_variables: 
    variables_list=FCFF_bis_df.loc[variables]
    avg_variable.append(stat.mean(variables_list))
    median_variable.append(stat.median(variables_list))
    last_variable.append(variables_list[-1])
stat_FCFF_df = pd.DataFrame({
    'Average': avg_variable,
    'Median': median_variable,
    'Last Value': last_variable
}) 
stat_FCFF_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','Net Income','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']

list_variables = ['Revenue', 'COGS', 'Gross_profit', 'SGA', 'RD',
       'Other_Operating_Expense', 'DA', 'Opex', 'Ebit_', 'Income_Tax_Expense',
       'Net_income', 'Depreciation_and_Amortization',
       'Stock-BasedCompensation', 'OtherOperatingActivities',
       '(Gain)LossOnSaleOfInvest.', 'Asset_Writedown_Restructuring_Costs',
       'Change_nwc', 'CFO', 'Capital_Expenditure', 'FCFF']
avg_pct_variable=[]
median_pct_variable=[]
last_pct_variable=[]
for variables in list_variables: 
    variables_list=FCFF_bis_vertical.loc[variables]
    avg_pct_variable.append(stat.mean(variables_list))
    median_pct_variable.append(stat.median(variables_list))
    last_pct_variable.append(variables_list[-1])
stat_FCFF_vertical_df = pd.DataFrame({
    'Average': avg_pct_variable,
    'Median': median_pct_variable,
    'Last Value': last_pct_variable
}) 
stat_FCFF_vertical_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','Net Income','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']

Assumption_revenues_list=[]
Assumption_revenues_df=pd.DataFrame(Assumption_revenues_list)
Assumption_revenues_df.index=['2023','2024','2025','2026','2027','2028']
labels_revenues=['Sales Estelle','License Estelle','Sales Donesta','License Donesta','Sales Myring','License Myring','Others']
Assumption_revenues_df['Estelle']=[47.93,64.14,113.51,92.43,132.64,192.96]
Assumption_revenues_df['Donesta']=[15,12.8,34.9,14.2,31.1,33.3]
Assumption_revenues_df['Myring']=[7.14,8.94,14.92,10.64,11.79,12.74]
Assumption_revenues_df['Other']=[0.92,1.41,2.15,3.29,5.04,7.72]
total_sales_list=[]
for year in Assumption_revenues_df.index:
    total_sales_list.append(sum(Assumption_revenues_df.T[year]))
    
Assumption_revenues_df['Total']=total_sales_list
Assumption_revenues_df=Assumption_revenues_df.T




stat.mean(FCFF_vertical.loc['SG&A',['2017-12-31','2018-12-31','2019-12-31','2021-12-31','2022-12-31']])
        
        
choice_list=[]
draft_futurV_FCFF_df=pd.DataFrame(choice_list)

draft_futurV_FCFF_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A1','Opex',
                'EBIT','Income tax expenses','Net Income','D&A2',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']
year_to_forecast=['2023','2024','2025','2026','2027','2028']

for year in year_to_forecast:
    draft_futurV_FCFF_df[year]=avg_pct_variable
    draft_futurV_FCFF_df[year]['SG&A']=stat.mean(FCFF_vertical.loc['SG&A',['2017-12-31','2018-12-31','2019-12-31','2021-12-31','2022-12-31']])
    draft_futurV_FCFF_df[year]['R&D']=stat.mean(FCFF_vertical.loc['R&D',['2018-12-31','2019-12-31']])
    draft_futurV_FCFF_df[year]['Other Operating Expense']=stat.mean(FCFF_vertical.loc['Other Operating Expense',['2017-12-31','2018-12-31','2019-12-31','2021-12-31','2022-12-31']])
    draft_futurV_FCFF_df[year]['Opex']=sum([draft_futurV_FCFF_df[year]['SG&A'],draft_futurV_FCFF_df[year]['R&D'],draft_futurV_FCFF_df[year]['Other Operating Expense']])
    draft_futurV_FCFF_df[year]['Capital Expenditure']=stat.mean(FCFF_vertical.loc['Capital Expenditure',['2017-12-31','2018-12-31','2019-12-31','2021-12-31','2022-12-31']])
    draft_futurV_FCFF_df[year]['D&A1']=stat.mean(FCFF_bis_vertical.loc['DA',['2019-12-31','2020-12-31','2021-12-31',]]/-FCFF_vertical.loc['Capital Expenditure',['2019-12-31','2020-12-31','2021-12-31']])
    draft_futurV_FCFF_df[year]['D&A2']=stat.mean(FCFF_bis_vertical.loc['DA',['2019-12-31','2020-12-31','2021-12-31',]]/-FCFF_vertical.loc['Capital Expenditure',['2019-12-31','2020-12-31','2021-12-31']])

draft_futurV_FCFF_df.loc['COGS']=[0.2874, 0.2662, 0.1834, 0.2083, 0.1627, 0.1289]
draft_futurV_FCFF_df.loc['Income tax expenses']=stat.median(np.absolute(FCFF_df.loc['Income tax expenses',['2017-12-31','2018-12-31','2019-12-31','2020-12-31','2021-12-31']]/FCFF_df.loc['EBIT',['2017-12-31','2018-12-31','2019-12-31','2020-12-31','2021-12-31']]))

forecasted_revenue_list=Assumption_revenues_df.loc['Total']
forecasted_cost_list=draft_futurV_FCFF_df.loc['COGS']*Assumption_revenues_df.loc['Total']
forecasted_ar_list=stat_NWC_df.loc['Account receivables in % of Sales','Average']*forecasted_revenue_list
forecasted_inv_list=stat_NWC_df.loc['Inventory Turnover','Average']*forecasted_cost_list
forecasted_peoca_list=stat_NWC_df.loc['Prepaid Expenses in % of Sales','Average']*forecasted_revenue_list
forecasted_ap_list=stat_NWC_df.loc['Account payable in % of COGS','Average']*forecasted_cost_list
forecasted_ae_list=stat_NWC_df.loc['Accrued_Liabilites in % of Sales','Average']*forecasted_revenue_list
forecasted_oca_list=stat_NWC_df.loc['Other Current Assets in % of Sales','Average']*forecasted_revenue_list
forecasted_taxe_list=stat_NWC_df.loc['Taxes Payable in % of Sales','Average']*forecasted_revenue_list
forecasted_ocl_list=stat_NWC_df.loc['Other Current Liabilities in % of Sales','Average']*forecasted_revenue_list
forecasted_net_current_assets=forecasted_ar_list+forecasted_inv_list+forecasted_peoca_list+forecasted_oca_list
forecasted_net_current_liabilities=forecasted_ap_list+forecasted_ae_list+forecasted_taxe_list+forecasted_ocl_list
forecasted_NWC=forecasted_net_current_assets-forecasted_net_current_liabilities
forecasted_change_NWC= forecasted_NWC - forecasted_NWC.shift(1)
forecasted_change_NWC_2023=forecasted_NWC['2023']-NWC['2022-12-31']
forecasted_change_NWC=forecasted_change_NWC.fillna(forecasted_change_NWC_2023)
forecasted_NWC_df = pd.DataFrame({'Revenue': forecasted_revenue_list,
        'COGS': forecasted_cost_list,
        'Currents_Assets':Currents_Assets,
        'Accounts_Receivable': forecasted_ar_list,
        'Inventory': forecasted_inv_list,
        'Prepaid_Exp.': forecasted_peoca_list,
        'Other_Current_Assets': forecasted_oca_list,
        'Net_current_assets':forecasted_net_current_assets,
        'Currents_Liabilities':Currents_Liabilities,
        'Accounts Payable':forecasted_ap_list,
        'AccruedExp.':forecasted_ae_list,
        'Curr.IncomeTaxesPayable':forecasted_taxe_list,
        'OtherCurrentLiabilities':forecasted_ocl_list,
        'Net_Current_Liabilities':forecasted_net_current_liabilities,
        'NWC':forecasted_NWC,       
        'Change_nwc':forecasted_change_NWC})
forecasted_NWC_df=forecasted_NWC_df.T

list_variables = ['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A1','Opex',
                'EBIT','Income tax expenses','Net Income','D&A2',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']
variables = []

for variable in list_variables:
    x = draft_futurV_FCFF_df.loc[variable]*Assumption_revenues_df.loc['Total']
    variables.append(x)
Futur_FCFF_df=pd.DataFrame(variables)
Futur_FCFF_df.columns = draft_futurV_FCFF_df.columns

Futur_FCFF_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','NOPAT','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']
Futur_FCFF_df.loc['Change nwc']=forecasted_NWC_df.loc['Change_nwc']
Futur_FCFF_df.loc['Gross_profit']=Futur_FCFF_df.loc['Revenue']-Futur_FCFF_df.loc['COGS']
Futur_FCFF_df.loc['EBIT']=Futur_FCFF_df.loc['Gross_profit']-Futur_FCFF_df.loc['Opex']
Futur_FCFF_df.loc['NOPAT']=Futur_FCFF_df.loc['EBIT']-Futur_FCFF_df.loc['Income tax expenses']
Futur_FCFF_df.loc['CFO']=Futur_FCFF_df.loc['NOPAT']+Futur_FCFF_df.loc['Stock Based Compensation']+Futur_FCFF_df.loc['(Gain) Loss On Sale Of Invest']+Futur_FCFF_df.loc['Asset Writedown Restructuring Costs']-Futur_FCFF_df.loc['Change nwc']
Futur_FCFF_df.loc['FCFF']=Futur_FCFF_df.loc['CFO']-Futur_FCFF_df.loc['Capital Expenditure']
Futur_FCFF_df.loc['Income tax expenses']=draft_futurV_FCFF_df.loc['Income tax expenses']*Futur_FCFF_df.loc['EBIT']

Futur_FCFF_df.columns = draft_futurV_FCFF_df.columns

Futur_FCFF_df.index=['Revenue','COGS','Gross_profit','SG&A','R&D','Other Operating Expense','D&A','Opex',
                'EBIT','Income tax expenses','Net Income','D&A',
                'Stock Based Compensation','Other Operating Activities',
                '(Gain) Loss On Sale Of Invest',
                'Asset Writedown Restructuring Costs',
                'Change nwc','CFO','Capital Expenditure','FCFF']




import requests
from bs4 import BeautifulSoup
import pandas as pd
# Send a GET request to the website URL
url = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html"
response = requests.get(url,verify=False)

soup = BeautifulSoup(response.content, 'html.parser')

table = soup.find('table')

rows = table.find_all('tr')

data = []
for row in rows:
    cols = row.find_all('td')
    cols = [col.text.strip() for col in cols]
    data.append(cols)
tb = pd.DataFrame(data)
tb.columns=(tb.iloc[0])
tb=tb.drop(index=[0,16])
tb.iloc[:, 0:1] = tb.iloc[:, 0:1].astype(float)
tb.iloc[:, 3] = (tb.iloc[:, 3].astype(str).str.replace('%', '').astype(float))/100

interest=inc_df.loc['  Net Interest Exp.']
ratio=Ebit_[-1]/interest[-1]

i=0
while i < len(tb.index):
    if float(tb.iloc[i, 0]) < ratio < float(tb.iloc[i, 1]):
        spread = tb.iloc[i, 3]
    i+=1
    

import yfinance as yf
url = "http://www.worldgovernmentbonds.com/"
response = requests.get(url)
soup = BeautifulSoup(response.content, "html.parser")
country='Belgium'
table = soup.find("table", {"class": "homeBondTable"})
for row in table.find_all("tr"):
    if country in row.text:
        rf = row.find("td", {"class": "w3-right-align w3-bold"}).text.strip()
        rf = (float(rf.strip("%")))/100
        break  
tot_debt=510330000
year=2023
debts = [
    [2023, 25796],
    [2024, 22996],
    [2025, 105761],
    [2028, 301637],
    [2033, 53141],
]
total_debt = sum(debt[1] for debt in debts)
w_avg_maturity = sum((debt[0] - year) * (debt[1] / total_debt) for debt in debts)
int_exp=inc_df.loc['Interest Expense']
int_exp = abs(int_exp[-1])
cost_of_debt=rf + spread
step_1 = (1 - (1/(1+cost_of_debt)**w_avg_maturity))/cost_of_debt
step_2 = tot_debt/(1+cost_of_debt)**w_avg_maturity
mv_debt = int_exp * step_1 + step_2
taxe=draft_futurV_FCFF_df.loc['Income tax expenses'][0]    
    

import requests
from bs4 import BeautifulSoup
import pandas as pd
# Send a GET request to the website URL
url = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html"
response = requests.get(url,verify=False)
soup = BeautifulSoup(response.content, 'html.parser')
table = soup.find("table", {"width": "919"})
for row in table.find_all("tr"):
    if country in row.text:
        erp = row.find_all("td")[2].text.strip()
        erp = (float(erp.strip("%")))/100
        break
        
mitra = yf.Ticker('APPL')
marketCap = 100000000000
beta=2.13
cost_of_equity = beta * erp + rf
cost_of_debt = rf + spread
cap_structure = marketCap + mv_debt
w_cost_equity = marketCap / cap_structure
w_cost_debt = mv_debt / cap_structure
wacc = cost_of_equity * w_cost_equity + cost_of_debt * w_cost_debt

wacc_tab= pd.DataFrame({
    'cost of equity':[cost_of_equity],
    'beta': [beta],
    'erp': [erp],
    'Market Value of Equity': [marketCap],
    'Weight of Equity':[w_cost_equity],
    'cost of debt':[cost_of_debt],
    'rf': [rf],
    'Spread': [spread],
    'Taxe rate':[taxe],
    'Market Value of Debt':[mv_debt],
    'Weight of Debt':[w_cost_debt],
    'WACC':[wacc]
})




import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import seaborn as sns
import plotly.express as px
import plotly.graph_objs as go
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import base64

def format_float(x):
    if isinstance(x, (float, int)):
        return '{:.2f}'.format(x)
    else:
        return x

df1 = pd.DataFrame(FCFF_df)
df1 = df1.reset_index()
df1.columns = ['Index'] + list(df1.columns[1:])


df2 = pd.DataFrame(NWC_df)
df2 = df2.reset_index()
df2.columns = ['Index'] + list(df2.columns[1:])

df3 = pd.DataFrame(Assumption_revenues_df)
df3 = df3.reset_index()
df3.columns = ['Index'] + list(df3.columns[1:])

df4 = pd.DataFrame(Assumption_revenues_df)
df4 = df4.reset_index()
df4.columns = ['Index'] + list(df4.columns[1:])

df5 = pd.DataFrame(Assumption_revenues_df)
df5 = df5.reset_index()
df5.columns = ['Index'] + list(df5.columns[1:])

df6=tb

df7= wacc_tab.transpose()
df7 = df7.applymap(format_float)

df11=FCFF_df
df11 = df11.applymap(format_float)
df22=NWC_df
df22 = df22.applymap(format_float)
df33=Assumption_revenues_df
df33 = df33.applymap(format_float)
df44=forecasted_NWC_df
df44 = df44.applymap(format_float)
df55=Futur_FCFF_df
df55 = df55.applymap(format_float)
df11.columns = ['2017', '2018', '2019','2020','2021','2022']
df22.columns = ['2017', '2018', '2019','2020','2021','2022']
df33.columns = ['2023', '2024', '2025','2026','2027','2028']
df44.columns = ['2023', '2024', '2025','2026','2027','2028']
df55.columns = ['2023', '2024', '2025','2026','2027','2028']


# Define the layout of the Streamlit app
st.title('Mitra Valuation')
menu = ['Company', 'FCFF', 'Produit', 'Futur FCFF', 'Spread', 'WACC']

choice = st.sidebar.selectbox('Choisir un tableau', menu)


if choice == 'Company':
    st.write("""
    <div style='text-align: justify'>
    Mithra is a leading Belgian pharmaceutical company specializing in women's and genital health. Founded in 1999 and headquartered in Liège, Belgium, the company is publicly listed on Euronext Brussels and Euronext Amsterdam.

    Mithra focuses on researching, developing, and commercializing innovative health products to meet the unmet needs of women and men worldwide. The company has a broad range of products, including hormonal contraceptives and treatments for menopause, fertility products, and treatments for infectious and inflammatory diseases.

    One of Mithra's flagship products is Estelle®, a combined contraceptive pill based on estetrol, a natural hormone produced by the fetus during pregnancy. Mithra has also developed a patented technology platform called Estetra®, which allows for the economic production of estetrol on a large scale.

    In addition to its research and development activities, Mithra is actively involved in promoting women's health and gender equality. Through its commitment to innovation, sustainability, and its broad range of health products for women and men worldwide, Mithra continues to advance the healthcare industry and strengthen its position as a global leader in women's and genital health.

    Mithra is a Belgian pharmaceutical company that focuses on researching, developing, and commercializing innovative health products to meet the unmet needs of women and men worldwide. The company's activities include:

    1) Research and development: Mithra invests heavily in research and development of new health products, particularly in women's and genital health. The company uses cutting-edge technologies to develop innovative products, such as hormonal contraceptives, menopause treatments, fertility products, and treatments for infectious and inflammatory diseases.

    2) Production: Mithra has state-of-the-art production facilities in Belgium to manufacture its pharmaceutical products. The company is committed to sustainable and environmentally friendly production, using production processes that minimize environmental impact.

    3) Commercialization: Mithra markets its pharmaceutical products in many countries around the world, working closely with local commercial partners and distributors. The company aims to expand its presence in emerging markets to meet the growing health needs of consumers.

    4) Commitment to women's health: Mithra is committed to promoting women's health and gender equality. The company participates in awareness campaigns, events, and initiatives to help improve women's health worldwide.
    </div>
    """, unsafe_allow_html=True)


elif choice == 'FCFF':
    st.subheader('FCFF')
    st.table(df11)

    y_value = st.selectbox('Choisir la valeur y', df1['Index'].unique())
    x_values = df22.columns[0:]
    y_values = df1.loc[df1['Index'] == y_value].values[0][1:]

    fig = go.Figure()
    fig.add_trace(go.Bar(x=x_values, y=y_values, marker_color='green'))
    fig.update_layout(title=y_value, xaxis_title='', yaxis_title='')
    st.plotly_chart(fig)


elif choice == 'NWC':
    st.subheader('NWC')
    st.table(df22)

    y_value = st.selectbox('Choisir la valeur y', df2['Index'].unique())
    x_values = df22.columns[0:]
    y_values = df2.loc[df2['Index'] == y_value].values[0][1:]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x_values, y=y_values, marker_color='green'))
    fig.update_layout(title=y_value, xaxis_title='', yaxis_title='')
    st.plotly_chart(fig)


    
elif choice == 'Produit':
    st.subheader('Produit')
    st.table(df33)
    Assumption_revenues_df = Assumption_revenues_df.T
    melted_df = pd.melt(Assumption_revenues_df, ignore_index=False, var_name='Product', value_name='Sales')
    melted_df.reset_index(inplace=True)
    melted_df.rename(columns={'index': 'Year'}, inplace=True)
    # Filter out 'Total Sales' rows
    melted_df = melted_df[melted_df['Product'] != 'Total']
    labels = {
    'Year': 'Year',
    'Product': 'Product',
    'Sales': 'Sales (in millions)',}
    color_scale = 'Blues'
    fig = px.sunburst(melted_df, path=['Year', 'Product'], values='Sales', color='Sales',
                  color_continuous_scale=color_scale, hover_name='Year', labels=labels,
                  branchvalues='total', title='Revenues by Year and Product')
    st.plotly_chart(fig)

    
elif choice == 'NWC Forecast':
    st.subheader('NWC Forecast')
    st.table(df44)
    
elif choice == 'Futur FCFF':
    st.subheader('Futur FCFF')
    st.table(df55)
        
        
elif choice == 'Spread':
    st.subheader('spread')
    st.table(df6)
    st.write(f"The value of the ratio is {ratio:.2}")
    st.write(f"The value of the spread is {spread:.2%}")
        
elif choice == 'WACC':
    st.subheader('WACC')
    st.table(df7)
    st.write(f"The value of the WACC is {wacc:.2%}")
        
        
        
