import pandas as pd
import win32com.client
from datetime import date
from datetime import timedelta
from sqlalchemy import create_engine
import os
from urllib.parse import quote

#date
today = date.today()
yesterday = (today - timedelta(days = 1)).strftime('%d%b%Y')
strt_day = (today - timedelta(days = 1)).strftime('01%b%Y')


#raw_data = pd.read_excel('data.xlsx')

cnx = create_engine('mysql+pymysql://analytics:%s@34.93.197.76:61306/analytics' % quote('Secure@123'))
query_1 = f"select branch,	URN,	AccountNumber,	funder_name,	funding_txn_type,	" \
          f"funding_txn_remark,	Product,	Account_Status,	Prin_OS_as_on_{yesterday} as POS,	" \
          f"Overdue_Days_as_on_{yesterday} as DPD_Days,	customer_name, DisbursementDate, Last_Repayment_Date" \
          f" from perdix_cdr.quick_cbs_loan_dump_{strt_day}to{yesterday}" \
          f" where funding_txn_type in ('Direct Assignment', 'IDFC DA- Rs 10.51 Cr - Jan 2020')" \
          f"and funding_txn_remark<>'Dvara KGFS DA Feb 2020 - Karambayam' "
raw_data = pd.read_sql(query_1, cnx, index_col=None)
cnx.dispose()
print('raw_data_created')

branch_master = pd.read_excel('branch_master.xlsx')

#branch Corrections

raw_data['funding_txn_remark'].replace('Dvara KGFS Habitat DA Feb 2022 Ã¢â‚¬â€œ Catalyst Trusteeship Limited', 'Dvara KGFS Habitat DA Feb 2022 â€“ Catalyst Trusteeship Limited', inplace=True)
raw_data['branch'] = raw_data['branch'].str.upper()
print(raw_data.shape)
final_raw_data = pd.merge(raw_data, branch_master, on='branch', how='left')
print(final_raw_data.shape)

final_raw_data = final_raw_data.astype({"DPD_Days": int})

print('data correction completed')
# DPD Cat
def cat(x):
    if x['DPD_Days'] == 0:
        return 'schd'
    elif x['DPD_Days'] > 0 and x['DPD_Days'] < 31:
        return '1to30'
    elif x['DPD_Days'] > 30 and x['DPD_Days'] < 61:
        return '31to60'
    elif x['DPD_Days'] > 60 and x['DPD_Days'] < 91:
        return '61to90'
    elif x['DPD_Days'] > 90 and x['DPD_Days'] < 181:
        return '91to180'
    elif x['DPD_Days'] > 180 and x['DPD_Days'] < 366:
        return '181to365'
    return '>365'

final_raw_data['DPD_cat'] = final_raw_data.apply(lambda x: cat(x), axis=1)



#arrange dataframe
final_raw_data = final_raw_data[['KGFS','branch','URN','AccountNumber','funder_name','funding_txn_type','funding_txn_remark','Product',
                                 'Account_Status', 'POS','DPD_Days','DPD_cat','customer_name','state','DisbursementDate', 'Last_Repayment_Date']]
print('export to excel started.....')
#Write excel
for i in (final_raw_data['funding_txn_remark'].unique()):
    data = final_raw_data[final_raw_data['funding_txn_remark'] == i].copy()
    pivot = data.pivot_table(index=['KGFS', 'state'],columns='DPD_cat', values='POS', aggfunc='sum', fill_value=0, margins=True, margins_name='Total')
    #rename column
    data.rename(columns={'POS': f'Prin_OS_as_on_{yesterday}', 'DPD_Days': f'Overdue_Days_as_on_{yesterday}'},
                inplace=True)
    writer = pd.ExcelWriter(f'{i}.xlsx', engine='xlsxwriter')
    pivot.to_excel(writer, sheet_name='summary')
    data.to_excel(writer, sheet_name='raw_data', index=False)

    writer.sheets['raw_data'].set_column(0, 16, 17)
    writer.sheets['summary'].set_column(0, 10, 12)
    writer.save()

    print(f'{i}.created')

print('export to excel completed.....')
#e-mail
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'vijayakumar.g1@dvarakgfs.com; Shilpa.B@Dvarakgfs.com; vinodkumar.m@dvarakgfs.com; Dilli.Kumar@dvarakgfs.com; devina.r@dvarakgfs.com'
mail.CC = 'prakash.r@dvarakgfs.com; krishnakumar.a@dvarakgfs.com'
#----
mail.Subject = f'Direct Assignment Transactions as on {yesterday}'
mail.HTMLBody = f"""Dear All<br><br> PFA the DA Transaction data as on {yesterday}.<br><br> Regards<br>Analytics Team
"""
cwd = os.getcwd()
for i in (final_raw_data['funding_txn_remark'].unique()):
    print(f'{i}.xlsx')
    mail.Attachments.Add(f'{cwd}/{i}.xlsx')
mail.Send()
print('send email.....')

