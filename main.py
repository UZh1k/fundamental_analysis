from time import sleep

import FundamentalAnalysis as fa
import pandas as pd

api_key = '891eb7484d76ff0dd412e8df2cdc3d74'

tickers = ['MSFT', 'AAPL', 'SNP', 'UBI', 'ATVI']
df = pd.DataFrame()

for ticker in tickers:
    for _ in range(10):
        try:
            one_companie_df = fa.balance_sheet_statement(ticker, api_key, period="annual", limit=5).T
            break
        except ValueError:
            sleep(1)
            continue

    one_companie_df['ticker'] = ticker
    df = df.append(one_companie_df)

cols = df.columns.tolist()[:-1]
cols.insert(0, 'ticker')
cols_to_remove = ['cik', 'fillingDate', 'acceptedDate', 'period', 'link', 'finalLink']
for ctr in cols_to_remove:
    cols.remove(ctr)

df = df[cols]

writer = pd.ExcelWriter("fundamental_analysis_result.xlsx", engine = 'openpyxl')

df.to_excel(writer,
            sheet_name='balance_sheet')

df_cash_flow = pd.DataFrame()
for ticker in tickers:
    for _ in range(10):
        try:
            one_companie_df = fa.cash_flow_statement(ticker, api_key, period="annual", limit=5).T
            break
        except ValueError:
            sleep(1)
            continue

    one_companie_df['ticker'] = ticker
    df_cash_flow = df_cash_flow.append(one_companie_df)

cols = df_cash_flow.columns.tolist()[:-1]
cols.insert(0, 'ticker')
for ctr in cols_to_remove:
    cols.remove(ctr)

df_cash_flow = df_cash_flow[cols]
df_cash_flow.to_excel(writer,
                      sheet_name='cash_flow')

df_income_statement = pd.DataFrame()
for ticker in tickers:
    for _ in range(10):
        try:
            one_companie_df = fa.income_statement(ticker, api_key, period="annual", limit=5).T
            break
        except ValueError:
            sleep(1)
            continue
    one_companie_df['ticker'] = ticker
    df_income_statement = df_income_statement.append(one_companie_df)

cols = df_income_statement.columns.tolist()[:-1]
cols.insert(0, 'ticker')
for ctr in cols_to_remove:
    cols.remove(ctr)

df_income_statement = df_income_statement[cols]
df_income_statement.to_excel(writer,
                             sheet_name='income_statement')

writer.save()
