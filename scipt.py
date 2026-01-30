import pandas as pd
import matplotlib.pyplot as plt
import os

path=("G:/excel_files/supermarket.xlsx")
if os.path.exists(path):
    data=pd.read_excel(path,usecols='B:I',skiprows=7)
    data.head()

    # clean dateformat colums
    data['Order Date']=pd.to_datetime(data['Order Date'],errors='coerce').dt.strftime('%y-%m-%d')
    data['Ship Date']=pd.to_datetime(data['Ship Date'],errors='coerce').dt.strftime('%y-%m-%d')

    data.head(4)
    data.columns

    # checking for the total USD > 150
    total_usd=data[data['Total (USD)']>150]
    total_usd.to_excel('Total_usd.xlsx',index=False)
    print(total_usd)

    ## fill unfilled column with zeros
    data.fillna(0,inplace=True)
    ## drop the duplicates data
    data.drop_duplicates(inplace=True)

    ## Generate the report for each day
    daily_report=data.groupby('Order Date').agg({'Total (USD)':'sum','Order No':'count'}).reset_index()

    print(daily_report)
    ## Rename column to the daily report
    daily_report.rename(columns={'Total (USD)':'Daily_Total_USD','Order No':'Number_of_orders'},inplace=True)
    daily_report.to_excel('Daily_report.xlsx',index=False)

    ## data visualization
    daily_report.plot(x='Order Date',y='Daily_Total_USD', kind='bar',figsize=(12,6),color='skyblue',edgecolor='black')
    plt.title("Daily UDS sales report")
    plt.xlabel("Order date")
    plt.ylabel("Total Usd")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    ## Generate the column remark
    data['Remark']=data['Tax (USD)'].apply(lambda x: 'Hight Tax' if x>20 else 'Low Tax')

    ### Rename some colums
    data.rename(columns={'Total USD()':'Total_USD ','Tax USD()':'Tax_USD '},inplace=True) 

    ## standardize column names
    data['Customer Name']=data['Customer Name'].str.strip().str.title()

    ### save cleaned file
    with pd.ExcelWriter("cleaned_report.xlsx") as writer:
        data.to_excel(writer,sheet_name='report',index=False)