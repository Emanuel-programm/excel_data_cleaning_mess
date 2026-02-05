import pandas as pd
import matplotlib.pyplot as plt
import os

# define the cleaning Engine:
def clean_data(df):
    """Handles all cleaning data and standardization"""
    df=df.copy()  ## Working on a copy to avoid modifying original data
    df.fillna(0,inplace=True)
    df.drop_duplicates(inplace=True)

    ## standardize Strings
    if 'Customer Name' in df.columns:
        df['Customer Name']=df['Customer Name'].str.strip().str.title()

    ## clean dateformat colums
    date_cols = ['Order Date', 'Ship Date']
    for col in date_cols:
        if col in df.columns:
            df[col]=pd.to_datetime(df[col],errors='coerce').dt.strftime('%y-%m-%d')
    

    ## Add Remarks business logic
    if 'Tax (USD)' in df.columns:
        df['Remark']=df['Tax (USD)'].apply(lambda x: 'Hight Tax' if x>20 else 'Low Tax')

    
    return df

## Define the reporting engine
def generate_daily_reports(df):
    """Create a business specific summary"""
    #Daily sales report
    daily=df.groupby('Order Date').agg({'Total (USD)':'sum','Order No':'count'}).reset_index()

    daily.rename(columns={'Total (USD)':'Daily_Total_USD','Order No':'Number_of_orders'},inplace=True)
    return daily

## Define the visualizarion Engine
def create_visuals(report_df):
    """Generate the charts for the clients"""
    plt.figure(figsize=(5,5))
    plt.bar(report_df['Order Date'],report_df['Daily_Total_USD'],color='skyblue',edgecolor='black')
    plt.title("Daily USDS final report")
    plt.xlabel("order date")
    plt.xticks(rotation=45)
    plt.ylabel("Total USD")
    plt.tight_layout()
    plt.savefig("sales_chart.png")
    print("Chart saved as sales_chart")

    plt.show()


##### the main automation block #############
def main():
    path=("G:/excel_files/supermarket.xlsx")
    if not os.path.exists(path):
        print("No file found in that direcyory or in your system")
        return
    try:
        # load data
        raw_data=pd.read_excel(path,usecols='B:I',skiprows=7)

        # process data-cleaning
        print("Cleaning data...")
        cleaned_data=clean_data(raw_data)

        ## Generate the reports for each day
        print("Generating daily report...")
        daily_report=generate_daily_reports(cleaned_data)

        ## Generate the charts for the clients
        print("Creating visuals..")
        create_visuals(daily_report)

        ## save cleaned file (Export professional mult-sheet excel report)
        with pd.ExcelWriter("professional_report.xlsx") as writer:
            cleaned_data.to_excel(writer,sheet_name='cleaned_data',index=False)
            daily_report.to_excel(writer,sheet_name='daily_report',index=False)
        print("Done! Professional report saved as professional_report.xlsx")

    except Exception as e:
        print(f"An error occurred: {e}")

# This line ensures the code only runs if you play the file directly         
if __name__ =="__main__":
    main()


    