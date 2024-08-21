"""my data in excel file i want to create table using same excel file data and read data
then after delete data in excel file after
i add some data into same excel file again update that data into same table if same excel file new date
is available automate program every 30 minutes  for update and delete
same excel file data into excel file and export data into different excel file
using python code with sql server windows authentication so i need proper full code"""

import pandas as pd
#import pyodbc
from sqlalchemy import create_engine
import time
import os

# SQL Server connection details
server = 'DESKTOP-60AKLBN\SQLEXPRESS'
database = 'OWN_DB'
# Windows Authentication (username and password can be omitted)
connection_string = f'mssql+pyodbc://@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server'

# Create SQL Alchemy engine
engine = create_engine(connection_string)

# File paths
input_file = r'C:\Users\malle\Downloads\SampleSuperstorekm.xlsx'
export_file = r'C:\Users\malle\Downloads\output_excel_filethree.xlsx'


def process_excel_file():
    # Read data from Excel file
    df = pd.read_excel(input_file, sheet_name=None)

    for sheet_name, sheet_data in df.items():
        # Create or update table in SQL Server
        sheet_data.to_sql(sheet_name, con=engine, if_exists='replace', index=False)

    # Clear the data in the Excel file by writing empty DataFrames
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet_name in df.keys():
            empty_df = pd.DataFrame()
            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Export data from SQL Server to a different Excel file
    with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
        for sheet_name in df.keys():
            query = f'SELECT * FROM {sheet_name}'
            df_export = pd.read_sql(query, con=engine)
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)

    print("Processed Excel file and updated SQL Server table.")


def main():
    while True:
        if os.path.exists(input_file):
            process_excel_file()
        else:
            print(f"Input file {input_file} not found.")

        time.sleep(1800)  # Sleep for 30 minutes (1800 seconds)


if __name__ == "__main__":
    main()






