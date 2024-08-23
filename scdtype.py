"""my data in excel file with 1 to 12 months data so  i want to create table using
same excel file data and read data
then after delete data in excel file after
i add some data into same excel file again update that data into same table if same excel file have new data
is available automate program every 30 minutes  insert or update and delete
same excel file data into excel file and export data into different excel file than
i want to i want maintain only 2 to 3 months data like using scd types
using python code with sql server windows authentication so i need proper full code"""

import pandas as pd
import pyodbc
import os
from datetime import datetime, timedelta
import time

# Database connection parameters
server = 'DESKTOP-60AKLBN\SQLEXPRESS'
database = 'OWN_DB'
conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'

# File paths
input_file = r'C:\Users\malle\Downloads\SampleStore.xlsx'
export_file = r'C:\Users\malle\Downloads\dataSampledate.xlsx'

# Connect to SQL Server
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()


# Create a table based on the Excel file data
def create_table_from_excel():
    df = pd.read_excel(input_file)

    # Check if the table exists
    table_check_query = "SELECT COUNT(*) FROM sys.tables WHERE name = 'Orders'"
    cursor.execute(table_check_query)
    table_exists = cursor.fetchone()[0]

    # If the table doesn't exist, create it
    if table_exists == 0:
        columns = ', '.join(
            [f"[{col.replace(' ', '_')}]" + " NVARCHAR(MAX)" for col in df.columns])  # Replace spaces in column names
        create_table_query = f"CREATE TABLE Orders ({columns})"
        cursor.execute(create_table_query)
        conn.commit()
        print("Table created:", create_table_query)

    # Insert data into the SQL table
    for index, row in df.iterrows():
        columns = ', '.join([f"[{col.replace(' ', '_')}]" for col in df.columns])
        placeholders = ', '.join(['?' for _ in row])
        insert_query = f"INSERT INTO Orders ({columns}) VALUES ({placeholders})"
        cursor.execute(insert_query, tuple(row))
    conn.commit()

    # Clear the Excel file
    writer = pd.ExcelWriter(input_file, engine='openpyxl')
    pd.DataFrame().to_excel(writer, index=False)
    writer.close()


# Update the table if new data is present in the Excel file
def update_table_from_excel():
    df = pd.read_excel(input_file)
    for index, row in df.iterrows():
        update_query = f"UPDATE Orders SET Column1 = ?, Column2 = ? WHERE id = ?"
        cursor.execute(update_query, row['Column1'], row['Column2'], row['id'])
    conn.commit()


# Maintain only the last 2-3 months of data
def maintain_recent_data():
    actual_date_column = '[Order_Date]'  # Update this line with the correct column name
    cutoff_date = datetime.now() - timedelta(days=90)  # Adjust for 2-3 months
    delete_query = f"DELETE FROM Orders WHERE {actual_date_column} < ?"
    cursor.execute(delete_query, cutoff_date)
    conn.commit()


# Export data to a different Excel file
def export_data_to_excel():
    query = "SELECT * FROM Orders"
    df = pd.read_sql(query, conn)
    df.to_excel(export_file, index=False)


# Main loop to automate every 30 minutes
while True:
    create_table_from_excel()
    update_table_from_excel()
    maintain_recent_data()
    export_data_to_excel()

    # Wait for 30 minutes
    time.sleep(1800)











