"""my data in excel file i want to create table using same excel file data and read data
then after delete data in excel file after
i add some data into same excel file again update that data into same table and delete
same excel file data into excel file and export data into different excel file
using python code with sql server windows authentication so i need full code"""

"""1.Create a table using data from an Excel file.
2.Read the data and insert it into the SQL Server table.
3.Delete the data in the Excel file.
4.Allow you to add new data into the same Excel file.
5.Update the SQL Server table with the new data.
6.Delete the data in the Excel file again.
7.Export the data into a different Excel file."""

import pandas as pd
import pyodbc

# Step 1: Load data from the Excel file
excel_file = r'C:\Users\malle\Downloads\SampleSuperstoreOrders.xlsx'  # Use .xlsx file
sheet_name = 'Orders'  # Adjust this to your specific sheet name
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')

# Define the SQL Server connection using Windows Authentication
server = 'DESKTOP-60AKLBN\SQLEXPRESS'  # Replace with your server name
database = 'OWN_DB'  # Replace with your database name

conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f'SERVER={server};DATABASE={database};'
    'Trusted_Connection=yes;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Step 2: Create a table in SQL Server
table_name = 'Orders'  # Replace with your desired table name

# Drop table if no need
"""drop_table_sql = f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE [{table_name}]"
try:
    cursor.execute(drop_table_sql)
    conn.commit()
    print(f"Table '{table_name}' dropped successfully.")
except pyodbc.ProgrammingError as e:
    print(f"ProgrammingError: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    cursor.close()
    conn.close()"""

# Define data type mapping
def infer_sql_type(dtype):
    if pd.api.types.is_integer_dtype(dtype):
        return 'INT'
    elif pd.api.types.is_float_dtype(dtype):
        return 'FLOAT'
    elif pd.api.types.is_datetime64_any_dtype(dtype):
        return 'DATETIME'
    elif pd.api.types.is_bool_dtype(dtype):
        return 'BIT'
    else:
        return 'NVARCHAR(MAX)'

# Generate a SQL CREATE TABLE statement based on the DataFrame
columns = ", ".join([f"[{col}] {infer_sql_type(df[col].dtype)}" for col in df.columns])
create_table_sql = f"CREATE TABLE [{table_name}] ({columns})"
print(create_table_sql)  # Print SQL for debugging

try:
    cursor.execute(create_table_sql)
    conn.commit()
    print("Table created successfully!")
except pyodbc.ProgrammingError as e:
    print(f"ProgrammingError: {e}")
except Exception as e:
    print(f"An error occurred: {e}")

# Step 3: Insert data into SQL Server table
for index, row in df.iterrows():
    placeholders = ", ".join("?" * len(row))
    insert_sql = f"INSERT INTO [{table_name}] VALUES ({placeholders})"
    cursor.execute(insert_sql, tuple(row))
conn.commit()

# Step 4: Clear the data in the Excel file
df.iloc[0:0].to_excel(excel_file, sheet_name=sheet_name, index=False)

# Step 5: Add new data into the same Excel file manually
# Manually add data to the Excel file here or modify the file as needed

# Step 6: Read new data from the Excel file
new_df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')

# Step 7: Update the table with the new data
for index, row in new_df.iterrows():
    placeholders = ", ".join("?" * len(row))
    insert_sql = f"INSERT INTO [{table_name}] VALUES ({placeholders})"
    cursor.execute(insert_sql, tuple(row))
conn.commit()

# Step 8: Clear the data in the Excel file again
new_df.iloc[0:0].to_excel(excel_file, sheet_name=sheet_name, index=False)

# Step 9: Export the data from SQL Server into a different Excel file
output_excel_file = r'C:\\Users\\malle\\Downloads\\output_excel_fileget.xlsx'
df_from_sql = pd.read_sql(f"SELECT * FROM [{table_name}]", conn)
df_from_sql.to_excel(output_excel_file, index=False)

# Clean up
cursor.close()
conn.close()
print("Process completed successfully.")
