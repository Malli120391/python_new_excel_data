import pandas as pd
import pyodbc

# import openpyxl

# Step 1: Load data from the Excel file
excel_file = r'C:\Users\malle\Downloads\SampleSuperstoreone.xlsx'
sheet_name = 'Orders'  # Adjust this to your specific sheet name
df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')

# Step 2: Define the SQL Server connection using Windows Authentication
server = 'DESKTOP-60AKLBN\SQLEXPRESS'  # Replace with your server name
database = 'OWN_DB'  # Replace with your database name

conn_str = (
    # 'DRIVER={SQL Server Native Client 11.0};'
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f'SERVER={server};DATABASE={database};'
    'Trusted_Connection=yes;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Step 3: Create a table in SQL Server
table_name = 'Orders'  # Replace with your desired table name


# Determine column types
def infer_sql_type(dtype):
    if pd.api.types.is_integer_dtype(dtype):
        return 'INT'
    elif pd.api.types.is_float_dtype(dtype):
        return 'FLOAT'
    elif pd.api.types.is_datetime64_any_dtype(dtype):
        return 'DATETIME'
    else:
        return 'NVARCHAR(MAX)'


# Generate a SQL CREATE TABLE statement based on the DataFrame
# columns = ", ".join([f"{col} NVARCHAR(MAX)" for col in df.columns])

columns = ", ".join([f"{col} {infer_sql_type(df[col].dtype)}" for col in df.columns])
create_table_sql = f"CREATE TABLE {table_name} ({columns})"
print(create_table_sql)
cursor.execute(create_table_sql)
conn.commit()

# Step 4: Insert data into SQL Server table
for index, row in df.iterrows():
    placeholders = ", ".join("?" * len(row))
    insert_sql = f"INSERT INTO {table_name} VALUES ({placeholders})"
    cursor.execute(insert_sql, tuple(row))
conn.commit()

# Step 5: Clear the data in the Excel file
df.iloc[0:0].to_excel(excel_file, sheet_name=sheet_name, index=False)

# Step 6: Add new data into the same Excel file manually
# Manually add data to the Excel file here or modify the file as needed

# Step 7: Read new data from the Excel file
new_df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Step 8: Update the table with the new data
for index, row in new_df.iterrows():
    placeholders = ", ".join("?" * len(row))
    insert_sql = f"INSERT INTO {table_name} VALUES ({placeholders})"
    cursor.execute(insert_sql, tuple(row))
conn.commit()

# Step 9: Clear the data in the Excel file again
new_df.iloc[0:0].to_excel(excel_file, sheet_name=sheet_name, index=False)

# Step 10: Export the data from SQL Server into a different Excel file
output_excel_file = r'C:\Users\malle\Downloads\output_excel_file.xlsx'
df_from_sql = pd.read_sql(f"SELECT * FROM {table_name}", conn)
df_from_sql.to_excel(output_excel_file, index=False)

# Clean up
cursor.close()
conn.close()

print("Process completed successfully.")
