import pandas as pd
import pyodbc

# SQL Server connection setup
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=DESKTOP-60AKLBN\SQLEXPRESS;'
    r'DATABASE=OWN_DB;'
    r'TRUSTED_CONNECTION=yes;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Function to check if a table exists
def table_exists(table_name):
    query = f"""
    SELECT * FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = '{table_name}'
    """
    return cursor.execute(query).fetchone() is not None


# Path to your Excel workbook
excel_path = r'C:\Users\malle\Downloads\SampleSuperstorelast.xlsx'

# Load Excel workbook
excel_file = pd.ExcelFile(excel_path)

# Print available sheet names for debugging
print("Available sheets:", excel_file.sheet_names)

# Iterate over each sheet in the workbook
for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Create table in SQL Server
    columns_with_types = []
    for col in df.columns:
        # Infer SQL data type for each column
        if pd.api.types.is_integer_dtype(df[col]):
            sql_type = 'INT'
        elif pd.api.types.is_float_dtype(df[col]):
            sql_type = 'FLOAT'
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            sql_type = 'DATETIME'
        else:
            sql_type = 'NVARCHAR(MAX)'

        # Use square brackets around column names
        columns_with_types.append(f"[{col}] {sql_type}")

    create_table_query = f"""
    IF OBJECT_ID('[{sheet_name}]', 'U') IS NOT NULL
        DROP TABLE [{sheet_name}];

    CREATE TABLE [{sheet_name}] (
        {', '.join(columns_with_types)}
    );
    """
    print(f"Creating table with query:\n{create_table_query}")  # Debug print
    try:
        cursor.execute(create_table_query)
        conn.commit()
    except pyodbc.Error as e:
        print(f"Error creating table: {e}")

    # Insert data into the table
    insert_query = f"INSERT INTO [{sheet_name}] ({', '.join([f'[{col}]' for col in df.columns])}) VALUES ({', '.join(['?' for _ in df.columns])})"
    for index, row in df.iterrows():
        print(f"Insert query: {insert_query}")  # Debug print
        print(f"Row data: {tuple(row)}")  # Debug print
        try:
            cursor.execute(insert_query, tuple(row))
        except pyodbc.Error as e:
            print(f"Error inserting data: {e}")
            print(f"Data causing error: {tuple(row)}")  # Print problematic data
    conn.commit()

# Verify table existence and export data
for sheet_name in excel_file.sheet_names:
    if table_exists(sheet_name):
        query = f"SELECT * FROM [{sheet_name}]"
        try:
            df = pd.read_sql(query, conn)
            export_path = r'C:\Users\malle\Downloads\path_to_exported_excel_filetwo.xlsx'
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Exported data from table '{sheet_name}' to Excel.")
        except pd.io.sql.DatabaseError as e:
            print(f"Error querying table {sheet_name}: {e}")
    else:
        print(f"Table '{sheet_name}' does not exist in the database.")

# Drop tables if not needed
"""for sheet_name in excel_file.sheet_names:
    if table_exists(sheet_name):
        drop_table_query = f"DROP TABLE IF EXISTS [{sheet_name}];"
        cursor.execute(drop_table_query)
        conn.commit()
        print(f"Dropped table '{sheet_name}'.")"""

# Clean up
cursor.close()
conn.close()
print("Operation completed successfully.")
