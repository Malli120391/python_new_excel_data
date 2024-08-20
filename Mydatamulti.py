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

# Path to your Excel workbook
excel_path = r'C:\Users\malle\Downloads\SampleSuperstoreons.xlsx'

# Load Excel workbook
excel_file = pd.ExcelFile(excel_path)

# Print available sheet names for debugging
print("Available sheets:", excel_file.sheet_names)

# Iterate over each sheet in the workbook
for sheet_name in excel_file.sheet_names:
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except ValueError as e:
        print(f"Error reading sheet {sheet_name}: {e}")
        continue

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

    # Delete data from worksheet
    df = pd.DataFrame(columns=df.columns)  # Clear the DataFrame
    df.to_excel(excel_path, sheet_name=sheet_name, index=False)

# Export data to a new Excel file
export_path = r'C:\Users\malle\Downloads\path_to_exported_excel_file.xlsx'
with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
    for sheet_name in excel_file.sheet_names:
        query = f"SELECT * FROM [{sheet_name}]"
        df = pd.read_sql(query, conn)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Drop tables if not needed
for sheet_name in excel_file.sheet_names:
    drop_table_query = f"DROP TABLE IF EXISTS [{sheet_name}];"
    cursor.execute(drop_table_query)
    conn.commit()

# Clean up
cursor.close()
conn.close()
print("Operation completed successfully.")
