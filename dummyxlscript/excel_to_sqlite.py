import pandas as pd
import sqlite3
import os

def excel_to_sqlite(excel_dir, db_name="my_database.db"):
    """
    Converts all sheets from all XLSX files in a given directory into
    tables in an SQLite database.

    Args:
        excel_dir (str): The path to the directory containing the XLSX files.
        db_name (str): The name of the SQLite database file to create/connect to.
                       Defaults to 'my_database.db'.
    """
    # Connect to SQLite database (creates it if it doesn't exist)
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    print(f"Connected to SQLite database: {db_name}")

    # Iterate through all files in the specified directory
    for filename in os.listdir(excel_dir):
        if filename.endswith(".xlsx"):
            excel_filepath = os.path.join(excel_dir, filename)
            print(f"\nProcessing Excel file: {filename}")

            try:
                # Load all sheets from the Excel file
                xls = pd.ExcelFile(excel_filepath)

                # Iterate through each sheet in the Excel file
                for sheet_name in xls.sheet_names:
                    # Read the sheet into a pandas DataFrame
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    # Clean up column names: replace spaces with underscores, remove special characters
                    # and make them lowercase to be valid SQL column names
                    df.columns = [
                        "".join(c if c.isalnum() else "_" for c in str(col)).lower()
                        for col in df.columns
                    ]

                    # Generate a valid table name from the sheet name
                    # Remove spaces, special characters, and make lowercase
                    table_name = "".join(
                        c if c.isalnum() else "_" for c in sheet_name
                    ).lower()
                    
                    # Ensure table name starts with a letter or underscore if it doesn't already
                    if not table_name[0].isalpha() and table_name[0] != '_':
                        table_name = '_' + table_name

                    print(f"  Processing sheet: '{sheet_name}' -> Table: '{table_name}'")

                    # Drop the table if it already exists to ensure a clean import
                    cursor.execute(f"DROP TABLE IF EXISTS {table_name}")

                    # Write the DataFrame to the SQLite table
                    # if_exists='replace' will drop and recreate the table
                    # index=False prevents pandas from writing the DataFrame index as a column
                    df.to_sql(table_name, conn, if_exists="replace", index=False)
                    print(f"    Successfully imported '{sheet_name}' into table '{table_name}'.")

            except Exception as e:
                print(f"  Error processing {filename}: {e}")

    # Commit changes and close the connection
    conn.commit()
    conn.close()
    print(f"\nDatabase '{db_name}' created/updated successfully with data from XLSX files.")

# --- How to use this script ---
if __name__ == "__main__":
    # IMPORTANT: Replace 'your_excel_files_directory' with the actual path
    # where your .xlsx files are located.
    # Example: excel_directory = "/Users/yourusername/Documents/MyExcelData"
    # Example: excel_directory = "C:\\Users\\yourusername\\Documents\\MyExcelData"
    # For demonstration, let's assume your XLSX files are in a folder named 'excel_data'
    # in the same directory as this script.
    
    # Create a dummy directory and files for demonstration if they don't exist
    demo_dir = "excel_data"
    if not os.path.exists(demo_dir):
        os.makedirs(demo_dir)
        print(f"Created demo directory: {demo_dir}")

        # Create a dummy Excel file 1
        df1 = pd.DataFrame({
            'Product ID': [1, 2, 3],
            'Product Name': ['Laptop', 'Mouse', 'Keyboard'],
            'Price (USD)': [1200, 25, 75]
        })
        df2 = pd.DataFrame({
            'Customer ID': [101, 102],
            'Customer Name': ['Alice', 'Bob'],
            'Email': ['alice@example.com', 'bob@example.com']
        })
        with pd.ExcelWriter(os.path.join(demo_dir, 'products_customers.xlsx')) as writer:
            df1.to_excel(writer, sheet_name='Products', index=False)
            df2.to_excel(writer, sheet_name='Customers Data', index=False)
        print(f"Created dummy file: {os.path.join(demo_dir, 'products_customers.xlsx')}")

        # Create a dummy Excel file 2
        df3 = pd.DataFrame({
            'Order_ID': [1001, 1002],
            'Item': ['Laptop', 'Mouse'],
            'Quantity': [1, 2]
        })
        with pd.ExcelWriter(os.path.join(demo_dir, 'orders.xlsx')) as writer:
            df3.to_excel(writer, sheet_name='Orders List', index=False)
        print(f"Created dummy file: {os.path.join(demo_dir, 'orders.xlsx')}")

    excel_directory = '/Users/pratikshamehta/Desktop/dummyxlscript' # Use the demo directory for execution

    # Name of your SQLite database file
    sqlite_database_name = "security_data.db"

    # Call the function to perform the conversion
    excel_to_sqlite(excel_directory, sqlite_database_name)

    # --- Verification (Optional) ---
    print("\n--- Verifying database content ---")
    conn_verify = sqlite3.connect(sqlite_database_name)
    cursor_verify = conn_verify.cursor()

    # Get list of tables
    cursor_verify.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor_verify.fetchall()
    print(f"Tables in '{sqlite_database_name}': {tables}")

    # Read data from a sample table (e.g., 'products')
    try:
        print("\nContent of 'products' table:")
        sample_df = pd.read_sql_query("SELECT * FROM products", conn_verify)
        print(sample_df)
    except Exception as e:
        print(f"Could not read 'products' table (it might not exist or be named differently): {e}")

    conn_verify.close()
    print("\nVerification complete.")
