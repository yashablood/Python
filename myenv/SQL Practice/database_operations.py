import os
import sqlite3

def execute_sql_script(sql_file):
    conn = sqlite3.connect('Books_database.db')  
    cursor = conn.cursor()

    if not os.path.exists(sql_file):
        raise FileNotFoundError(f"SQL file '{sql_file}' not found.")

    # Read and execute the script
    with open(sql_file, 'r') as f:
        sql_script = f.read()
        cursor.executescript(sql_script)
        conn.commit()

    conn.close()

def main():
    # Construct absolute path to database_setup.sql
    script_dir = os.path.dirname(__file__)
    sql_file = os.path.join(script_dir, 'database_setup.sql')

    # Execute the SQL script
    execute_sql_script(sql_file)
    print("SQL script executed successfully.")

if __name__ == "__main__":
    main()
