import requests
import pandas as pd
from sqlalchemy import create_engine
from bs4 import BeautifulSoup
import os
import urllib.parse


# Define the URL where the XLSX files are located
url = "https://www.nrb.org.np/category/monthly-statistics/?department=bfr"

# Define the directory to save downloaded files
download_dir = "downloaded_xlsx_files"

# Create the download directory if it doesn't exist
os.makedirs(download_dir, exist_ok=True)

# Database connection settings (use your connection string)
server = ''
database = ''
username = ''
password = ''
dsn_name = ''
conn_str = f'mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(f"DSN={dsn_name};UID={username};PWD={password};DATABASE={database}")}'
#engine = create_engine(conn_str)
try:
    # Create an SQLAlchemy engine for SQL Server connection
    engine = create_engine(conn_str)

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content of the page
        soup = BeautifulSoup(response.text, 'html.parser')

        # Find all the links to XLSX files on the page
        links = soup.find_all('a', href=True)

        # Iterate through the links and download XLSX files
        for link in links:
            file_url = link.get('href')
            if file_url.endswith('.xlsx'):
                # Get the filename from the URL
                filename = os.path.join(download_dir, os.path.basename(file_url))

                # Download the XLSX file
                with open(filename, 'wb') as file:
                    file_response = requests.get(file_url)
                    file.write(file_response.content)

                # Load "C10" sheets from the XLSX file into DataFrames
                xls = pd.ExcelFile(os.path.join(download_dir, os.path.basename(filename)))
                if "C10" in xls.sheet_names:
                    df_c10 = xls.parse("C10", header=None)  # Set header=None to read without header

                    # Drop the first two rows (header and units rows)
                    df_c10 = df_c10.iloc[2:, :]
                    df_c10 = df_c10.transpose()
                    # Remove columns where all row data is empty or null
                    
                    df_c10.dropna(axis=1, how='all', inplace=True)

                    
                    df_c10 = df_c10.iloc[1:, :]
                    #df_c10.columns = df_c10.iloc[0, :]
                  
                    df_c10.reset_index(drop=True, inplace=True)
                    df_c10.columns = df_c10.iloc[0]

                    # Drop the first row since it's now the header
                    df_c10 = df_c10[1:]

                    # Reset the index 
                    df_c10.reset_index(drop=True, inplace=True)
    
                    # Resetting index to the default RangeIndex
                    df_c10.reset_index(drop=True, inplace=True)  # Resetting to a default RangeIndex

                    # Setting a specific column as the index
              
                   # Add a new column 'source_link' with the URL
                    df_c10['source_link'] = file_url
                    df_c10['extracted_source_link'] = df_c10['source_link'].str.extract(r'([^/]+)_Publish')
                    
                    # Insert data into the SQL Server database for "C10" (adjust table name accordingly)
                    df_c10.to_sql('LOANS_ADVANCES', engine, if_exists='replace', index=False)

               
    else:
        print("Failed to fetch the webpage.")

except Exception as e:
    print(f"An error occurred: {str(e)}")
