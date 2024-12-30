import requests
import pandas as pd
from datetime import datetime, timedelta
import time

# API endpoint and headers
url = "https://api.finanvo.in/company/newcompanies/list"
headers = {
    "Content-Type": "application/json",
    "x-api-key": "OliINWwpCj",
    "x-api-secret-key": "smdlme2Y9n25v8htM0uQkWGqxGdN3D8O8aLkf9Zl"
}

# Step 1: Fetch data from API for a specific date
def fetch_data_from_api(url, headers, date):
    try:
        params = {'type': 'date', 'bulkMasterType': 'CIN', 'date': date}
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()  # Raise an error for HTTP issues
        return response.json()  # Return the parsed JSON response
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for {date}: {e}")
        return {}

# Step 2: Process and convert data to DataFrame
def process_data_to_dataframe(data):
    if 'cin' in data:  # Ensure the 'cin' key exists in the response
        companies = data['cin']
        company_list = []

        for company in companies:
            # Extract company details
            company_info = {
                "CIN": company.get("CIN", ""),
                "COMPANY_NAME": company.get("COMPANY_NAME", ""),
                "ACTIVITY_CODE": company.get("ACTIVITY_CODE", ""),
                "DATE_OF_REGISTRATION": company.get("DATE_OF_REGISTRATION", ""),
                "STATE": company.get("STATE", ""),
                "ROC": company.get("ROC", ""),
                "COMPANY_STATUS": company.get("COMPANY_STATUS", ""),
                "CATEGORY": company.get("CATEGORY", ""),
                "CLASS": company.get("CLASS", ""),
                "SUBCATEGORY": company.get("SUBCATEGORY", ""),
                "AUTHORIZED_CAPITAL": company.get("AUTHORIZED_CAPITAL", ""),
                "PAIDUP_CAPITAL": company.get("PAIDUP_CAPITAL", ""),
                "ACTIVITY_DESCRIPTION": company.get("ACTIVITY_DESCRIPTION", ""),
                "REGISTERED_OFFICE_ADDRESS": company.get("REGISTERED_OFFICE_ADDRESS", ""),
                "ADDRESS_OTHER_THAN_RO": company.get("ADDRESS_OTHER_THAN_RO", ""),
                "EMAIL": company.get("EMAIL", ""),
                "LAST_AGM": company.get("LAST_AGM", ""),
                "LAST_BALANCESHEET": company.get("LAST_BALANCESHEET", ""),
                "LISTING_STATUS": company.get("LISTING_STATUS", ""),
                "ACTIVE_COMPLIANCE": company.get("ACTIVE_COMPLIANCE", ""),
                "COMPANY_FILING_STATUS_16_17_18": company.get("COMPANY_FILING_STATUS_16_17_18", ""),
                "SUSPENDED_AT_STOCK": company.get("SUSPENDED_AT_STOCK", ""),
                "NUMBER_OF_MEMBERS": company.get("NUMBER_OF_MEMBERS", ""),
                "PARTNERS": company.get("PARTNERS", ""),
                "CIRP": company.get("CIRP", ""),
                "TIMESTAMP": company.get("TIMESTAMP", ""),
                "PINCODE": company.get("PINCODE", ""),
                "COUNTRY": company.get("COUNTRY", ""),
                "CITY": company.get("CITY", ""),
                "DISTRICT": company.get("DISTRICT", ""),
                "TOTAL_OBLIGATION_CONTRIBUTION": company.get("TOTAL_OBLIGATION_CONTRIBUTION", ""),
                "TYPE_OF_COMPANY": company.get("TYPE_OF_COMPANY", ""),
                # Capture the directors information
                "DIRECTORS": company.get("DIRECTORS", [])
            }
            company_list.append(company_info)

        df = pd.DataFrame(company_list)
        return df
    else:
        print("No 'cin' data found in the response.")
        return pd.DataFrame()

# Step 3: Save the DataFrame to Excel, appending if CIN doesn't exist
def append_to_excel(df, filename="company_data1.xlsx"):
    try:
        # Check if the Excel file already exists
        try:
            existing_df = pd.read_excel(filename)
        except FileNotFoundError:
            existing_df = pd.DataFrame()

        if not df.empty:
            # Merge the existing data with the new data, ensuring no duplicate CINs
            # Filter the new data to exclude rows with existing CINs
            new_data_filtered = df[~df['CIN'].isin(existing_df['CIN'])]
            
            # Append only the new data (filtered) to the existing dataframe
            if not new_data_filtered.empty:
                merged_df = pd.concat([existing_df, new_data_filtered]).drop_duplicates(subset='CIN', keep='last')
                # Save the merged dataframe back to Excel
                merged_df.to_excel(filename, index=False)
                print(f"Data successfully appended to {filename}")
            else:
                print("No new data to append.")

        else:
            print("No new data to append.")

    except Exception as e:
        print(f"Error while saving to Excel: {e}")

# Step 4: Create a pivot table for count of `DATE_OF_REGISTRATION`
def create_pivot_table(filename="company_data1.xlsx"):
    try:
        # Read the existing data from Excel
        df = pd.read_excel(filename)

        if not df.empty:
            # Create a pivot table with count of rows for each DATE_OF_REGISTRATION
            pivot_df = df.groupby('DATE_OF_REGISTRATION').size().reset_index(name='COUNT')

            # Write the pivot table to a new sheet called 'companies_pivot'
            with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
                pivot_df.to_excel(writer, sheet_name='companies_pivot', index=False)
                print(f"Pivot table successfully written to 'companies_pivot' sheet in {filename}")
        else:
            print("No data available to create pivot table.")
    
    except Exception as e:
        print(f"Error creating pivot table: {e}")


# Step 5: Iterate over dates until today's date
def fetch_and_save_data_until_today(start_date):
    # Convert start date to datetime object
    current_date = datetime.strptime(start_date, "%Y-%m-%d")
    today = datetime.today()

    while current_date <= today:
        # Format the date to string (e.g., '2024-12-27')
        date_str = current_date.strftime("%Y-%m-%d")
        print(f"Fetching data for {date_str}...")

        # Fetch data for the current date
        raw_data = fetch_data_from_api(url, headers, date_str)

        if raw_data:
            # Process data to DataFrame
            df = process_data_to_dataframe(raw_data)

            # Append data to Excel
            append_to_excel(df, "company_data1.xlsx")

        # Increment the date by 1 day
        current_date += timedelta(days=1)

# Main script
if __name__ == "__main__":
    start_date = "2024-12-24"  # Starting date

    while True:
        # Record the current time (run start time)
        current_time = datetime.now()
        print(f"Current run started at: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Run the fetch and save function
        fetch_and_save_data_until_today(start_date)

        # Record the time when the current run finishes
        finish_time = datetime.now()
        print(f"Current run finished at: {finish_time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Calculate and print the next run time (5 minutes later)
        next_run_time = finish_time + timedelta(minutes=5)
        print(f"Next run scheduled at: {next_run_time.strftime('%Y-%m-%d %H:%M:%S')}")

        # Create the pivot table after each run
        create_pivot_table("company_data1.xlsx")

        # Wait for 5 minutes before the next run
        print("Waiting for 5 minutes before next run...")
        time.sleep(300)  # Sleep for 5 minutes (300 seconds)
