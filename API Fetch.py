import requests
import pandas as pd

# API endpoint and headers
url = "https://api.finanvo.in/company/newcompanies/list?type=date&bulkMasterType=CIN&date=2024-12-27"
headers = {
    "Content-Type": "application/json",
    "x-api-key": "OliINWwpCj",
    "x-api-secret-key": "smdlme2Y9n25v8htM0uQkWGqxGdN3D8O8aLkf9Zl"
}

# Step 1: Fetch data from API
def fetch_data_from_api(url, headers):
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise an error for HTTP issues
        return response.json()  # Return the parsed JSON response
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data from API: {e}")
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

# Step 3: Save the DataFrame to Excel
def save_to_excel(df, filename="company_data.xlsx"):
    if not df.empty:
        df.to_excel(filename, index=False)  # Save without row index
        print(f"Data successfully saved to {filename}")
    else:
        print("No data to save.")

# Main script
if __name__ == "__main__":
    # Fetch data from API
    raw_data = fetch_data_from_api(url, headers)
    
    if raw_data:
        # Process data to DataFrame
        df = process_data_to_dataframe(raw_data)
        
        if not df.empty:
            # Save the DataFrame to an Excel file
            save_to_excel(df, filename="company_data.xlsx")
        else:
            print("No valid data to display.")
    else:
        print("No data received or invalid format.")
