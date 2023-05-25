import pandas as pd
from datetime import datetime
import os
from dateutil.relativedelta import relativedelta

def execute_contracts(xlsx1):
    

    df = pd.read_excel(xlsx1)

    # Delete Column "Contract Start Date"
    del df['Contract Start Date']

    # Add Column "Contract Start Date" and populate with data from column "Contract End Date" + 1 day and add it right after column "Contract End Date"
    df.insert(6, 'Contract Start Date', df['Contract End Date'] + pd.Timedelta(days=1))
    # Change the format to mm/dd/yyyy
    df['Contract Start Date'] = df['Contract Start Date'].dt.strftime('%m/%d/%Y')

    # Delete Column "Contract End Date"
    del df['Contract End Date']

    # Change the name of column Contract Term (Months) to Contract Term
    df.rename(columns={'Contract Term (months)': 'Contract Term'}, inplace=True)

    # Add Column "Contract Name" and use this excel formula to populate it, using the column headers as indicators: = Account Name & " " & Contract Type & " Contract " & Term & " Months " & TEXT(Contract Start Date, "mm/dd/yyyy") Contract Start date is not in datetime format, so can not use dt/strftime
    df.insert(3, 'Contract Name', df['Account Name'] + " " + df['Contract Type'] + " Contract " + df['Contract Term'].astype(str) + " Months " + df['Contract Start Date'])

    # Add Column after Contract Name "Len" and apply the len formula to the Contract Name column
    df.insert(4, 'Len', df['Contract Name'].str.len())

    # Add colulmn "Owner Expiration Notice" and populate each row with 15
    df.insert(9, 'Owner Expiration Notice', 15)

    # Change the values in column "Status" to "Draft"
    df['Status'] = 'Draft'

    # Add Column "Renewal" set values equal to "TRUE"
    df.insert(11, 'Renewal', 'TRUE')

    # Change name of Case Safe ID to Account ID
    df.rename(columns={'Case Safe ID': 'Account ID'}, inplace=True)
    # Delete Columns "Case Safe ID" and "Contract Number" and "Case Safe Contract ID" and "Corporate Contract Status"
    del df['Contract Number']
    del df['Case Safe Contract ID']
    del df['Corporate Contract Status']

# Create a folder for storing the generated files
    now = datetime.now()
    current_month = now.month
    next_month = now + relativedelta(months=2)
    folder_name = f"{next_month.strftime('%B')} Renewals"
    year = next_month.year

    folder_path = os.path.join("C:/Users/lindb/Documents/CBL", folder_name)
    os.makedirs(folder_path, exist_ok=True)

    # Save a copy of the file as a CSV UTF-8 inside the new folder
    file_path = os.path.join(folder_path, "RenewalsContracts.csv")
    df.to_csv(file_path, index=False)
        # Save a copy of the file as a CSV UTF-8
        # df.to_csv('RenewalsContracts.csv', index=False)

# Execute the contracts when this file is run directly
if __name__ == '__main__':
    execute_contracts('RenewalsContracts.xlsx')
