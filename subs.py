import pandas as pd
from datetime import datetime
import os
from dateutil.relativedelta import relativedelta

def execute_subs(xlsx1, xlsx2):
    now = datetime.now()
    current_month = now.month
    next_month = now + relativedelta(months=2)
    folder_name = f"{next_month.strftime('%B')} Renewals"
    year = next_month.year

    folder_path = os.path.join("C:/Users/lindb/Documents/CBL", folder_name)

    if os.path.exists(folder_path):
        
        df1 = pd.read_excel(xlsx1)
        df2 = pd.read_excel(xlsx2)

        if 'subscription' in xlsx2.name.lower():
            df_subs = df2.copy()  # Create a copy of df2 to avoid modifying the original DataFrame
            df_opp = df1.copy()  # Create a copy of df1 to avoid modifying the original DataFrame
        else:
            print("Please make sure to include both files, Opportunity Insert Success file and the Subscription Export.")
            return

        # Config the Subs export file
        # Delete all rows where the column "Subscription Status" is not "Active" or "Setup" or "Suspended"
        df_subs = df_subs[df_subs['Subscription Status'].isin(['Active', 'Setup', 'Suspended'])]

        # Rename columns using .rename() and .loc accessor
        df_subs = df_subs.rename(columns={'Recurring Charge': 'Sales Price'})
        df_subs['Key1'] = df_subs['Account'] + df_subs['Contract Type']
        
        # rename df_opp column ID to Opportunity ID
        df_opp = df_opp.rename(columns={'ID': 'Opportunity ID'})

        # Create Key1 column on opp insert success which consists of Account and Type
        df_opp['Key1'] = df_opp['Account ID'] + df_opp['Type']
        # Create Key1 column on subs export which consists of Account and Contract Type
        df_subs['Key1'] = df_subs['Account'] + df_subs['Contract Type']

        df_merged = df_subs.merge(df_opp, left_on='Key1', right_on='Key1')

        df_merged['Key2'] = df_merged['Plan__r.Product ID'] + df_merged['Department']

        # Create a Quantity Column to count how many have the same exact department and plan__r.product.id pairing
        df_merged['Quantity'] = df_merged.groupby('Key2')['Key2'].transform('count')

        # Remove duplicates from df_merged based on column Key2
        df_merged.drop_duplicates(subset='Key2', keep='first', inplace=True)

        # Delete all columns except: Account, ID, Sales Price, Quantity, Department, Plan__r.Product ID, Plan__r.Product Family Detail, Plan__r.Product Family
        df_merged = df_merged[['Account', 'Opportunity ID', 'Sales Price', 'Quantity', 'Department', 'Plan__r.Product ID', 'Plan__r.Product Family Detail', 'Plan__r.Product Family']]

        # Create new columns "SIM" = "Existing Postpaid" and "Customer Sale Type" = "EC Renewal"
        df_merged.insert(0, 'SIM', 'Existing Postpaid')
        df_merged.insert(1, 'Customer Sale Type', 'EC - Renewal')

        # Rename columns: Plan__r.Product ID to Product ID, Plan__r.Product Family Detail to Product Family Detail, Plan__r.Product Family to Product Family
        df_merged.rename(columns={'Plan__r.Product ID': 'Product ID', 'Plan__r.Product Family Detail': 'Product Family Detail', 'Plan__r.Product Family': 'Product Family'}, inplace=True)

        print("Columns in df_subs:", df_subs.columns)
        print("Columns in df_opp:", df_opp.columns)
        # Reorder the columns
        df_merged = df_merged[['Account', 'Opportunity ID', 'SIM', 'Customer Sale Type', 'Sales Price', 'Quantity', 'Department', 'Product ID', 'Product Family Detail', 'Product Family']]

        # Save a copy of the file as a CSV UTF-8 inside the existing folder
        file_path = os.path.join(folder_path, "renewalOLIinsert.csv")
        df_merged.to_csv(file_path, index=False)
    else:
        print("File path does not exist. Make sure to execute 'contracts.py' first.")


if __name__ == '__main__':
    execute_subs('OpportunityInsertSuccess.xlsx', 'SubscriptionExport.xlsx')
