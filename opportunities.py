import pandas as pd

def execute_opportunities(filepath):
    df = pd.read_excel(io=filepath)

    # Delete column 1 and 2
    del df['STATUS']
    del df['COMMENTS']
    del df['Owner Expiration Notice']

    # Rename Column ID to Contract ID
    df.rename(columns={'ID': 'Contract ID'}, inplace=True)

    # Create new column "Previous Contract End Date" and populate it with Column "Contract Start Date" - 1 day
    df.insert(6, 'Previous Contract End Date', df['Contract Start Date'] - pd.Timedelta(days=1))
    df.insert(6, 'Close Date', df['Contract Start Date'] - pd.Timedelta(days=30))
    df['Previous Contract End Date'] = df['Previous Contract End Date'].dt.strftime('%m/%d/%Y')
    df['Close Date'] = df['Close Date'].dt.strftime('%m/%d/%Y')

    # Add "Months" to the end of the Contract Term
    df['Contract Term'] = df['Contract Term'].astype(str) + " Months"

    # Rename Column "Contract Type" to "Type" and "Status" to "Contract Status"
    df.rename(columns={'Contract Type': 'Type'}, inplace=True)
    df.rename(columns={'Status': 'Contract Status'}, inplace=True)
    df.rename(columns={'Renewal': 'Renewal Opportunity'}, inplace=True)

    #insert column "Stage" and populate it with "Qualifying"
    df.insert(6, 'stage', 'Qualifying')

    # insert column next step and set to Renew and column Approved by manager and set to FALSE
    df.insert(7, 'Next Step', 'Renew')
    df.insert(8, 'Approved by Manager', 'FALSE')


    # change column name "Active Corporate Subscriptions" to "Potential Renewals"
    df.rename(columns={'Active Corporate Subscriptions': 'Potential Renewals'}, inplace=True)

    # Create a column "Previous Contract MRR" and if Type is Corporate then populate it with "ALIV Monthly Spend" - "ALIV MRR PTT" and if Type is PTT then populate it with "ALIV MRR PTT"
    df.insert(9, 'Previous Contract MRR', df['ALIV Monthly Spend'] - df['ALIV MRR PTT'])
    df.loc[df['Type'] == 'PTT', 'Previous Contract MRR'] = df['ALIV MRR PTT']

    # Create Column "Price book ID" set to "01s41000007XSvaAAG" 
    df.insert(10, 'Price Book ID', '01s41000007XSvaAAG')

    #Create column "Porting required" and set to Not Required
    df.insert(11, 'Porting Required', 'Not Required')

    df['Contract Start Date'] = df['Contract Start Date'].dt.strftime('%m/%d/%Y')

    # Insert column "Name" and populate it with Account Name + " Renewal" + Contract Term + Type + " Contract " + Contract Start Date. TypeError: unsupported operand type(s) for +: 'Timestamp' and 'str'
    df.insert(3, 'Name', df['Account Name'] + " Renewal " + df['Contract Term'] + " " + df['Type'] + " Contract " + df['Contract Start Date'])
    # Add Column after Contract Name "Len" and apply the len formula to the Name column. first delete the current len column
    # del df['Len']
    # df.insert(4, 'Len', df['Name'].str.len())

    # insert column recordtype ID = 01241000000L1DSAA0
    df.insert(5, 'Record Type ID', '01241000000L1DSAA0')

    # delete columns Account Name, Contract Name
    del df['Account Name']
    del df['Contract Name']




# save a new excel file "renewalsOpportunities.xlsx"
    df.to_csv('C:/Users/lindb/Documents/CBL/Renewals/renewalsOpportunities.csv', index=False)

if __name__ == '__main__':
    execute_opportunities('RenewalOpportunities.xlsx')
