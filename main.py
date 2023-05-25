import os
from datetime import datetime
import streamlit as st
from dotenv import load_dotenv
from contracts import execute_contracts
from opportunities import execute_opportunities
from subs import execute_subs

def main():
    load_dotenv()
    st.set_page_config(page_title="Monthly Renewal Process")
    st.header("Monthly Renewal Process")

    # Sidebar guide
    st.sidebar.header("User Guide")
    st.sidebar.markdown(
        """
        Start with the Corporate Renewals file

        1. Upload the Renewals file, and click "Generate Contracts".
        2. Upload the Contract Insert Success file, and click "Generate Opportunities".
        3. Upload the Opportunity Insert Success file, and a Subscription Export in the second uploader, and click "Generate OLIs".
        4. Redo steps 1-3 with the PTT files.
        """
    )

    xlsx1 = st.file_uploader("To generate Renewal Contracts and Opportunities", type="xlsx")
    xlsx2 = st.file_uploader("Upload Subscription Export", type="xlsx")

    if xlsx1 is not None and xlsx2 is not None:
        if st.button("Generate Subs"):
            execute_subs(xlsx1, xlsx2)
            st.success("Subscriptions generated successfully")
            show_download_button("renewalOLIinsert.csv")
    elif xlsx1 is not None:
        generate_contracts = st.button("Generate Contracts")
        generate_opportunities = st.button("Generate Opportunities")

        if generate_contracts:
            execute_contracts(xlsx1)
            st.success("Contracts generated successfully")
            show_download_button("contracts_output.xlsx")
        if generate_opportunities:
            execute_opportunities(xlsx1)
            st.success("Opportunities generated successfully")
            show_download_button("opportunities_output.xlsx")
    else:
        st.info("Upload files to get started!")

def show_download_button(file_name):
    folder_name = get_next_month_renewals_folder()
    folder_path = os.path.join("C:/Users/lindb/Documents/CBL", folder_name)
    file_path = os.path.join(folder_path, file_name)

    if os.path.exists(file_path):
        with open(file_path, "rb") as file:
            contents = file.read()
        st.download_button("Download", data=contents, file_name=file_name)
    else:
        st.warning(f"File '{file_name}' not found in the specified folder.")

def get_next_month_renewals_folder():
    now = datetime.now()
    next_month = now.month + 2 if now.month + 2 <= 12 else (now.month + 2) % 12
    year = now.year if now.month + 2 <= 12 else now.year + 1
    folder_name = f"{datetime(year, next_month, 1).strftime('%B')} Renewals"
    return folder_name

if __name__ == '__main__':
    main()
