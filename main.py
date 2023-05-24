import streamlit as st
from dotenv import load_dotenv
from contracts import execute_contracts
from opportunities import execute_opportunities
# from subs import execute_subs

def main():
    load_dotenv()
    st.set_page_config(page_title="Monthly Renewal Process")
    st.header("Monthly Renewal Process")

    xlsx = st.file_uploader("Upload your excel file", type="xlsx")

    if xlsx is not None:
        generate_contracts = st.button("Generate Contracts")
        generate_opportunities = st.button("Generate Opportunities")
        generate_subs = st.button("Generate Subs")

        if generate_contracts:
            execute_contracts(xlsx)
            st.success("Contracts generated successfully")

        if generate_opportunities:
            execute_opportunities(xlsx)
            st.success("Opportunities generated successfully")

        # if generate_subs:
        #     execute_subs(xlsx)
        #     st.success("Subs generated successfully")

        

if __name__ == '__main__':
    main()
