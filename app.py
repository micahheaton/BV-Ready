import streamlit as st

# Import your different applications
import TGAreport_generator

def main():
    st.set_page_config(page_title="Report Generators", page_icon=None, layout="wide", initial_sidebar_state="auto")
    st.title("Report Generators")
    st.write("Select a report generator to run:")

    option = st.radio("Select an option", ("M365 Threat Gap Analysis Report Generator", "Report Generator", "Sentinel Report Generator"))

    if option == "M365 Threat Gap Analysis Report Generator":
        TGAreport_generator.app()

if __name__ == "__main__":
    main()