import streamlit as st
import subprocess

def main():
    st.set_page_config(page_title="Report Generators", page_icon=None, layout="wide", initial_sidebar_state="auto")
    
    st.title("Report Generators")
    
    st.write("Select a report generator to run:")
    
    if st.button("M365 Threat Gap Analysis Report Generator"):
        subprocess.Popen(["streamlit", "run", "TGAreport_generator.py"])
    
    if st.button("Report Generator"):
        subprocess.Popen(["streamlit", "run", "report_generator.py"])
    
    if st.button("Sentinel Report Generator"):
        subprocess.Popen(["streamlit", "run", "sentinel_report_generator.py"])

if __name__ == "__main__":
    main()