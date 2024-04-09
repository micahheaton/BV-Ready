import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from io import BytesIO
from analysis_functions import *

def select_experiences():
    experience_options = ["M365 Analysis", "Sentinel Analysis", "Defender for Cloud + Microsoft XDR Analysis", "Existing Customer (QBR/Renewal)"]
    selected_experiences = st.multiselect("Select Experiences", experience_options, key="select_experiences")
    st.session_state.selected_experiences = selected_experiences
    return selected_experiences

def load_and_process_data(file, form_data):
    # Load data from the XLSX file
    xlsx_data = BytesIO(file.read())
    xlsx_file = pd.ExcelFile(xlsx_data)

    # Load each sheet into a DataFrame
    customer_profile_df = pd.read_excel(xlsx_file, sheet_name="Customer_Profile")
    m365_scoping_df = pd.read_excel(xlsx_file, sheet_name="M365_Scoping")
    log_sources_df = pd.read_excel(xlsx_file, sheet_name="Log Sources")
    data_volume_df = pd.read_excel(xlsx_file, sheet_name="Step 1 - MS Data Volume Estimation")
    cost_estimation_df = pd.read_excel(xlsx_file, sheet_name="Step 2 - MS Sentinel Cost Estimation")
    commitment_tier_df = pd.read_excel(xlsx_file, sheet_name="Commitment Tier")

    # Extract key information from customer_profile_df
    company_name = customer_profile_df.iloc[0]["Customer Name"]
    total_employees = customer_profile_df.iloc[0]["Total Number of Employees"]
    compliance_requirements = customer_profile_df.iloc[0]["Compliance Requirements"]
    existing_siem_solution = customer_profile_df.iloc[0]["Existing SIEM Solution"]
    data_retention_policy = customer_profile_df.iloc[0]["Data Retention Policy"]
    number_of_workstations = customer_profile_df.iloc[0]["Number of Endpoint Environment - Workstations"]
    number_of_servers = customer_profile_df.iloc[0]["Number of Endpoint Environment - Servers"]
    current_edr_technology = customer_profile_df.iloc[0]["Current EDR Technology"]

    # Analyze M365 scoping data
    m365_products = m365_scoping_df.columns[m365_scoping_df.iloc[0] == "Yes"].tolist()
    m365_licenses = m365_scoping_df.iloc[0]["Microsoft Licensing Type"]
    aad_integration = m365_scoping_df.iloc[0]["Is AAD already populated with user accounts?"]
    aad_connect_configured = m365_scoping_df.iloc[0]["Is AAD Connect already configured?"]
    on_premises_domain = m365_scoping_df.iloc[0]["Is there a on-premises domain that is connected or needs to be connected to AAD?"]
    email_hosted_in_m365 = m365_scoping_df.iloc[0]["Is Email already hosted in M365?"]

    # Analyze log sources data
    log_source_counts = log_sources_df.iloc[:, 0].value_counts()
    log_source_list = log_sources_df.iloc[:, 0].tolist()

    # Analyze data volume estimation
    total_data_volume = data_volume_df["GB/day"].sum()
    data_volume_by_source = data_volume_df.groupby("Data Sources")["GB/day"].sum()
    data_volume_by_tier = data_volume_df.groupby("Commitment Tier Pricing")["GB/day"].sum()
    azure_defender_benefit = data_volume_df.iloc[0]["Ingestion Offset Insights"]

    # Analyze cost estimation data
    selected_commitment_tier = cost_estimation_df.iloc[0]["Commitment Tier"]
    estimated_monthly_cost = cost_estimation_df.iloc[0]["Estimated Spend"]
    log_analytics_cost = cost_estimation_df.iloc[0]["Log Analytics Subtotal"]
    sentinel_cost = cost_estimation_df.iloc[0]["Sentinel Subtotal"]
    data_retention_cost = cost_estimation_df.iloc[0]["Data Retention"]

    # Analyze commitment tier data
    commitment_tier_prices = commitment_tier_df.set_index("Tier")["Effective Per GB Price1"]
    commitment_tier_savings = commitment_tier_df.set_index("Tier")["Savings Over Pay-As-You-Go"]

    # Prepare data for the report template
    report_data = {
        "executive_summary": {
            "company_name": company_name,
            "total_employees": total_employees,
            "current_setup": f"Existing SIEM: {existing_siem_solution}, Data Retention Policy: {data_retention_policy}",
            "key_findings": [
                "Placeholder for key finding 1",
                "Placeholder for key finding 2",
                "Placeholder for key finding 3"
            ],
            "recommendations": [
                "Placeholder for recommendation 1",
                "Placeholder for recommendation 2",
                "Placeholder for recommendation 3"
            ]
        },
        "sentinel_deployment_assessment": {
            "current_state": form_data.get("current_usage"),
            "log_sources": {
                "current_sources": log_source_list,
                "potential_gaps": "Placeholder for potential log source gaps"
            },
            "external_ingestion": {
                "challenges": form_data.get("log_source_pain_points"),
                "areas_for_expansion": "Placeholder for areas of expansion"
            },
            "data_volume_and_cost": {
                "total_data_volume": total_data_volume,
                "data_volume_by_source": data_volume_by_source,
                "estimated_monthly_cost": estimated_monthly_cost,
                "log_analytics_cost": log_analytics_cost,
                "sentinel_cost": sentinel_cost,
                "data_retention_cost": data_retention_cost
            }
        },
        "content_and_detection_optimization": {
            "mitre_attack_alignment": "Placeholder for MITRE ATT&CK alignment assessment",
            "detection_strategy": "Placeholder for detection strategy assessment",
            "hunting_opportunities": "Placeholder for hunting opportunities assessment"
        },
        "next_steps": {
            "quick_wins": [
                "Placeholder for quick win 1",
                "Placeholder for quick win 2"
            ],
            "strategic_initiatives": [
                "Placeholder for strategic initiative 1",
                "Placeholder for strategic initiative 2"
            ]
        },
        "additional_data": {
            "compliance_requirements": compliance_requirements,
            "endpoints": f"{number_of_workstations} workstations, {number_of_servers} servers",
            "current_edr": current_edr_technology,
            "m365_products": m365_products,
            "m365_licenses": m365_licenses,
            "aad_integration": aad_integration,
            "aad_connect_configured": aad_connect_configured,
            "on_premises_domain": on_premises_domain,
            "email_hosted_in_m365": email_hosted_in_m365,
            "log_source_counts": log_source_counts,
            "data_volume_by_tier": data_volume_by_tier,
            "azure_defender_benefit": azure_defender_benefit,
            "commitment_tier_prices": commitment_tier_prices,
            "commitment_tier_savings": commitment_tier_savings
        },
        "user_input": {
            "key_challenges": form_data.get("key_challenges"),
            "proactive_hunting_interest": form_data.get("proactive_hunting_interest"),
            "current_threat_detection": form_data.get("current_threat_detection"),
            "specific_log_challenges": form_data.get("specific_log_challenges"),
            "cost_optimization_priority": form_data.get("cost_optimization_priority"),
            "additional_notes": form_data.get("additional_notes")
        }
    }

    return report_data

def generate_consolidated_report(selected_experiences, uploaded_files):
    data = {}

    if "M365 Analysis" in selected_experiences:
        m365_files = [file for file in uploaded_files if file.name.endswith(".csv")]
        data["m365_results"] = process_m365_files(m365_files)

    if "Sentinel Analysis" in selected_experiences:
        sentinel_files = [file for file in uploaded_files if file.name.endswith(".xlsx")]
        data["sentinel_results"] = process_sentinel_files(sentinel_files)

    if "Defender for Cloud + Microsoft XDR Analysis" in selected_experiences:
        defender_cloud_files = [file for file in uploaded_files if file.name.endswith(".csv") or file.name.endswith(".xlsx")]
        data["defender_cloud_results"] = process_defender_cloud_files(defender_cloud_files)

    return render_report(data)

def render_report(data):
    file_loader = FileSystemLoader('templates')
    env = Environment(loader=file_loader)
    template = env.get_template("report_template.html")

    m365_results = data.get("m365_results", {})
    sentinel_results = data.get("sentinel_results", {})
    defender_cloud_results = data.get("defender_cloud_results", {})

    executive_summary = sentinel_results.get("executive_summary", {})
    sentinel_optimization_assessment = sentinel_results.get("sentinel_deployment_assessment", {})
    cost_analysis = sentinel_results.get("data_volume_and_cost", {})

    secure_score = m365_results.get("secure_score", {})
    category_scores = secure_score.get("category_scores", []) if secure_score else []

    incidents = m365_results.get("incidents", {})
    incident_severity_counts = incidents.get("incident_severity_counts", {})

    output = template.render(
        m365_results=m365_results,
        sentinel_results=sentinel_results,
        defender_cloud_results=defender_cloud_results,
        executive_summary=executive_summary,
        sentinel_optimization_assessment=sentinel_optimization_assessment,
        cost_analysis=cost_analysis,
        secure_score=secure_score,
        category_scores=category_scores,
        incident_severity_counts=incident_severity_counts or {},
        device_risk=m365_results.get("device_risk", {}),
        product_deployment=m365_results.get("product_deployment", {}),
        user_risk=m365_results.get("user_risk", {}),
        incidents=incidents,
        tvm=m365_results.get("tvm", {}),
        current_usage=sentinel_optimization_assessment.get("current_state", ""),
        key_challenges=sentinel_results.get("user_input", {}).get("key_challenges", ""),
        proactive_hunting_interest=sentinel_results.get("user_input", {}).get("proactive_hunting_interest", ""),
        current_threat_detection=sentinel_results.get("user_input", {}).get("current_threat_detection", ""),
        log_source_pain_points=sentinel_optimization_assessment.get("external_ingestion", {}).get("challenges", []),
        specific_log_challenges=sentinel_results.get("user_input", {}).get("specific_log_challenges", ""),
        cost_optimization_priority=sentinel_results.get("user_input", {}).get("cost_optimization_priority", ""),
        additional_notes=sentinel_results.get("user_input", {}).get("additional_notes", "")
    )

    return output

def main():
    st.cache_data.clear()
    st.title("M365 TGA | Sentinel COA | MDC COA and Copilot Readiness Report Generator")
    st.write("Select the desired analyses and upload the corresponding files to get started.")

    selected_experiences = select_experiences()

    # Display the Sentinel Optimization Assessment form if "Sentinel Analysis" is selected
    form_data = {}
    if "Sentinel Analysis" in st.session_state.get("selected_experiences", []):
        st.header("Sentinel Optimization Assessment")

        # 1. Sentinel Optimization
        current_usage = st.selectbox("Current Usage", ["Initial configuration, out-of-the-box setup", "Some customizations, limited content development", "Moderate customization, own detections and rules", "Other (please specify)"])
        key_challenges = ""  # Initialize key_challenges with an empty string
        if current_usage == "Other (please specify)":
            key_challenges = st.text_area("Key Challenges with Current Sentinel Setup", height=100)

        # 2. Proactive Security
        proactive_hunting_interest = st.selectbox("Interest in Threat Hunting", ["Not at this time", "Interested, but no dedicated resources", "Yes, want to develop a hunting program"])
        current_threat_detection = st.text_area("Current Threat Detection Approach", height=100)

        # 3. Log Management & Complexity
        log_source_pain_points = st.multiselect("Main Log Source Pain Points (Select all that apply)", ["Integrating new log sources", "Managing data volume and costs", "Maintaining consistent log pipelines", "Complexity due to multi-vendor environment"])
        specific_log_challenges = st.text_area("Specific Log-related Challenges", height=100)

        # 4. Cost vs. Value
        cost_optimization_priority = st.selectbox("Cost Optimization Priority", ["High – cost reduction is crucial", "Moderate – balanced with security improvements", "Low – security improvements are paramount"])

        # 5. Additional Notes
        additional_notes = st.text_area("Other Concerns/Considerations", height=100)

        # Store the form data in a dictionary
        form_data = {
            "current_usage": current_usage,
            "key_challenges": key_challenges,
            "proactive_hunting_interest": proactive_hunting_interest,
            "current_threat_detection": current_threat_detection,
            "log_source_pain_points": log_source_pain_points,
            "specific_log_challenges": specific_log_challenges,
            "cost_optimization_priority": cost_optimization_priority,
            "additional_notes": additional_notes
        }

    uploaded_files = st.file_uploader("Upload Files", type=["csv", "xlsx"], accept_multiple_files=True)

    if st.button("Generate Report"):
        if selected_experiences and uploaded_files:
            consolidated_report = generate_consolidated_report(selected_experiences, uploaded_files)
            st.components.v1.html(consolidated_report, height=600, scrolling=True)
            st.download_button(label="Download Report", data=consolidated_report, file_name="Consolidated_Report.html")
        else:
            st.warning("Please select at least one experience and upload the corresponding files.")

if __name__ == "__main__":
    main()