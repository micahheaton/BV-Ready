import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from io import BytesIO

def load_and_process_data(file):
    # Load data from the XLSX file
    xlsx_data = BytesIO(file.read())
    xlsx_file = pd.ExcelFile(xlsx_data)

    # Load each sheet into a DataFrame
    customer_profile_df = pd.read_excel(xlsx_file, sheet_name="Customer_Profile")
    m365_scoping_df = pd.read_excel(xlsx_file, sheet_name="M365_Scoping")
    log_sources_df = pd.read_excel(xlsx_file, sheet_name="Log Sources")
    data_volume_df = pd.read_excel(xlsx_file, sheet_name="Step 1 - MS Data Volume Estimat")
    cost_estimation_df = pd.read_excel(xlsx_file, sheet_name="Step 2- MS Sentinel Cost Estima")
    commitment_tier_df = pd.read_excel(xlsx_file, sheet_name="Commitment Tier")

    # Extract key information from customer_profile_df
    company_name = customer_profile_df.iloc[3][1]["Customer Name"]
    total_employees = customer_profile_df.iloc[5][1]["Total Number of Employees"]
    hq_country = customer_profile_df.iloc[4][1]["HQ Country"]
    delivery_location = customer_profile_df.iloc[6][1]["Delivery Location/Azure region"]
    compliance_requirements = customer_profile_df.iloc[8][1]["Compliance Requirements"]
    existing_siem_solution = customer_profile_df.iloc[11][1]["Existing SIEM Solution"]
    data_retention_policy = customer_profile_df.iloc[13][1]["Data Retention Policy"]
    workstations_count = customer_profile_df.iloc[14][1]["Number of Endpoint Environment - Workstations"]
    servers_count = customer_profile_df.iloc[15][1]["Number of Endpoint Environment - Servers"]
    current_edr_technology = customer_profile_df.iloc[17][1]["Current EDR Technology"]
    domain_names = customer_profile_df.iloc[18][1]["Domain Name(s)"]
    azure_tenant_id = customer_profile_df.iloc[19][1]["Azure Tenant ID"]
    sentinel_instances_count = customer_profile_df.iloc[20][1]["Number of Sentinel Instances"]

    # Extract key information from m365_scoping_df
    m365_customer = m365_scoping_df.iloc[2][1]["Microsoft 365 Customer?"]
    m365_licensing_type = m365_scoping_df.iloc[3][1]["Microsoft Licensing Type:"]
    azure_infrastructure = m365_scoping_df.iloc[4][1]["Azure Infrastructure:"]
    endpoint_management_tools = m365_scoping_df.iloc[5][1]["Endpoint Management Tools:"]

    # Analyze log sources data
    log_source_counts = log_sources_df['Needed Log Sources'].value_counts()
    log_source_list = log_sources_df['Needed Log Sources'].tolist()
    planned_integrations = log_sources_df[log_sources_df['Notes'].str.contains('Palo Alto', na=False, case=False)]['Needed Log Sources'].tolist()

    # Load data into a DataFrame
    data_volume_df = pd.DataFrame({
        'Data Sources': ['Azure AD Audit (Users)', 'Azure AD Sign-ins (Users)', 'Windows Servers w/ high EPS', 'Windows Servers w/ medium EPS', 'Windows Servers w/ low EPS', 'Windows Domain Server', 'Windows Desktops (Laptops, Tablets, POS)', 'HyperVisor (ESXi, Hyper-V etc)', 'Linux / Unix Servers', 'Network Firewalls (DMZ)', 'Network Firewalls (Internal)', 'Network Flows (NetFlow/S-Flow)', 'Network IPS/IDS', 'Network Load-Balancers', 'Network Gateway/Routers', 'Network Switches', 'Network VPN / SSL VPN', 'Network Web Proxy', 'Network Wireless LAN', 'Other Network Devices', 'Other Security Devices', 'Azure Tenants', 'Ingestion Offset Insights', 'E5 Licensed Users', 'Defender for Server P2 Servers'],
        'Nodes or End Points or Users': [600, 600, 20, 20, 20, 5, 600, None, None, 1, 3, 3, 3, 3, 3, 3, 2, 2, 2, 2, 2, 1, None, 600, 60],
        'Avg. Event Size (bytes)': [2048, 800, 700, 700, 700, 1000, 746, 1000, 300, 250, 250, 400, 300, 150, 250, 100, 300, 650, 150, 250, 750, None, None, 5, 500],
        'Avg. EPS per node or end-point or user': [0.000174, 0.001736, 7, 3, 1, 7, 0.0005, 15, 3, 50, 240, 30, 100, 5, 1, 30, 2, 20, 5, 10, 5, None, None, 1000, 1000],
        'EPS/Source': [0, 1, 140, 60, 20, 35, 0.30, None, None, 50, 720, 90, 300, 15, 3, 90, 4, 40, 10, 20, 10, None, None, None, None],
        'GB/day': [0, 0, 8, 3, 1, 3, 0.02, 0, 0, 1, 14, 3, 7, 0, 0, 1, 0, 2, 0, 0, 1, None, 3, 3, 30],
        'Commitment Tier Pricing': ['Region', 'East US 2', 'Tier', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None]
    })

        # Analyze data volume estimation
    total_data_volume = data_volume_df['GB/day'].sum()
    data_volume_by_source = data_volume_df.groupby('Data Sources')['GB/day'].sum()
    data_volume_by_tier = data_volume_df[data_volume_df['Commitment Tier Pricing'].notnull()].groupby('Commitment Tier Pricing')['GB/day'].sum()
    azure_defender_benefit = data_volume_df[data_volume_df['Data Sources'] == 'Ingestion Offset Insights']['GB/day'].values[0]


    # Analyze cost estimation data
    expected_data_ingestion = cost_estimation_df.iloc[7][1]  # Value from row 8, column 1
    free_data_ingestion = cost_estimation_df.iloc[8][1]  # Value from row 9, column 1
    net_billable_data = expected_data_ingestion - free_data_ingestion
    include_log_analytics = cost_estimation_df.iloc[9][1]  # Value from row 10, column 1
    azure_region = cost_estimation_df.iloc[10][1]  # Value from row 11, column 1
    currency = cost_estimation_df.iloc[11][1]  # Value from row 12, column 1
    customer_discount = cost_estimation_df.iloc[13][1]  # Value from row 14, column 1
    data_retention_period = cost_estimation_df.iloc[14][1]  # Value from row 15, column 1

    # Log Analytics
    log_analytics_pay_go_gb_day = cost_estimation_df.iloc[18][1]  # Value from row 19, column 1
    log_analytics_capacity_reservation_gb_day = cost_estimation_df.iloc[19][1]  # Value from row 20, column 1
    log_analytics_cost = log_analytics_pay_go_gb_day + log_analytics_capacity_reservation_gb_day

    # Azure Sentinel
    sentinel_pay_go_gb_day = cost_estimation_df.iloc[21][1]  # Value from row 22, column 1
    sentinel_capacity_reservation_gb_day = cost_estimation_df.iloc[22][1]  # Value from row 23, column 1
    sentinel_cost = sentinel_pay_go_gb_day + sentinel_capacity_reservation_gb_day

    # Data Retention
    data_retention_gb_day = cost_estimation_df.iloc[24][1]  # Value from row 25, column 1
    data_retention_cost = data_retention_gb_day

        # Set default values for undefined variables
    estimated_monthly_cost = 0
    azure_defender_vm_count = 0
    azure_defender_benefit_gb_per_day = 0

    # Analyze commitment tier data
    commitment_tier_prices = commitment_tier_df.set_index("Tier")["Effective Per GB Price1"]
    commitment_tier_savings = commitment_tier_df.set_index("Tier")["Savings Over Pay-As-You-Go"]

    # Prepare data for visualizations
    log_source_counts_data = {
        "labels": log_source_counts.index.tolist(),
        "values": log_source_counts.values.tolist()
    }

    data_volume_by_source_data = {
        "labels": data_volume_by_source.index.tolist(),
        "values": data_volume_by_source.values.tolist()
    }

    data_volume_by_tier_data = {
        "labels": data_volume_by_tier.index.tolist(),
        "values": data_volume_by_tier.values.tolist()
    }

    cost_breakdown_data = {
        "labels": ["Log Analytics", "Sentinel", "Data Retention"],
        "values": [log_analytics_cost, sentinel_cost, data_retention_cost]
    }

    commitment_tier_prices_data = {
        "labels": commitment_tier_prices.index.tolist(),
        "values": commitment_tier_prices.values.tolist()
    }

    commitment_tier_savings_data = {
        "labels": commitment_tier_savings.index.tolist(),
        "values": commitment_tier_savings.values.tolist()
    }

    # Prepare data for the report template
    report_data = {
        "executive_summary": {
            "company_name": company_name,
            "total_employees": total_employees,
            "hq_country": hq_country,
            "delivery_location": delivery_location,
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
            "current_state": {
                "sentinel_setup": st.session_state.sentinel_setup,
                "m365_licensing": m365_licensing_type,
                "defender_products": m365_scoping_df.iloc[0]["Which M365 Defender products are currently deployed?"],
                "workstations_count": workstations_count,
                "servers_count": servers_count,
                "current_edr_technology": current_edr_technology,
                "domain_names": domain_names,
                "azure_tenant_id": azure_tenant_id,
                "sentinel_instances_count": sentinel_instances_count
            },
            "log_sources": {
                "current_sources": log_source_list,
                "potential_gaps": "Placeholder for potential log source gaps",
                "planned_integrations": planned_integrations
            },
            "external_ingestion": {
                "challenges": st.session_state.external_ingestion_challenges,
                "solutions": "Placeholder for external ingestion solutions"
            },
            "data_volume_and_cost": {
                "total_data_volume": total_data_volume,
                "data_volume_by_source": data_volume_by_source_data,
                "data_volume_by_tier": data_volume_by_tier_data,
                "estimated_monthly_cost": estimated_monthly_cost,
                "cost_breakdown": cost_breakdown_data,
                "azure_defender_benefit": azure_defender_benefit,
                "azure_defender_vm_count": azure_defender_vm_count,
                "azure_defender_benefit_gb_per_day": azure_defender_benefit_gb_per_day
            }
        },
        "content_and_detection_optimization": {
            "mitre_attack_alignment": {
                "current_state": st.session_state.mitre_alignment,
                "recommendations": "Placeholder for MITRE ATT&CK alignment recommendations"
            },
            "detection_strategy": {
                "current_state": st.session_state.detection_strategy,
                "recommendations": "Placeholder for detection strategy recommendations"
            },
            "hunting_opportunities": {
                "current_state": st.session_state.hunting_opportunities,
                "recommendations": "Placeholder for hunting opportunities recommendations"
            }
        },
        "integration_and_log_source_management": {
            "integration_challenges": {
                "current_state": st.session_state.integration_challenges,
                "solutions": "Placeholder for integration challenge solutions"
            },
            "log_source_prioritization": {
                "current_state": st.session_state.log_source_prioritization,
                "recommendations": "Placeholder for log source prioritization recommendations"
            }
        },
        "cost_benefit_analysis": {
            "data_costs_vs_detection_opportunities": {
                "current_state": st.session_state.data_costs_vs_detection,
                "analysis": "Placeholder for data costs vs. detection opportunities analysis"
            },
            "compliance_and_legal_requirements": {
                "current_state": st.session_state.compliance_and_legal,
                "analysis": "Placeholder for compliance and legal requirements analysis"
            },
            "commitment_tier_prices": commitment_tier_prices_data,
            "commitment_tier_savings": commitment_tier_savings_data
        },
        "roadmap_and_next_steps": {
            "key_recommendations": [
                "Placeholder for key recommendation 1",
                "Placeholder for key recommendation 2",
                "Placeholder for key recommendation 3"
            ],
            "implementation_roadmap": "Placeholder for implementation roadmap",
            "next_steps": "Placeholder for next steps"
        },
        "additional_data": {
            "compliance_requirements": compliance_requirements,
            "log_source_counts": log_source_counts_data,
            "m365_customer": m365_customer,
            "azure_infrastructure": azure_infrastructure,
            "endpoint_management_tools": endpoint_management_tools,
            "azure_region": azure_region,
            "currency": currency,
            "customer_discount": customer_discount,
            "data_retention_period": data_retention_period
        }
    }

    return report_data

def render_report(data):
    file_loader = FileSystemLoader('templates')
    env = Environment(loader=file_loader)
    template = env.get_template("sentinel_report_template.html")
    output = template.render(data)
    return output

def main():
    st.title("Sentinel Optimization Assessment")

    # Collect additional information not present in the Excel file
    st.header("Additional Information")

    # Current Sentinel Setup
    sentinel_setup_options = [
        "Out-of-the-box configuration",
        "Partially customized",
        "Highly customized",
        "Other"
    ]
    st.session_state.sentinel_setup = st.selectbox("Current Sentinel Setup", sentinel_setup_options)
    if st.session_state.sentinel_setup == "Other":
        st.session_state.sentinel_setup_other = st.text_input("Please specify")

    # MITRE ATT&CK Alignment
    mitre_alignment_options = [
        "Not aligned",
        "Partially aligned",
        "Fully aligned",
        "Unknown"
    ]
    st.session_state.mitre_alignment = st.selectbox("MITRE ATT&CK Alignment", mitre_alignment_options)
    if st.session_state.mitre_alignment == "Partially aligned":
        st.session_state.mitre_alignment_details = st.text_area("Please provide more details")

    # Detection Strategy
    detection_strategy_options = [
        "No defined strategy",
        "Basic detections",
        "Advanced detections",
        "Other"
    ]
    st.session_state.detection_strategy = st.selectbox("Current Detection Strategy", detection_strategy_options)
    if st.session_state.detection_strategy == "Other":
        st.session_state.detection_strategy_other = st.text_input("Please specify")

    # Hunting Opportunities
    hunting_opportunities_options = [
        "Not explored",
        "Ad-hoc hunting",
        "Regular hunting exercises",
        "Other"
    ]
    st.session_state.hunting_opportunities = st.selectbox("Hunting Opportunities", hunting_opportunities_options)
    if st.session_state.hunting_opportunities == "Other":
        st.session_state.hunting_opportunities_other = st.text_input("Please specify")

    # External Ingestion Challenges
    st.session_state.external_ingestion_challenges = st.multiselect(
        "External Ingestion Challenges",
        ["Cost", "Complexity", "Data quality", "Other"]
    )
    if "Other" in st.session_state.external_ingestion_challenges:
        st.session_state.external_ingestion_challenges_other = st.text_input("Please specify")

    # Integration Challenges
    st.session_state.integration_challenges = st.multiselect(
        "Integration Challenges",
        ["Inter-departmental", "Technical complexity", "Vendor support", "Other"]
    )
    if "Other" in st.session_state.integration_challenges:
        st.session_state.integration_challenges_other = st.text_input("Please specify")

    # Log Source Prioritization
    log_source_prioritization_options = [
        "Not prioritized",
        "Partially prioritized",
        "Fully prioritized",
        "Other"
    ]
    st.session_state.log_source_prioritization = st.selectbox("Log Source Prioritization", log_source_prioritization_options)
    if st.session_state.log_source_prioritization == "Other":
        st.session_state.log_source_prioritization_other = st.text_input("Please specify")

    # Data Costs vs. Detection Opportunities
    data_costs_vs_detection_options = [
        "Cost is the primary focus",
        "Balance between cost and detection",
        "Detection is the primary focus",
        "Other"
    ]
    st.session_state.data_costs_vs_detection = st.selectbox("Data Costs vs. Detection Opportunities", data_costs_vs_detection_options)
    if st.session_state.data_costs_vs_detection == "Other":
        st.session_state.data_costs_vs_detection_other = st.text_input("Please specify")

    # Compliance and Legal Requirements
    compliance_and_legal_options = [
        "Not a priority",
        "Moderate priority",
        "High priority",
        "Other"
    ]
    st.session_state.compliance_and_legal = st.selectbox("Compliance and Legal Requirements", compliance_and_legal_options)
    if st.session_state.compliance_and_legal == "Other":
        st.session_state.compliance_and_legal_other = st.text_input("Please specify")

    # File Upload
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file is not None:
        report_data = load_and_process_data(uploaded_file)
        report_output = render_report(report_data)

        # Display the generated report
        st.header("Generated Report")
        st.download_button(
            label="Download Report",
            data=report_output,
            file_name="sentinel_optimization_assessment.html",
            mime="text/html"
        )
        st.components.v1.html(report_output, height=600, scrolling=True)

if __name__ == "__main__":
    main()