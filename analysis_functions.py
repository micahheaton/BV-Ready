import pandas as pd
from io import BytesIO

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

def process_m365_files(uploaded_files):
    m365_required_files = [
        "Microsoft Secure Score - Microsoft Defender.csv",
        "export-tvm-security-recommendations",
        "User detections.csv",
        "Devices.csv",
        "incidents-queue-",  # Starts with ...
    ]

    m365_files = [file for file in uploaded_files if file.name in m365_required_files]
    m365_results = {}

    for file in m365_files:
        if file.name == "Microsoft Secure Score - Microsoft Defender.csv":
            secure_score_df = pd.read_csv(file)
            m365_results["secure_score"] = analyze_m365_secure_score(secure_score_df)
        elif file.name == "export-tvm-security-recommendations":
            tvm_data_df = pd.read_csv(file)
            m365_results["product_deployment"] = analyze_product_deployment(tvm_data_df)
            m365_results["tvm"] = analyze_tvm_data(tvm_data_df)
        elif file.name == "User detections.csv":
            user_detections_df = pd.read_csv(file)
            m365_results["user_risk"] = analyze_user_risk(user_detections_df, None)
        elif file.name == "Devices.csv":
            devices_df = pd.read_csv(file)
            m365_results["device_risk"] = analyze_devices_risk(devices_df)
        elif file.name.startswith("incidents-queue-"):
            incidents_df = pd.read_csv(file)
            if "incidents" not in m365_results:
                m365_results["incidents"] = {}
            m365_results["incidents"].update(analyze_m365_incidents(incidents_df, m365_results.get("user_risk", {}), m365_results.get("device_risk", {})))

    return m365_results

def process_sentinel_files(uploaded_files):
    sentinel_results = {}
    sentinel_tq_file = None

    for file in uploaded_files:
        if "TQ" in file.name.lower() and file.name.endswith('.xlsx'):
            sentinel_tq_file = file
            break

    if sentinel_tq_file:
        report_data = load_and_process_data(sentinel_tq_file, {})
        sentinel_results = report_data

    return sentinel_results

def process_defender_cloud_files(uploaded_files):
    defender_cloud_results = {}
    # ... (process Defender for Cloud files)
    return defender_cloud_results

def analyze_m365_secure_score(df):
    def parse_points_achieved(points_achieved):
        if pd.isna(points_achieved) or points_achieved.strip() == '':
            return 0, 0
        try:
            points_achieved = points_achieved.strip().replace('"', '').replace(' ', '')
            numerator, denominator = map(float, points_achieved.split('/'))
            return numerator, denominator
        except (ValueError, ZeroDivisionError):
            return 0, 0

    df['Points Achieved'], df['Points Possible'] = zip(*df['Points achieved'].apply(parse_points_achieved))
    df['Status'] = df.apply(lambda row: 'Completed' if row['Points Achieved'] == row['Points Possible'] else 'To Address', axis=1)
    df['Score Difference'] = df['Points Possible'] - df['Points Achieved']

    category_scores = df.groupby('Category').agg({'Points Achieved': 'sum', 'Points Possible': 'sum'}).reset_index()
    category_scores['Completed Percentage'] = category_scores['Points Achieved'] / category_scores['Points Possible'] * 100
    category_scores['To Address Percentage'] = 100 - category_scores['Completed Percentage']

    category_order = ['Apps', 'Data', 'Identity', 'Device']
    category_scores['Category'] = pd.Categorical(category_scores['Category'], categories=category_order, ordered=True)
    category_scores = category_scores.sort_values('Category')

    overall_completed_percentage = category_scores['Points Achieved'].sum() / category_scores['Points Possible'].sum() * 100
    overall_to_address_percentage = 100 - overall_completed_percentage

    top_actions_to_prioritize = df.nlargest(15, 'Score Difference')[['Recommended action', 'Points Possible', 'Points Achieved', 'Score Difference', 'Category']].values.tolist()

    return {
        "secure_score_completed": round(overall_completed_percentage, 1),
        "secure_score_to_address": round(overall_to_address_percentage, 1),
        "category_scores": category_scores[['Category', 'Completed Percentage', 'To Address Percentage']],
        "top_actions_to_prioritize": top_actions_to_prioritize
    }

def analyze_product_deployment(df):
    product_deployment = df.groupby('Product')[['Points Achieved', 'Points Possible']].sum()
    product_deployment['Deployment Percentage'] = product_deployment['Points Achieved'] / product_deployment['Points Possible'] * 100
    return product_deployment.reset_index()

def analyze_tvm_data(df):
    active_recommendations = df[df['Status'] == 'Active']
    exposure_score_avg = df['Exposure Score impact'].mean()
    top_recommendations = (
        active_recommendations
        .nlargest(10, 'Exposure Score impact')
        [['Security recommendation', 'Exposure Score impact', 'Exposed Machines', 'Total Machines']]
        .reset_index(drop=True)
    )

    return {
        "total_recommendations": len(df),
        "active_recommendations": len(active_recommendations),
        "exposure_score_avg": round(exposure_score_avg, 2),
        "top_recommendations": top_recommendations
    }

def analyze_user_risk(df, user_incident_correlation):
    at_risk_users = df[(df['Risk state'] == 'At risk') & (df['Risk level'] == 'High')]
    
    important_detection_types = [
        'Leaked credentials',
        'Malicious IP address',
        'Impossible travel',
        'Password spray',
        'New country',
        'Suspicious inbox forwarding',
        'Suspicious inbox manipulation rules',
        'Verified threat actor IP',
        'Admin confirmed user compromised',
        'User reported suspicious activity'
    ]
    
    important_detections_df = at_risk_users[at_risk_users['Detection type'].isin(important_detection_types)][['User', 'Detection type']]
    important_detections_df['Documentation URL'] = 'https://example.com/placeholder'  # Placeholder URL
    
    user_incident_count = at_risk_users.groupby('UPN').size().reset_index(name='Incident Count')
    user_risk_details = at_risk_users.merge(user_incident_count, on='UPN')
    
    top_users = user_risk_details.drop_duplicates('UPN').nlargest(10, 'Incident Count')
    
    top_users_with_prominent_detection = top_users.copy()
    top_users_with_prominent_detection['Most Prominent Detection'] = top_users_with_prominent_detection.apply(lambda row: at_risk_users[at_risk_users['UPN'] == row['UPN']]['Detection type'].mode().iloc[0] if not at_risk_users[at_risk_users['UPN'] == row['UPN']]['Detection type'].empty else 'N/A', axis=1)
    top_users_with_prominent_detection['Detection Type URL'] = 'https://example.com/placeholder'  # Placeholder URL
    
    return {
        "important_detections": important_detections_df,
        "top_users_with_prominent_detection": top_users_with_prominent_detection[['User', 'UPN', 'Incident Count', 'Risk level', 'Most Prominent Detection', 'Detection Type URL', 'Location']]
    }

def analyze_devices_risk(df):
    os_summary = df.groupby('OS Platform')['Device ID'].count()
    patch_summary = df.groupby('Domain')['OS Version'].count()
    risk_summary = df.groupby('Risk Level')['Device ID'].count()

    active_devices_count = len(df[df['Health Status'] == 'Active'])
    inactive_devices_count = len(df[df['Health Status'] == 'Inactive'])

    high_risk_devices = df[(df['Risk Level'] == 'High') & (df['Health Status'] == 'Active')]
    top_high_risk_devices = high_risk_devices[['Device Name', 'Risk Level']].head(15)

    high_exposure_devices = df[(df['Risk Level'] == 'High') & (df['Exposure Level'] == 'High')]  # Adjust the threshold as needed
    high_exposure_devices_summary = high_exposure_devices[['Device Name', 'Risk Level', 'Exposure Level']].head(10)
    
    return {
        "os_summary": os_summary,
        "patch_summary": patch_summary,
        "risk_summary": risk_summary,
        "active_devices_count": active_devices_count,
        "inactive_devices_count": inactive_devices_count,
        "high_risk_devices": len(high_risk_devices),
        "top_high_risk_devices": top_high_risk_devices,
        "high_exposure_devices_summary": high_exposure_devices_summary
    }

def analyze_m365_incidents(df, user_risk_data, device_risk_data):
    incident_summary = (
        df[['Incident name', 'Severity', 'Impacted assets', 'Tags']]
        .reset_index(drop=True)
        .head(10)
    )

    high_severity_incidents = df[df['Severity'] == 'high']
    high_risk_users = user_risk_data[user_risk_data['Risk level'] == 'High']
    high_risk_devices = device_risk_data[device_risk_data['Risk Level'] == 'High']

    def extract_upns_and_users(impacted_assets):
        upns = []
        users = []
        for asset in impacted_assets.split(','):
            if 'Accounts:' in asset:
                accounts = asset.split('Accounts:')[1].strip()
                upns.extend([upn.strip() for upn in accounts.split(',')])
                users.extend([user.strip() for user in accounts.split(',')])
            elif 'Mailboxes:' in asset:
                mailboxes = asset.split('Mailboxes:')[1].strip()
                upns.extend([upn.strip() for upn in mailboxes.split(',')])
                users.extend([user.strip() for user in mailboxes.split(',')])
        return upns, users

    high_severity_incidents['UPNs'], high_severity_incidents['Users'] = zip(*high_severity_incidents['Impacted assets'].apply(extract_upns_and_users))

    user_incident_correlation = pd.DataFrame(columns=['Incident Name', 'User/UPN', 'User Risk Level'])
    for _, incident_row in high_severity_incidents.iterrows():
        incident_name = incident_row['Incident name']
        incident_upns = incident_row['UPNs']
        incident_users = incident_row['Users']
        for upn in incident_upns:
            if not high_risk_users.empty and upn in high_risk_users['UPN'].values:
                user_risk_level = high_risk_users.loc[high_risk_users['UPN'] == upn, 'Risk level'].values[0]
                user_incident_correlation = pd.concat([user_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'User/UPN': [upn], 'User Risk Level': [user_risk_level]})], ignore_index=True)
        for user in incident_users:
            if not high_risk_users.empty and user in high_risk_users['User'].values:
                user_risk_level = high_risk_users.loc[high_risk_users['User'] == user, 'Risk level'].values[0]
                user_incident_correlation = pd.concat([user_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'User/UPN': [user], 'User Risk Level': [user_risk_level]})], ignore_index=True)

    device_incident_correlation = pd.DataFrame(columns=['Incident Name', 'Device Name', 'Device Risk Level'])
    for _, incident_row in high_severity_incidents.iterrows():
        incident_name = incident_row['Incident name']
        incident_upns = incident_row['UPNs']
        incident_users = incident_row['Users']
        for upn in incident_upns:
            if not high_risk_devices.empty and upn in high_risk_devices['Device Name'].values:
                device_risk_level = high_risk_devices.loc[high_risk_devices['Device Name'] == upn, 'Risk Level'].values[0]
                device_incident_correlation = pd.concat([device_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'Device Name': [upn], 'Device Risk Level': [device_risk_level]})], ignore_index=True)
        for user in incident_users:
            if not high_risk_devices.empty and user in high_risk_devices['Device Name'].values:
                device_risk_level = high_risk_devices.loc[high_risk_devices['Device Name'] == user, 'Risk Level'].values[0]
                device_incident_correlation = pd.concat([device_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'Device Name': [user], 'Device Risk Level': [device_risk_level]})], ignore_index=True)

    user_device_matches = (
        high_risk_users
        .merge(high_risk_devices, left_on='UPN', right_on='Device Name', how='inner')
        [['User', 'UPN', 'Device Name']]
    )

    incident_severity_counts = df['Severity'].value_counts().to_dict()

    return {
        "incident_summary": incident_summary,
        "incident_severity_counts": incident_severity_counts,
        "user_incident_correlation": user_incident_correlation,
        "device_incident_correlation": device_incident_correlation,
        "user_device_matches": user_device_matches
    }