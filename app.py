import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os


    
# Analysis functions 
def analyze_m365_secure_score(df):
    print("Column names:", df.columns)
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
    top_users_with_prominent_detection['Most Prominent Detection'] = top_users_with_prominent_detection.apply(lambda row: at_risk_users[at_risk_users['UPN'] == row['UPN']]['Detection type'].mode().iloc[0] if not at_risk_users[at_risk_users['UPN'] == row['UPN']].empty else 'N/A', axis=1)
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

import pandas as pd

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
            if upn in high_risk_users['UPN'].values:
                user_risk_level = high_risk_users.loc[high_risk_users['UPN'] == upn, 'Risk level'].values[0]
                user_incident_correlation = pd.concat([user_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'User/UPN': [upn], 'User Risk Level': [user_risk_level]})], ignore_index=True)
        for user in incident_users:
            if user in high_risk_users['User'].values:
                user_risk_level = high_risk_users.loc[high_risk_users['User'] == user, 'Risk level'].values[0]
                user_incident_correlation = pd.concat([user_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'User/UPN': [user], 'User Risk Level': [user_risk_level]})], ignore_index=True)

    device_incident_correlation = pd.DataFrame(columns=['Incident Name', 'Device Name', 'Device Risk Level'])
    for _, incident_row in high_severity_incidents.iterrows():
        incident_name = incident_row['Incident name']
        incident_upns = incident_row['UPNs']
        incident_users = incident_row['Users']
        for upn in incident_upns:
            if upn in high_risk_devices['Device Name'].values:
                device_risk_level = high_risk_devices.loc[high_risk_devices['Device Name'] == upn, 'Risk Level'].values[0]
                device_incident_correlation = pd.concat([device_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'Device Name': [upn], 'Device Risk Level': [device_risk_level]})], ignore_index=True)
        for user in incident_users:
            if user in high_risk_devices['Device Name'].values:
                device_risk_level = high_risk_devices.loc[high_risk_devices['Device Name'] == user, 'Risk Level'].values[0]
                device_incident_correlation = pd.concat([device_incident_correlation, pd.DataFrame({'Incident Name': [incident_name], 'Device Name': [user], 'Device Risk Level': [device_risk_level]})], ignore_index=True)

    user_device_matches = (
        high_risk_users
        .merge(high_risk_devices, left_on='UPN', right_on='Device Name', how='inner')
        [['User', 'UPN', 'Device Name']]
    )

    incident_severity_counts = df['Severity'].value_counts()

    return {
        "incident_summary": incident_summary,
        "incident_severity_counts": incident_severity_counts,
        "user_incident_correlation": user_incident_correlation,
        "device_incident_correlation": device_incident_correlation,
        "user_device_matches": user_device_matches
    }
# Function to load and process CSV data
def process_csv_data(uploaded_files):
    required_files = [
        "Microsoft Secure Score - Microsoft Defender.csv",
        "export-tvm-security-recommendations",
        "User detections.csv",
        "Devices.csv",
        "incidents-queue-"
    ]

    uploaded_file_names = [file.name for file in uploaded_files]

    for required_file in required_files:
        if required_file not in uploaded_file_names and not any(file_name.startswith(required_file) for file_name in uploaded_file_names):
            raise FileNotFoundError(f"Required CSV file not found: {required_file}")

    secure_score_data = pd.read_csv(next((file for file in uploaded_files if file.name == "Microsoft Secure Score - Microsoft Defender.csv"), None))
    tvm_data = pd.read_csv(next((file for file in uploaded_files if file.name.startswith("export-tvm-security-recommendations")), None), skiprows=1)
    user_risk_data = pd.read_csv(next((file for file in uploaded_files if file.name == "User detections.csv"), None))
    device_risk_data = pd.read_csv(next((file for file in uploaded_files if file.name == "Devices.csv"), None))
    incident_data = pd.read_csv(next((file for file in uploaded_files if file.name.startswith("incidents-queue-")), None))

    incident_analysis = analyze_m365_incidents(incident_data, user_risk_data, device_risk_data)

    data = {
        "secure_score": analyze_m365_secure_score(secure_score_data),
        "tvm": analyze_tvm_data(tvm_data),
        "user_risk": analyze_user_risk(user_risk_data, incident_analysis["user_incident_correlation"]),
        "device_risk": analyze_devices_risk(device_risk_data),
        "incidents": incident_analysis,
        "product_deployment": analyze_product_deployment(secure_score_data),
    }

    return data

# Streamlit App Structure 
def main():
    st.title("M365 Threat Gap Analysis and Copilot Readiness Report Generator")
    st.write("Upload your CSV files to get started.")

    uploaded_files = st.file_uploader("Choose CSV files", type="csv", accept_multiple_files=True)

    if uploaded_files:
        try:
            csv_data = process_csv_data(uploaded_files)
            render_report(csv_data)
        except FileNotFoundError as e:
            st.error(str(e))

# Rendering the Report
def render_report(data):
    file_loader = FileSystemLoader('templates')
    env = Environment(loader=file_loader)
    template = env.get_template('TGAreport_template.html')
    output = template.render(data)  

    # Display the rendered HTML on the screen
    st.components.v1.html(output, height=600, scrolling=True)

    # Option to download the rendered HTML:
    st.download_button(label="Download Report", data=output, file_name="M365_Report.html")

if __name__ == "__main__":
    main()