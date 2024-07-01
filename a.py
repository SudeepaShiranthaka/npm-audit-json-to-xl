import json
import pandas as pd

def json_to_excel(json_file_path, excel_file_path):
    # Load the JSON data from the file
    with open(json_file_path, 'r') as file:
        json_data = json.load(file)
    
    # Extracting the relevant information
    data = []
    for vulnerability in json_data['vulnerabilities'].values():
        entry = {
            "name": vulnerability['name'],
            "severity": vulnerability['severity'],
            "isDirect": vulnerability['isDirect'],
            "via": ', '.join(v['name'] if isinstance(v, dict) else v for v in vulnerability['via']),
            "effects": ', '.join(vulnerability['effects']),
            "range": vulnerability['range'],
            "nodes": ', '.join(vulnerability['nodes']),
            "fixAvailable": vulnerability['fixAvailable'] if isinstance(vulnerability['fixAvailable'], bool) else f"{vulnerability['fixAvailable']['name']}@{vulnerability['fixAvailable']['version']}"
        }
        data.append(entry)

    # Creating a DataFrame
    df = pd.DataFrame(data)

    # Saving the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

    print(f"Excel file has been created: {excel_file_path}")

# Specify the input JSON file and output Excel file paths
json_file_path = 'vulnerabilities.json'
excel_file_path = 'vulnerabilities.xlsx'

# Convert JSON to Excel
json_to_excel(json_file_path, excel_file_path)
