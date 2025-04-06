import pandas as pd
import requests
import os

# Input Excel file path
input_file = r"C:\Users\user\OneDrive\Documents\Deepaklal\resumes_data.xlsx.xlsx"
output_file = r"C:\Users\user\OneDrive\Documents\Deepaklal\invalid_links.xlsx"
download_folder = r"C:\Users\user\OneDrive\Documents\Deepaklal\resumes"

# Ensure the download folder exists
os.makedirs(download_folder, exist_ok=True)

# Read the Excel file
df = pd.read_excel(input_file)

# List to store invalid links
invalid_links = []

for index, row in df.iterrows():
    name = row['Name']
    email = row['Mail ID']
    resume_url = row['Resume Link']
    
    if pd.isna(resume_url) or not isinstance(resume_url, str):
        invalid_links.append({'Mail ID': email, 'Error': 'No Link Provided'})
        print(f"{email}: No link provided, skipping.")
        continue
    
    try:
        # Download the resume
        response = requests.get(resume_url, timeout=10)
        
        if response.status_code == 200:
            file_extension = resume_url.split('.')[-1]  # Extract file extension
            file_name = os.path.join(download_folder, f"{email}.{file_extension}")
            with open(file_name, 'wb') as file:
                file.write(response.content)
            print(f"{email}: Resume downloaded successfully.")
        else:
            invalid_links.append({'Mail ID': email, 'Error': f"Status Code: {response.status_code}"})
            print(f"{email}: Failed to download (Status Code: {response.status_code}).")
    except requests.RequestException as e:
        invalid_links.append({'Mail ID': email, 'Error': str(e)})
        print(f"{email}: Failed to download ({e}).")

# Save invalid links in a new Excel file
if invalid_links:
    pd.DataFrame(invalid_links).to_excel(output_file, index=False)

print("Process Completed! Resumes downloaded and invalid links logged.")
