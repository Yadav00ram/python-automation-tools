import os
import pandas as pd
import shutil

# Paths define karen
excel_file = r"C:\Automation Coding Work\Ram Data\All Record.xlsx"  # Excel file ka path
main_folder = r"C:\Jaipur Candidates"  # Data ka main folder
destination_folder = r"C:\Automation Coding Work\Ram Data\Data11"  # Jahan copy karna hai
not_found_excel = r"C:\Automation Coding Work\Ram Data\Data11\Not_Found.xlsx"  # Missing data ki report

# Destination folder agar exist nahi karta, to create karen
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# Missing files report ke liye empty list
not_found = []

# Excel file se mail IDs read karna (blank cells ko handle karna)
df = pd.read_excel(excel_file)
mail_ids = df['Email'].dropna().astype(str).str.lower().tolist()

# Files process karna
for mail_id in mail_ids:
    found = False  # Tracking if file found
    for root, dirs, files in os.walk(main_folder):
        for file in files:
            if file.lower().endswith(".msg") and mail_id in file.lower():
                # File ka path define karen
                source_path = os.path.join(root, file)
                dest_path = os.path.join(destination_folder, file)
                
                # File copy karna
                shutil.copy(source_path, dest_path)
                print(f"Copied: {file} to {destination_folder}")
                found = True
                break  # Agar file mil gayi, to aur search mat karo
        if found:
            break
    
    # Agar file nahi mili
    if not found:
        print(f"Not Found: {mail_id}")
        not_found.append(mail_id)

# Missing data ki Excel report banana
if not_found:
    not_found_df = pd.DataFrame({'Not Found Emails': not_found})
    os.makedirs(os.path.dirname(not_found_excel), exist_ok=True)  # Folder ensure karna
    not_found_df.to_excel(not_found_excel, index=False)
    print(f"Not found report saved at {not_found_excel}")

print("Process complete!")
