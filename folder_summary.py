import os

def print_folder_summary(folder_path):
    # Check if the folder path exists
    if not os.path.exists(folder_path):
        print("Invalid folder path. Please try again.")
        return

    # Print the main folder name
    print(f"Folder: {os.path.basename(folder_path)}")
    
    # Iterate through subfolders
    for subfolder in os.listdir(folder_path):
        subfolder_path = os.path.join(folder_path, subfolder)
        if os.path.isdir(subfolder_path):  # Check if it's a subfolder
            file_count = len([file for file in os.listdir(subfolder_path) if os.path.isfile(os.path.join(subfolder_path, file))])
            # Only print details if the subfolder has 1 or more files
            if file_count > 0:
                print(f"  Subfolder {subfolder} - Files {file_count}")

def main():
    # Hardcoded folder path
    folder_path = r"C:\Users\Ramji\Desktop\Segregation Data\For Upload\PCG"
    print("\nProcessing...\n")
    print_folder_summary(folder_path)

if __name__ == "__main__":
    main()
