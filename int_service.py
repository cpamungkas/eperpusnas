import openpyxl
from smbprotocol.connection import Connection
from smbprotocol.session import Session
from smbprotocol.tree import TreeConnect
from smbprotocol.open import Open, Directory
from smbprotocol.file_info import FileInformationClass, FileAttributes

# Configuration for SMB connection
server = "xx.xx.xx.xx"
share = "Perpusnas"
username = "your_username"  # Ganti dengan username
password = "your_password"  # Ganti dengan password
folder_path = ""  # Kosongkan atau set sesuai folder di share kalau perlu

# Setup SMB connection
def connect_to_smb(server, share, username, password):
    conn = Connection(uuid="unique-uuid", server_name=server, port=445)
    conn.connect()
    session = Session(conn, username=username, password=password)
    session.connect()
    tree = TreeConnect(session, share)
    tree.connect()
    return tree

# List files in SMB server directory
def list_files_in_smb_directory(tree, folder_path):
    dir_open = Open(tree, folder_path, desired_access=FileAttributes.FILE_LIST_DIRECTORY)
    dir_open.create(FileInformationClass.FILE_DIRECTORY_INFORMATION)
    file_list = dir_open.query_directory()
    dir_open.close()
    
    file_names = []
    for file_info in file_list:
        if not file_info['file_name'].endswith('/'):  # Exclude directories
            file_names.append(file_info['file_name'])
    return file_names

# Download file from SMB server
def download_file_from_smb(tree, file_name):
    file_open = Open(tree, file_name)
    file_open.create(FileInformationClass.FILE_READ_DATA)
    file_data = file_open.read(0, file_open.size)
    file_open.close()
    return file_data

# Save the file locally
def save_file_locally(file_data, local_file_path):
    with open(local_file_path, "wb") as local_file:
        local_file.write(file_data)

# Read and process Excel file
def read_excel_file(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    # Example: Print data from the first column
    for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
        print(row[0])

# Main function to run the process
def main():
    try:
        # Step 1: Connect to SMB server
        tree = connect_to_smb(server, share, username, password)
        print(f"Connected to {server}/{share}")

        # Step 2: List all files in the directory
        files = list_files_in_smb_directory(tree, folder_path)
        print(f"Files in directory: {files}")

        # Step 3: Choose the file you want to download (e.g., based on some condition)
        excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xls')]
        
        if not excel_files:
            print("No Excel files found in the directory.")
            return
        
        # For simplicity, let's take the first Excel file we find
        selected_file = excel_files[0]
        print(f"Selected file: {selected_file}")

        # Step 4: Download the selected Excel file
        file_data = download_file_from_smb(tree, selected_file)
        print(f"Downloaded {selected_file}")

        # Step 5: Save the file locally
        local_file_path = "downloaded_excel_file.xlsx"
        save_file_locally(file_data, local_file_path)
        print(f"Saved file locally as {local_file_path}")

        # Step 6: Read the Excel file
        read_excel_file(local_file_path)
        print("Finished reading Excel file")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
