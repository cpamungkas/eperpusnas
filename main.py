from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.responses import FileResponse
from typing import List, Dict, Optional, Tuple
from fastapi.middleware.cors import CORSMiddleware
import os
import pandas as pd
import openpyxl
import win32wnet
import requests
import json
import urllib.parse

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Atau tambahin URL frontend lo
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize Basic Auth
security = HTTPBasic()


# Define usernae and password
VALID_USERNAME = "username"
VALID_PASSWORD = "password"


# Function handle auth check
def authenticate_user(credentials: HTTPBasicCredentials = Depends(security)):
    # Cek User Password
    correct_username = credentials.username == VALID_USERNAME
    correct_password = credentials.password == VALID_PASSWORD

    if not (correct_username and correct_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Basic"},
        )


# Hardcode nilai variabel global
host = "host domain"
username = "username"
password = "password"
share_path = "share path"
central_server_api_url = "url api server"

# Authorization Headers
auth_headers = {
    "Client-ID": "xc0000000000000000000",
    "Client-Secret": "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz",
    "Authable-ID": "XX0000000000000000000",
    "Content-Type": "application/json"  # Pastikan content type benar
}


def wnet_connect(host: str, username: str, password: str):
    unc = f"\\\\{host}"
    try:
        win32wnet.WNetAddConnection2(0, None, unc, None, username, password)
    except win32wnet.error as err:
        if err.winerror == 1219:
            # Handle multiple connections to the same resource
            win32wnet.WNetCancelConnection2(unc, 0, 0)
            return wnet_connect(host, username, password)
        raise HTTPException(
            status_code=500, detail=f"Failed to connect to network: {str(err)}")


def list_subfolders_and_files(host: str, username: str, password: str, share_path: str) -> Optional[List[str]]:
    try:
        wnet_connect(host, username, password)
        path = f"\\\\{host}\\{share_path.replace(':', '$')}"
        subfolders = []

        if os.path.exists(path):
            for dirpath, _, filenames in os.walk(path):
                subfolders.append((dirpath, filenames))
            return subfolders
        else:
            raise HTTPException(
                status_code=404, detail=f"Path {path} does not exist or is inaccessible.")
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error accessing network path: {str(e)}")


# Fungsi untuk list file difolder tertentu
def list_subfolders(host: str, username: str, password: str, share_path: str, subfolder: str) -> Optional[List[Tuple[str, List[str]]]]:
    try:
        wnet_connect(host, username, password)
        path = f"\\\\{host}\\{share_path.replace(':', '$')}"
        subfolders = []

        if os.path.exists(path):
            for dirpath, _, filenames in os.walk(path):
                if subfolder in dirpath:
                    subfolders.append((dirpath, filenames))
            return subfolders
        else:
            raise HTTPException(
                status_code=404, detail=f"Path {path} does not exist or is inaccessible.")
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error accessing network path: {str(e)}")


# Fungsi untuk list file per subfolder
def list_files_folder(host: str, username: str, password: str, share_path: str) -> Optional[Dict[str, List[Dict[str, str]]]]:
    try:
        wnet_connect(host, username, password)
        path = f"\\\\{host}\\{share_path.replace(':', '$')}"
        if not os.path.exists(path):
            raise HTTPException(status_code=404, detail="Path not found")

        files = {}
        for root, dirs, filenames in os.walk(path):
            subfolder = os.path.relpath(root, path)
            files[subfolder] = {
                "images": [f for f in filenames if f.lower().endswith(('.jpg', '.jpeg', '.png'))],
                "pdfs": [f for f in filenames if f.lower().endswith('.pdf')],
                "excels": [f for f in filenames if f.lower().endswith(('.xls', '.xlsx'))]
            }
        return files
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error accessing network path: {str(e)}")


# Function untuk menampilkan file image & pdf dalam folder
def list_data_subfolders_and_files(host: str, username: str, password: str, share_path: str) -> Optional[List[Dict]]:
    try:
        # Connect to the network share using win32wnet or similar
        wnet_connect(host, username, password)
        path = f"\\\\{host}\\{share_path.replace(':', '$')}"

        # Prepare to store the list of directories and their files
        subfolders = []

        # Check if the path exists and is accessible
        if os.path.exists(path):
            # Walk through the directory and list subfolders and files
            for dirpath, _, filenames in os.walk(path):
                # Filter filenames to only include image and PDF files
                filtered_files = [file for file in filenames if file.lower().endswith(
                    ('.jpg', '.jpeg', '.png', '.pdf'))]

                # If there are filtered files, append the directory and its files
                if filtered_files:
                    subfolders.append({
                        "dirpath": dirpath,
                        "filenames": filtered_files
                    })
            return subfolders
        else:
            raise HTTPException(
                status_code=404, detail=f"Path {path} does not exist or is inaccessible.")
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error accessing network path: {str(e)}")


# Fungsi untuk generate json payload
def create_json_payload(row_data) -> dict:
    payload = {
        "code": row_data.get('code', ''),
        "type": row_data.get('type', ''),
        "preview": row_data.get('preview', ''),
        "description": row_data.get('description', ''),
        "publication_year": row_data.get('publication_year', ''),
        "publication_month": row_data.get('publication_month', ''),
        "title": row_data.get('title', ''),
        "category": parse_categories(row_data('category')),
        "contributor": parse_contributors(row_data('contributor')),
        "file_original": row_data.get('file_original', ''),
        "file_cover": row_data.get('file_cover', ''),
        "access": row_data.get('access', ''),
        "price": row_data.get('price', '')
    }
    print("Payload yang akan dikirim:", payload)
    return payload


def custom_json_serializer(obj):
    if isinstance(obj, float):
        if obj != obj or obj == float('inf') or obj == float('-inf'):
            return str(obj)
    raise TypeError("Type not serializable")


# Fungsi untuk membaca file excel
def read_excel_file(file_path: str) -> Optional[List[dict]]:
    if os.path.exists(file_path) and file_path.endswith(('.xls', '.xlsx')):
        try:
            df = pd.read_excel(file_path)
            return df.to_dict(orient='records')
        except Exception as e:
            raise HTTPException(
                status_code=500, detail=f"Failed to read Excel file: {str(e)}")
    else:
        raise HTTPException(
            status_code=404, detail=f"Excel file not found or invalid format: {file_path}")


# Fungsi untuk mengupload data ke central server
def upload_to_central_server(data: dict):
    try:
        response = requests.post(
            central_server_api_url, json=data, headers=auth_headers)
        print("Response from server:", response.text)  # Debug: print response
        response.raise_for_status()
        return response.status_code
    except requests.RequestException as e:
        print(f"RequestException: {str(e)}")  # Debug: print error detail
        raise HTTPException(
            status_code=500, detail=f"Failed to upload data: {str(e)}")


# Fungsi untuk memindahkan file dari central server
def move_file_to_archive(file_path: str):
    archive_folder = os.path.join(os.path.dirname(file_path), "archive")
    os.makedirs(archive_folder, exist_ok=True)

    file_name = os.path.basename(file_path)
    new_path = os.path.join(archive_folder, file_name)

    os.rename(file_path, new_path)


# Fungsi untuk memecah data kontributor
def parse_contributors(contributor_string):
    contributors = []
    items = contributor_string.split(';')
    for item in items:
        item = item.strip()

        # Cek role berdasarkan kata kunci yang umum
        if 'penyunting' in item.lower():
            role = 'Penyunting'
            name = item.lower().replace('penyunting,', '', 1).strip()
        elif 'pengarang' in item.lower() or 'author' in item.lower():
            role = 'Pengarang'
            name = item.lower().replace('pengarang,', '', 1).strip()
        elif 'editor' in item.lower():
            role = 'Editor'
            name = item.lower().replace('editor,', '', 1).strip()
        elif 'ilustrasi' in item.lower():
            role = 'Ilustrasi'
            name = item.lower().replace('ilustrasi,', '', 1).strip()
        else:
            # Default ke pengarang jika role tidak jelas
            role = 'Pengarang'
            name = item.strip()

        # Format nama jadi huruf besar diawal tiap kata
        name = name.title()

        # Buat dict contributor
        contributor = {
            "name": role.capitalize(),
            "author_fullname": name,
            "author_title": "",  # Set default title jika ada di masa depan
            # Default tahun lahir, nanti bisa diganti kalau ada data lebih lanjut
            "author_year_of_birth": "2000",
            "author_year_of_death": ""  # Set default kosong
        }
        contributors.append(contributor)

    return contributors


# Fungsi untuk memecah data category
def parse_categories(category_string):
    return [category.strip() for category in category_string.split(',')]


# Endpoint untuk mendapatkan data file dalam folder
@app.get("/files")
async def get_files():
    return list_files_folder(host, username, password, share_path)


# EndPoint untuk mendapatkan data file dalam folder tertentu
@app.get("/file/{subfolder}/{filename}")
async def get_file(subfolder: str, filename: str):
    file_path = os.path.join(
        f"\\\\{host}\\{share_path.replace(':', '$')}", subfolder, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    return {"file_url": file_path}


# Endpoint untuk mendapatkan seluruh folder & file
@app.get("/list_files")
async def get_files():
    files = list_subfolders_and_files(host, username, password, share_path)
    if not files:
        raise HTTPException(
            status_code=404, detail=f"Path {share_path} does not exist")
    return {"data files": files}


@app.get("/get_files")
async def get_files():
    # Call the function to list files and subfolders
    files = list_data_subfolders_and_files(
        host, username, password, share_path)
    if not files:
        raise HTTPException(
            status_code=404, detail=f"Path {share_path} does not exist")
    # Return the files in the desired format
    return {"data files": files}


# Endpoint untuk membaca isi file excel di setiap folder
@app.get("/read_excel")
async def read_excel():
    subfolders = list_subfolders_and_files(
        host, username, password, share_path)
    if not subfolders:
        raise HTTPException(
            status_code=404, detail=f"Path {share_path} does not exist")

    all_data = []
    for dirpath, filenames in subfolders:
        for filename in filenames:
            if filename.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(dirpath, filename)
                try:
                    excel_data = read_excel_file(file_path)
                    for record in excel_data:
                        if 'contributor' in record:
                            record['contributor'] = parse_contributors(
                                record['contributor'])
                        if 'category' in record:
                            record['category'] = parse_categories(
                                record['category'])
                    all_data.append({
                        "path": dirpath,
                        "data_excel": excel_data
                    })
                    # all_data.extend(excel_data)
                except HTTPException as e:
                    # Jika ada error di file tertentu, lanjut ke file berikutnya
                    print(f"Error reading {file_path}: {e.detail}")

    return {"data": all_data}
    # return json.dumps({"data":all_data}, default=custom_json_serializer)


# Endpoint untuk membaca isi file excel dan memproses data ke API server
@app.post("/process_excel_files")
async def process_excel_files():
    subfolders = list_subfolders_and_files(
        host, username, password, share_path)
    if not subfolders:
        raise HTTPException(
            status_code=404, detail=f"Path {share_path} does not exist")

    errors = []  # Simpan error di list ini
    for dirpath, filenames in subfolders:
        for filename in filenames:
            if filename.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(dirpath, filename)
                try:
                    excel_data = read_excel_file(file_path)

                    for row_data in excel_data:
                        payload = create_json_payload(row_data)
                        try:
                            status_code = upload_to_central_server(payload)
                            if status_code != 200:
                                errors.append(
                                    f"Failed to upload data for {filename}, status code: {status_code}")
                        except HTTPException as e:
                            errors.append(
                                f"Exception during upload for {filename}: {e.detail}")

                    # Only move to archive if all rows for a file were processed successfully
                    if not errors:
                        move_file_to_archive(file_path)
                    else:
                        print(
                            f"Errors found, file {filename} will not be moved to archive.")

                except HTTPException as e:
                    errors.append(f"Error processing {file_path}: {e.detail}")

    if errors:
        # Return all errors found during processing
        raise HTTPException(
            status_code=500, detail={"message": "Some files failed to process", "errors": errors}
        )

    return {"detail": "All Excel files processed and moved to archive."}


# Endpoints for HIT BIP Data
@app.post("/process_data_bip")
async def process_excel_files():
    # subfolders = list_subfolders_and_files(host, username, password, share_path)
    subfolders = list_subfolders(host, username, password, share_path, 'BIP')
    if not subfolders:
        raise HTTPException(
            status_code=404, detail=f"Path {share_path} does not exist")

    errors = []
    all_data_bip = []
    for dirpath, filenames in subfolders:
        for filename in filenames:
            if filename.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(dirpath, filename)
                try:
                    # df = pd.read_excel(file_path)
                    # data = df.to_dict(orient='records')
                    excel_data = read_excel_file(file_path)

                    for row_data in excel_data:
                        payload = {
                            "code": row_data.get('code', ''),
                            "type": row_data.get('type', ''),
                            "preview": row_data.get('preview', ''),
                            "description": row_data.get('description', ''),
                            "publication_year": row_data.get('publication_year', ''),
                            "publication_month": row_data.get('publication_month', ''),
                            "title": row_data.get('title', ''),
                            "category": parse_categories(row_data.get('category', '')),
                            "contributor": parse_contributors(row_data['contributor']),
                            "file_original": row_data.get('file_original', ''),
                            "file_cover": row_data.get('file_cover', ''),
                            "price": row_data.get('price', ''),
                            "access": row_data.get('access', '')
                        }
                        print(f"Data berhasil dikirim: {payload}")
                        # response = requests.post(
                        #     central_server_api_url, json=payload, headers=auth_headers)
                        # if response.status_code != 200:
                        #     errors.append(
                        #         f"Failed to upload data for {filename}, status code: {response.status_code}")

                    all_data_bip.append(payload)
                    # if not errors:
                    #     archive_folder = os.path.join(
                    #         os.path.dirname(file_path), "archive")
                    #     os.makedirs(archive_folder, exist_ok=True)
                    #     os.rename(file_path, os.path.join(
                    #         archive_folder, os.path.basename(file_path)))
                except Exception as e:
                    errors.append(f"Error processing {file_path}: {e}")

    if errors:
        raise HTTPException(
            status_code=500, detail={"message": "Some files failed to process", "errors": errors}
        )

    return {"detail": "All Excel files processed and moved to archive.", "path": dirpath}
    # return {"{": all_data_bip}


# Endpoint untuk menampilkan preview file
@app.get("/api/preview")
# def preview_file(file_path: str):
def preview_file(file_path: str, credentials: HTTPBasicCredentials = Depends(authenticate_user)):
    # Decode URL encoded file path
    decoded_path = urllib.parse.unquote(file_path)

    # Cek apakah file path valid dan eksis
    if os.path.exists(decoded_path):
        return FileResponse(decoded_path)
    else:
        raise HTTPException(status_code=404, detail="File not found")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
