"""
Google Drive からのファイル取得ユーティリティ。
サービスアカウントの認証情報は環境変数 GCP_SERVICE_ACCOUNT_JSON (JSON文字列そのもの) から読む。
"""
import io
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


def get_drive_service():
    key_json = os.environ['GCP_SERVICE_ACCOUNT_JSON']
    info = json.loads(key_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)


def list_files_in_folder(service, folder_id):
    """フォルダ直下のファイル一覧 [{id, name, modifiedTime, size}, ...] を返す（trashed除く）。"""
    files = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"
    while True:
        resp = service.files().list(
            q=query,
            fields='nextPageToken, files(id, name, modifiedTime, size, mimeType)',
            pageSize=1000,
            pageToken=page_token,
        ).execute()
        files.extend(resp.get('files', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return files


def download_file(service, file_id, dest_path):
    request = service.files().get_media(fileId=file_id)
    os.makedirs(os.path.dirname(dest_path) or '.', exist_ok=True)
    with io.FileIO(dest_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()


def download_by_name(service, files, name, dest_dir):
    """files一覧(list_files_in_folderの戻り値)からファイル名が一致するものをダウンロード。"""
    match = next((f for f in files if f['name'] == name), None)
    if match is None:
        raise FileNotFoundError(f'Google Drive上に見つかりません: {name}')
    dest_path = os.path.join(dest_dir, name)
    download_file(service, match['id'], dest_path)
    return dest_path
