from __future__ import print_function
import os
import re
import io
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ===== ここを設定 =====
# 保存先ディレクトリ
# DEST = r"C:\Users\hikar\Desktop\検索除外フォルダ\EDR回避-評価用"
DEST = r"./Downloaded"
# ダウンロードするフォルダのURL
URL = "https://drive.google.com/drive/folders/1PFIo-uQmsIJzmvlBVXIEJxQrxVB3l9e_?usp=drive_link"
# ====================

# Google Drive 読み取り専用スコープ
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# OAuth情報の格納ディレクトリ
OAUTH_DIR = './Outh'

def get_service():
    """Google Drive APIの認証とサービスオブジェクト作成"""
    creds = None
    token_path = os.path.join(OAUTH_DIR, 'token.json')
    credentials_path = os.path.join(OAUTH_DIR, 'credentials.json')

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

def download_file(service, file_id, file_name, save_dir):
    """単一ファイルをダウンロード（危険ファイルも許可）"""
    os.makedirs(save_dir, exist_ok=True)
    file_path = os.path.join(save_dir, file_name)

    # acknowledgeAbuse=True でマルウェア警告を回避
    request = service.files().get_media(fileId=file_id, acknowledgeAbuse=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"{file_name}: {int(status.progress() * 100)}%")

    with open(file_path, 'wb') as f:
        f.write(fh.getvalue())
    print(f"✅ {file_name} is saved at {file_path}")

import re

def extract_folder_id(url_or_id: str) -> str:
    """
    Google DriveフォルダのURLまたは直接IDを渡すとフォルダIDを返す
    例:
      https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz
      または
      1AbCdEfGhIjKlMnOpQrStUvWxYz
    """
    # URLパターンにマッチしたらその部分を抽出
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', url_or_id)
    if match:
        return match.group(1)
    # URL形式じゃなければそのままIDとして返す
    return url_or_id

def main():
    # ===== ここを設定 =====
    # 保存先ディレクトリ
    dest = DEST
    # ダウンロードするフォルダのURL
    url = URL
    # ====================

    # URLからフォルダIDを抽出
    folder_id = extract_folder_id(url)
    service = get_service()

    # フォルダ名取得
    folder_meta = service.files().get(fileId=folder_id, fields="name").execute()
    folder_name = folder_meta.get('name', 'Downloaded_Folder')

    # 保存先ディレクトリ（例: ./Downloaded/<フォルダ名>/）
    save_dir = os.path.join(dest, folder_name)

    # ページネーションでフォルダ内すべてのファイル取得
    query = f"'{folder_id}' in parents and trashed=false"
    page_token = None
    all_files = []

    while True:
        results = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name)",
            pageSize=1000,
            pageToken=page_token
        ).execute()

        items = results.get('files', [])
        all_files.extend(items)

        page_token = results.get('nextPageToken')
        if not page_token:
            break

    print(f'📂 Number of files in "{folder_name}": {len(all_files)}')

    # ファイルを順次ダウンロード
    for item in all_files:
        download_file(service, item['id'], item['name'], save_dir)

if __name__ == '__main__':
    main()
