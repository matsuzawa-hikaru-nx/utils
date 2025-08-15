from __future__ import print_function
import os
import re
import io
import time
from tqdm import tqdm
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError  # 追加

# ===== ここを設定 =====
# 保存先ディレクトリ
# DEST = r"C:\Users\hikar\Desktop\検索除外フォルダ\EDR回避-評価用"
DEST = r"./Downloaded"  # 例: ./Downloaded
# DEST = r"./Downloaded"

# ダウンロードするフォルダのURL

# 12:06 - 12:50 開始
# benign files
# https://drive.google.com/drive/folders/1ltQlNnFrmpYCMFhhU492GVRpElD7ZDgc?usp=drive_link

# malware files
# https://drive.google.com/drive/folders/1iDxRSkdG8WHE_8AdegBdCfODUsgpUJss?usp=drive_link

# 13:00 - 13:00
URL = "https://drive.google.com/drive/folders/1OEfYuzHjdMqLtUm5eAGPFCX7fv53xtZB?usp=drive_link"
# ====================

# Google Drive 読み取り専用スコープ
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# OAuth情報の格納ディレクトリ
OAUTH_DIR = './Outh'

# カウンタ
block_count = 0
error_count = 0
success_count = 0


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
    """単一ファイルをダウンロード（危険ファイルも許可）
       成功: True, 失敗: False を返す
    """
    global block_count, error_count
    os.makedirs(save_dir, exist_ok=True)
    file_path = os.path.join(save_dir, file_name)

    try:
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
        return True  # 成功

    except HttpError as e:
        if e.resp.status == 403:
            block_count += 1
            print(f"🚫 Blocked by Zscaler (403) → {file_name}")
        else:
            error_count += 1
            print(f"❌ HTTP Error ({e.resp.status}) for {file_name}")
        return False
    except Exception as e:
        error_count += 1
        print(f"❌ Other error for {file_name}: {e}")
        return False


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
    global success_count
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
    for item in tqdm(all_files, desc="Downloading files", unit="file"):
        if download_file(service, item['id'], item['name'], save_dir):
            success_count += 1
        # ダウンロード間隔を空ける（例: 1秒）
        time.sleep(2)

    # 集計結果を表示
    print("\n=== Download Summary ===")
    print(f"Total files: {len(all_files)}")
    print(f"Successful downloads: {success_count} files")
    print(f"Blocked by Zscaler: {block_count} files")
    print(f"Other errors: {error_count} files")


if __name__ == '__main__':
    main()
