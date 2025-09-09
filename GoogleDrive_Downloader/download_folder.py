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

# ダウンロードするフォルダのURL
URL = "https://drive.google.com/drive/folders/1Nto9P2BJ-J9jT-Z9nwaliPfukfT3UgJB?usp=drive_link"
# ====================

# Google Drive 読み取り専用スコープ
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# OAuth情報の格納ディレクトリ
OAUTH_DIR = './Outh'

# カウンタ
block_count = 0
error_count = 0
success_count = 0

# MIME 定義
FOLDER_MIME   = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"

# Google ドキュメント形式 → 保存形式（必要に応じて変更可）
EXPORT_MAP = {
    "application/vnd.google-apps.document": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    "application/vnd.google-apps.spreadsheet": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "application/vnd.google-apps.presentation": ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
    "application/vnd.google-apps.drawing": ("image/png", ".png"),
    # PDF が良ければ上記の MIME/拡張子を pdf に変更してもよい
}

def sanitize(name: str) -> str:
    """Windows を含む OS 非対応文字を置換し、末尾の . や空白を除去"""
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = name.rstrip(' .')
    return name or "_"

def unique_path(path: str) -> str:
    """同名ファイル/フォルダが存在する場合に (1), (2), ... を付けて衝突回避"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    k = 1
    while True:
        cand = f"{base} ({k}){ext}"
        if not os.path.exists(cand):
            return cand
        k += 1

def get_service():
    """Google Drive APIの認証とサービスオブジェクト作成"""
    creds = None
    os.makedirs(OAUTH_DIR, exist_ok=True)
    token_path = os.path.join(OAUTH_DIR, 'token.json')
    credentials_path = os.path.join(OAUTH_DIR, 'credentials.json')

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            # 必要に応じて prompt='consent' なども可
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

def resolve_shortcut(service, item):
    """ショートカットなら実体に解決して返す（名前はショートカット名を優先）"""
    if item.get("mimeType") != SHORTCUT_MIME:
        return item
    meta = service.files().get(
        fileId=item["id"],
        fields="id,name,shortcutDetails/targetId,shortcutDetails/targetMimeType"
    ).execute()
    tid  = meta["shortcutDetails"]["targetId"]
    real = service.files().get(fileId=tid, fields="id,name,mimeType").execute()
    real["name"] = meta.get("name", real["name"])
    return real

def download_drive_item(service, item, save_dir) -> bool:
    """
    フォルダ以外の1アイテムを save_dir に保存。
    Google ドキュメント形式は export、通常ファイルは get_media。
    ショートカットは実体へ解決。
    成功: True / 失敗: False
    """
    global block_count, error_count, success_count
    os.makedirs(save_dir, exist_ok=True)

    # ショートカット解決
    item = resolve_shortcut(service, item)

    file_id = item["id"]
    file_name = sanitize(item["name"])
    mime = item.get("mimeType", "")

    try:
        # Google ドキュメント形式は export で保存
        if mime in EXPORT_MAP:
            export_mime, ext = EXPORT_MAP[mime]
            request = service.files().export_media(fileId=file_id, mimeType=export_mime)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            path = unique_path(os.path.join(save_dir, file_name + ext))
            with open(path, "wb") as f:
                f.write(fh.getvalue())
            success_count += 1
            return True

        # 通常のバイナリは get_media
        request = service.files().get_media(fileId=file_id, acknowledgeAbuse=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        path = unique_path(os.path.join(save_dir, file_name))
        with open(path, "wb") as f:
            f.write(fh.getvalue())
        success_count += 1
        return True

    except HttpError as e:
        if getattr(e, "resp", None) and e.resp.status == 403:
            block_count += 1
            print(f"🚫 403: {file_name}")
        else:
            error_count += 1
            code = getattr(getattr(e, "resp", None), "status", "?")
            print(f"❌ HTTP Error ({code}) for {file_name}: {e}")
        return False
    except Exception as e:
        error_count += 1
        print(f"❌ Other error for {file_name}: {e}")
        return False

def download_folder_recursive(service, folder_id, save_dir):
    """
    フォルダID配下を再帰走査し、Drive の階層を save_dir 配下に完全再現して保存。
    """
    os.makedirs(save_dir, exist_ok=True)

    page_token = None
    while True:
        res = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=page_token
        ).execute()

        items = res.get("files", [])
        # 進捗表示（任意）：フォルダ内のアイテム数が多い場合は tqdm で見やすく
        for it in tqdm(items, desc=f'In "{os.path.basename(save_dir)}"', unit="item"):
            name = sanitize(it["name"])
            mime = it.get("mimeType", "")
            if mime == FOLDER_MIME:
                # 階層を再現：サブフォルダ名を足して再帰
                child_dir = unique_path(os.path.join(save_dir, name))
                download_folder_recursive(service, it["id"], child_dir)
            else:
                download_drive_item(service, it, save_dir)
                time.sleep(1)  # レート調整（必要なら変更/削除）

        page_token = res.get("nextPageToken")
        if not page_token:
            break

def extract_folder_id(url_or_id: str) -> str:
    """
    Google DriveフォルダのURLまたは直接IDを渡すとフォルダIDを返す
    例:
      https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz
      または
      1AbCdEfGhIjKlMnOpQrStUvWxYz
    """
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', url_or_id)
    if match:
        return match.group(1)
    return url_or_id

def main():
    global success_count, block_count, error_count

    # ===== ここを設定 =====
    dest = DEST            # 保存先ディレクトリ
    url = URL              # ダウンロードするフォルダのURL/ID
    # ====================

    # URLからフォルダIDを抽出
    folder_id = extract_folder_id(url)
    service = get_service()

    # ルートフォルダ名を取得し、保存先のルートを決定
    folder_meta = service.files().get(fileId=folder_id, fields="name").execute()
    folder_name = sanitize(folder_meta.get('name', 'Downloaded_Folder'))
    root_save_dir = unique_path(os.path.join(dest, folder_name))

    print(f'📂 Start recursively downloading: "{folder_name}"')
    download_folder_recursive(service, folder_id, root_save_dir)

    # 集計結果を表示
    print("\n=== Download Summary ===")
    print(f"Successful downloads: {success_count} files")
    print(f"Blocked by 403:       {block_count} files")
    print(f"Other errors:         {error_count} files")

if __name__ == '__main__':
    main()
