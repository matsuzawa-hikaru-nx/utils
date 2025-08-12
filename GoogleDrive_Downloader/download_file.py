from __future__ import print_function
import os
import io
import re
from typing import Tuple
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ===== CONFIG =====
# 認証/保存まわり
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
OAUTH_DIR = './Outh'                          # credentials.json / token.json を置く場所
SAVE_DIR  = './Downloaded/SingleFile'         # ダウンロード先ディレクトリ
FILE_URL_OR_ID = '1mYQcAB8_lzfyuOw4tbvbkKG4vQsv5U7W'  # ファイルのURL or ID

# 変換/安全性オプション
EXPORT_GOOGLE_DOCS = True                     # Google Docs/Sheets/Slides をOffice/PDF等に変換して保存
ALLOW_ABUSE_DOWNLOAD = True                   # マルウェア/スパム判定ファイルを自己責任でDLする（acknowledgeAbuse）
CHUNK_SIZE_MB = 8                             # ダウンロードチャンク（MB）
# ==================

def extract_file_id(url_or_id: str) -> str:
    """
    ファイルのURLまたはID文字列から fileId を取り出す。
      例: https://drive.google.com/file/d/<ID>/view?..., ?id=<ID>, もしくは <ID> だけ
    """
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url_or_id)
    if m:
        return m.group(1)
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url_or_id)
    if m:
        return m.group(1)
    return url_or_id.strip()

def get_service():
    """Google Drive API 認証 & サービス生成"""
    os.makedirs(OAUTH_DIR, exist_ok=True)
    token_path = os.path.join(OAUTH_DIR, 'token.json')
    cred_path  = os.path.join(OAUTH_DIR, 'credentials.json')

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w', encoding='utf-8') as f:
            f.write(creds.to_json())

    return build('drive', 'v3', credentials=creds, cache_discovery=False)

def make_unique_path(path: str) -> str:
    """同名ファイルがある場合に filename (1).ext のように重複回避"""
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    idx = 1
    while True:
        cand = f"{root} ({idx}){ext}"
        if not os.path.exists(cand):
            return cand
        idx += 1

def pick_export_mime_and_ext(google_mime: str) -> Tuple[str, str]:
    """
    Google独自形式の mimeType からエクスポート先の MIME と拡張子を決める。
    足りない形式は PDF にフォールバック。
    """
    mapping = {
        'application/vnd.google-apps.document':     ('application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'),
        'application/vnd.google-apps.spreadsheet':  ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx'),
        'application/vnd.google-apps.presentation': ('application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx'),
        'application/vnd.google-apps.drawing':      ('image/png', '.png'),
    }
    if google_mime in mapping:
        return mapping[google_mime]
    return ('application/pdf', '.pdf')

def ensure_extension(filename: str, desired_ext: str) -> str:
    """拡張子が無ければ desired_ext を付与（既にあればそのまま）"""
    root, ext = os.path.splitext(filename)
    return filename if ext else (root + desired_ext)

def download_single_file(service, file_id_or_url: str, save_dir: str):
    """単一ファイルを元の名前・拡張子で指定フォルダに保存"""
    os.makedirs(save_dir, exist_ok=True)
    file_id = extract_file_id(file_id_or_url)

    # メタデータ取得（名前・MIME）
    meta = service.files().get(fileId=file_id, fields='name, mimeType').execute()
    name = meta.get('name', 'downloaded_file')
    mime = meta.get('mimeType', '')

    # 保存パス/リクエスト準備
    is_google_native = mime.startswith('application/vnd.google-apps')
    if is_google_native and EXPORT_GOOGLE_DOCS:
        export_mime, export_ext = pick_export_mime_and_ext(mime)
        save_name = ensure_extension(name, export_ext)
        dest_path = make_unique_path(os.path.join(save_dir, save_name))
        request = service.files().export_media(fileId=file_id, mimeType=export_mime)
    else:
        # 通常ファイルはDrive上の名前をそのまま利用
        dest_path = make_unique_path(os.path.join(save_dir, name))
        kwargs = {}
        if ALLOW_ABUSE_DOWNLOAD:
            kwargs['acknowledgeAbuse'] = True
        request = service.files().get_media(fileId=file_id, **kwargs)

    # ダウンロード（チャンクサイズ指定）
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request, chunksize=CHUNK_SIZE_MB * 1024 * 1024)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"{name}: {int(status.progress() * 100)}%")

    with open(dest_path, 'wb') as f:
        f.write(fh.getvalue())

    print(f"✅ 保存完了: {dest_path}")

def main():
    service = get_service()
    download_single_file(service, FILE_URL_OR_ID, SAVE_DIR)

if __name__ == '__main__':
    main()
