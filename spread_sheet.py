import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import csv
import sys
from datetime import datetime

def read_gspread_sheet_from_folder(auth_keyfile_path, gspread_folderid, book_name, sheet_title):
    # GoogleAPI
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    # 認証用キー
    json_keyfile_path = auth_keyfile_path
    # サービスアカウントキーを読み込む
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)
    # pydrive用の認証
    gauth = GoogleAuth()
    gauth.credentials = credentials
    drive = GoogleDrive(gauth)
    # スプレッドシート格納フォルダ（https://drive.google.com/drive/folders/<フォルダのID>）
    folder_id = gspread_folderid
    # スプレッドシート格納フォルダ内のファイル一覧取得
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % folder_id}).GetList()
    # ファイル一覧からファイル名のみを抽出
    title_list = [file['title'] for file in file_list]
    # 任意のブック名
    book_name = book_name
    # gspread用の認証
    gc = gspread.authorize(credentials)
    # スプレッドシートIDを取得
    sheet_id = [file['id'] for file in file_list if file['title'] == book_name]
    sheet_id = sheet_id[0]
    # ブックを開く
    workbook = gc.open_by_key(sheet_id)
    # 存在するワークシートの情報を全て取得
    worksheet = workbook.worksheet(sheet_title)
    # 全ての値を、辞書型の変数に設定
    # data_dic = worksheet.get_all_records()
    # 全ての値を、リスト型の変数に設定
    data_list = worksheet.get_all_values()

    return data_list


def read_gspread_sheet(worksheet):

    # 全ての値を、辞書型の変数に設定
    # data_dic = worksheet.get_all_records()
    # 全ての値を、リスト型の変数に設定
    data_list = worksheet.get_all_values()

    return data_list


def update_gspread_sheet(worksheet, cell_row: int, cell_col: int, update_value: str):
    # セルの更新
    worksheet.update_cell(cell_row, cell_col, update_value)

    return


def connect_gspread(auth_keyfile_path: str, spreadsheet_key: str, sheet_title: str):
    # GoogleAPI
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    # 認証用キー
    json_keyfile_path = auth_keyfile_path
    # サービスアカウントキーを読み込む
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)

    #認証情報を使ってスプレッドシートの操作権を取得
    gs = gspread.authorize(credentials)

    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    workbook = gs.open_by_key(spreadsheet_key)
    worksheet = workbook.worksheet(sheet_title)

    return worksheet


def import_gspread(auth_keyfile_path, csv_file_path, gspread_folderid, book_name, sheet_title):
    # GoogleAPI
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    # 認証用キー
    json_keyfile_path = auth_keyfile_path
    # サービスアカウントキーを読み込む
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)
    # pydrive用の認証
    gauth = GoogleAuth()
    gauth.credentials = credentials
    drive = GoogleDrive(gauth)
    # インポートするCSVファイル
    csv_path = csv_file_path
    # スプレッドシート格納フォルダ（https://drive.google.com/drive/folders/<フォルダのID>）
    folder_id = gspread_folderid
    # スプレッドシート格納フォルダ内のファイル一覧取得
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % folder_id}).GetList()
    # ファイル一覧からファイル名のみを抽出
    title_list = [file['title'] for file in file_list]
    # 任意のブック名
    book_name = book_name
    # ブックの存在判定
    if book_name not in title_list:
        # ブック新規作成
        f = drive.CreateFile({
            'title': book_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            "parents": [{"id": folder_id}]})
        f.Upload()
        # gspread用に認証
        gc = gspread.authorize(credentials)
        # スプレッドシートIDを取得
        sheet_id = f['id']
        # ブックを開く
        workbook = gc.open_by_key(sheet_id)
        # ワークシート名を任意のシート名に変更
        worksheet = workbook.worksheet('Sheet1')
        sheet_name = sheet_title
        worksheet.update_title(sheet_name)
    else:
        # gspread用の認証
        gc = gspread.authorize(credentials)
        # スプレッドシートIDを取得
        sheet_id = [file['id'] for file in file_list if file['title'] == book_name]
        sheet_id = sheet_id[0]
        # ブックを開く
        workbook = gc.open_by_key(sheet_id)
        # 存在するワークシートの情報を全て取得
        worksheets = workbook.worksheets()
        # ワークシート名のみをリストへ格納
        worksheets_title_list = [sheet.title for sheet in worksheets]
        # 任意のワークシート名
        sheet_name = sheet_title
        # 重複ファイル削除フラグ
        old_file_delflg = False
        old_file_name = sheet_name + "_old" + datetime.now().strftime("%Y%m%d_%H%M%S")
        # 同一ワークシート名の存在判定
        if sheet_name in worksheets_title_list:
            # シートが１つしかない場合、先に削除できないので、既存のシート名を一度変更する
            workbook.worksheet(sheet_name).update_title(old_file_name)
            old_file_delflg = True

        # ワークシート追加
        workbook.add_worksheet(title=sheet_name, rows=1000, cols=26)

        # ワークシート削除
        if old_file_delflg:
            workbook.del_worksheet(workbook.worksheet(old_file_name))

    # スプレッドシートにCSVをインポート
    workbook.values_update(
        sheet_name,
        params={'valueInputOption': 'USER_ENTERED'},
        body={'values': list(csv.reader(open(csv_path, encoding='utf_8_sig')))}
    )



if __name__ == '__main__':
    AUTH_KEY_PATH = './keyfile/pchislo-csv-upload-db490488affa.json'
    FOLDER_ID = '12xx-WHudb4rHn_D2WADyHLuDK4wo4kyV'
    BOOK_NAME = '商品リスト'
    SHEET_TITLE = 'シート1'
    read_gspread_sheet(AUTH_KEY_PATH, FOLDER_ID, BOOK_NAME, SHEET_TITLE)