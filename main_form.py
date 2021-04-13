import PySimpleGUI as sg
import os
import threading
import sys
import time
import traceback

import setting_read
import spread_sheet
import outlook_mail
from mail_item import MailItem

import mylogger
import datetime

# ログの定義
logger = mylogger.setup_logger(__name__)


class MainForm:
    def __init__(self):

        # デザインテーマの設定
        sg.theme('BlueMono')

        # ウィンドウの部品とレイアウト
        main_layout = [
            [sg.Text('送信リストを読み込み、自動でメールを送信します\n※メール手動送信：メールの作成画面を表示し、送信は手動で行います\n※メール自動送信：メールの作成画面を表示せず、送信まで自動で行います')],
            [sg.Text('処理状態：', size=(11, 1)), sg.Text('停止中', key='process_status', text_color='#191970', size=(11, 1))],
            [sg.Text(size=(40, 1), justification='center', text_color='#191970', key='message_text1'),
             sg.Button('メール手動送信', key='execute_manual_button'), sg.Button('メール自動送信', key='execute_button')]
        ]

        # ウィンドウの生成
        self.window = sg.Window('自動メール送信', main_layout)

        # 監視フラグ
        self.RUNNING_FLG = False

        # イベントループ
        while True:
            event, values = self.window.read(timeout=100, timeout_key='-TIMEOUT-')

            if event == sg.WIN_CLOSED:
                # ウィンドウのXボタンを押したときの処理
                break

            elif event == 'execute_button' or event == 'execute_manual_button':
                if not self.RUNNING_FLG:
                    self.RUNNING_FLG = True
                    self.MANUAL_FLG = False
                    self.window['execute_button'].update(disabled=True)
                    self.window['execute_manual_button'].update(disabled=True)
                    self.update_text('process_status', '送信処理中．．．')

                    if event == 'execute_manual_button':
                        self.MANUAL_FLG = True
                        logger.debug('メールの手動送信モード')
                    else:
                        logger.debug('メールの自動送信モード')

                    t1 = threading.Thread(target=self.mail_send_worker)
                    # スレッドをデーモン化する
                    t1.setDaemon(True)
                    t1.start()
                else:
                    self.RUNNING_FLG = False
                    self.update_text('process_status', '停止中')

            elif event == '-TIMEOUT-':
                if self.RUNNING_FLG:
                    if not self.window['execute_button'].Disabled:
                        self.window['execute_button'].update(disabled=True)
                        self.window['execute_manual_button'].update(disabled=True)
                        self.update_text('process_status', '送信処理中．．．')
                else:
                    if self.window['execute_button'].Disabled:
                        self.window['execute_button'].update(disabled=False)
                        self.window['execute_manual_button'].update(disabled=False)
                        self.update_text('process_status', '停止中')

            elif event == 'ERROR':
                sg.popup_error('メールの送信処理に失敗しました', title='エラー', button_color=('#f00', '#ccc'))

            elif event == 'SUCCESS':
                sg.Popup('メールの送信処理が完了しました', title='実行結果')

        self.window.close()

    def enable_button(self, key):
        self.window[key].update(disabled=False)

    def disable_button(self, key):
        self.window[key].update(disabled=True)

    def update_text(self, key, message):
        self.window[key].update(message)


    def mail_send_worker(self):

        try:
            # INIファイル読み込み
            self.gspread_info_dic = setting_read.read_config('GSPREAD_SHEET')
            AUTH_KEY_PATH = self.gspread_info_dic['AUTH_KEY_PATH']
            SPREAD_SHEET_KEY = self.gspread_info_dic['SPREAD_SHEET_KEY']
            MAIL_INFO_SHEETNAME = self.gspread_info_dic['MAIL_INFO_SHEETNAME']
            SEND_LIST_SHEETNAME = self.gspread_info_dic['SEND_LIST_SHEETNAME']

            # スプレッドシートからメールの件名と本文を取得
            logger.info('スプレッドシートからメールの件名と本文を取得します')
            self.mailinfo_worksheet = None
            self.mailinfo_worksheet = spread_sheet.connect_gspread(AUTH_KEY_PATH, SPREAD_SHEET_KEY, MAIL_INFO_SHEETNAME)
            mail_template_list = []
            mail_template_list = spread_sheet.read_gspread_sheet(self.mailinfo_worksheet)

            self.mail_message = ''
            self.mail_subject = ''
            for mail_template in mail_template_list[1:]:
                self.mail_subject = mail_template[0]
                self.mail_message = mail_template[1]

            # スプレッドシートから送信リストの取得
            logger.info('スプレッドシートから送信リストを取得します')
            self.sendlist_worksheet = None
            self.sendlist_worksheet = spread_sheet.connect_gspread(AUTH_KEY_PATH, SPREAD_SHEET_KEY, SEND_LIST_SHEETNAME)
            mail_list = []
            mail_list = spread_sheet.read_gspread_sheet(self.sendlist_worksheet)

            # 取得したデータをクラスに設定（ヘッダーを飛ばす）
            mail_send_list = []
            row_num = 2
            for item_data in mail_list[1:]:
                _mail_item = MailItem()
                _mail_item.spread_sheet_no = row_num
                _mail_item.exhibition_name = item_data[0]
                _mail_item.client_num = item_data[1]
                _mail_item.send_date = item_data[2]
                _mail_item.client_name = item_data[3]
                _mail_item.client_hp = item_data[4]
                _mail_item.client_mail_address = item_data[5]
                _mail_item.crowdfunding_url = item_data[6]
                _mail_item.product = item_data[7]
                _mail_item.person_in_charge = item_data[8]
                _mail_item.person_mail_address = item_data[9]
                _mail_item.skype_id = item_data[10]
                _mail_item.facebook_id = item_data[11]
                _mail_item.note = item_data[13]
                if item_data[14] == '○':
                    _mail_item.send_flg = True
                else:
                    _mail_item.send_flg = False

                _mail_item.mail_message = self.mail_message.format(client_hp=_mail_item.client_hp)
                _mail_item.mail_subject = self.mail_subject

                mail_send_list.append(_mail_item)

                row_num += 1

            logger.info('メールの送信処理を開始します')

            mail_send_count = 0
            mail_skip_count = 0
            SEND_DATE_COL = 3
            for mail_item in mail_send_list:
                if mail_item.send_flg:
                    logger.info('メール送信フラグTrueのため、送信する。スプレッドシート行番号：' + str(mail_item.spread_sheet_no))
                    if (mail_item.client_mail_address and mail_item.mail_subject and mail_item.mail_message):
                        # メールの送信処理
                        outlook_mail.mail_send(to=mail_item.client_mail_address, cc='', bcc='', subject=mail_item.mail_subject, body=mail_item.mail_message, body_format=1, manual_flg=self.MANUAL_FLG)
                        mail_send_count += 1
                        logger.info('メール送信完了')

                        # スプレッドシートに送信日時を書き込み
                        now = datetime.datetime.now()
                        now_str = now.strftime('%Y/%m/%d %H:%M:%S')
                        cell_index = mail_item.spread_sheet_no
                        spread_sheet.update_gspread_sheet(worksheet=self.sendlist_worksheet, cell_row=cell_index, cell_col=SEND_DATE_COL, update_value=now_str)
                        logger.info('スプレッドシートに送信日時書き込み完了')
                        time.sleep(1.0)
                    else:
                        mail_skip_count += 1
                        logger.info('メールに必要なパラメータがないため、送信しない。スプレッドシート行番号：' + str(mail_item.spread_sheet_no))
                else:
                    mail_skip_count += 1
                    logger.info('メール送信フラグFalseのため、送信しない。スプレッドシート行番号：' + str(mail_item.spread_sheet_no))

            logger.info('メール送信処理完了。送信件数：' + str(mail_send_count) + 'スキップ件数：' + str(mail_skip_count))

        except Exception as err:
            logger.error(err)
            logger.error(traceback.format_exc())

        finally:
            self.RUNNING_FLG = False


def expexpiration_date_check():
    import datetime
    now = datetime.datetime.now()
    expexpiration_datetime = now.replace(month=4, day=15, hour=12, minute=0, second=0, microsecond=0)
    logger.info("有効期限：" + str(expexpiration_datetime))
    if now < expexpiration_datetime:
        return True
    else:
        return False


# プログラム実行部分
if __name__ == "__main__":
    logger.info('プログラム起動開始')

    # # 有効期限チェック
    # if not (expexpiration_date_check()):
    #     logger.info("有効期限切れため、プログラム起動終了")
    #     sys.exit(0)

    app = MainForm()