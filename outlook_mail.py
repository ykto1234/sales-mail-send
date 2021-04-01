import win32com.client
import pythoncom
import mylogger

# ログの定義
logger = mylogger.setup_logger(__name__)



def mail_send(to: str, cc: str, bcc: str, subject: str, body: str, body_format=1, manual_flg=True):
    """
    Outlookを起動し、メールを送信します。

    Args:
        to (str): メールのTO
        cc (str): メールのCC
        subject (str): メールの件名
        body (str): メールの本文
        body_format (int): メール本文のフォーマット（1 テキスト, 2 HTML, 3 リッチテキスト）
    """
    # Excelを起動する前にこれを呼び出す
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)

    mail.to = to
    mail.cc = cc
    mail.bcc = bcc
    mail.subject = subject
    mail.bodyFormat = body_format
    mail.body = body

    # path = r'' # 添付ファイルは絶対パスで指定
    # mail.Attachments.Add (path)

    if manual_flg:
        # 手動の場合、出来上がったメールを表示
        logger.debug('メール画面を表示する')
        mail.Display(True)
    else:
        # 自動の場合、画面表示せず自動で送信する
        logger.debug('メール画面を表示せず送信する')
        mail.Send()

    return