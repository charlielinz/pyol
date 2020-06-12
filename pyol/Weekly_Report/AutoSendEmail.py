from win32com import client
import schedule
import time
import os



def send_mail(subject, body, to, cc='', bcc='', attachments=[], just_show=False):
    """
    The method of sending Email by outlook.
    """
    olMailItem = 0x0
    outlook_client = client.Dispatch("Outlook.Application")
    mail = outlook_client.CreateItem(olMailItem)
    mail.Subject = subject
    mail.Body = body
    mail.To = to
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    if attachments:
        for attachment in attachments:
            mail.Attachments.Add(attachment)
    if just_show:
        mail.display()
    else:
        mail.Send()

with open(r"C:\Users\charlielin\Envs\pyol\pyol\Weekly_Report\Body_weeklyreport.txt", encoding="utf-8") as content:
    body_content = content.readlines()

body = "".join(body_content)

path = r"C:\Users\charlielin\Desktop\檔案\工作週報"
a = os.listdir(path)
location = path + "\\" + a[-1]


def mail_job():
    """
    the format of attachments should be a list containing r-string like
    [r'path1', r'path2'].
    """
    send_mail(
        subject='本週週報_林定垣',
        body=body,
        to='sandycclin@iii.org.tw;sylviahuang@iii.org.tw',
        attachments=[location]
    )

schedule.every().friday.at("17:55").do(mail_job)
"""
For more schedule example:
https://github.com/dbader/schedule
"""

while True:
    schedule.run_pending()
    time.sleep(10)
    with open(r"C:\Users\charlielin\Envs\pyol\pyol\Weekly_Report\readingfile.txt", encoding="utf-8") as file1:
        lines = file1.readlines()
    if lines[0] != "0":
        break