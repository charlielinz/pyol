from win32com import client
import schedule
import time

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

def mail_job():
    """
    Customize your mail job here.
    Here comes an exmaple.
    Note that when you use this on a windows PC, 
    the format of attachments should be a list containing r-string like
    [r'path1', r'path2'].
    """
    with open(r"C:\Users\charlielin\Envs\pyol\pyol\Weekly_Report\Subject.txt", encoding="utf-8") as file1:
        list_subject = file1.readlines()
        subject = list_subject[0]

    with open(r"C:\Users\charlielin\Envs\pyol\pyol\Weekly_Report\To.txt", encoding="utf-8") as file2:
        list_to = file2.readlines()
        string_to = ''
        for to in list_to:
            string_to += to
            string_to += ';'


    body = ""
    
    send_mail(
        subject=subject,
        body=body,
        to=string_to,
        #cc='jimmy_lin@chief.com.tw',
        #attachments=[
        #    r'C:\Users\charlielin\Desktop\檔案\工作週報for自動寄信\週報 Charlie_lin.xlsx',
        #            ]
    )

schedule.every(10).seconds.do(mail_job)

"""
For more schedule example:
https://github.com/dbader/schedule
"""

while True:
    schedule.run_pending()
    time.sleep(10)
    
    
