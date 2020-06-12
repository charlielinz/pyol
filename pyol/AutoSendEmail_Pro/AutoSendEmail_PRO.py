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
    mail.Subject = subject_config
    mail.Body = body_config
    mail.To = to_config
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    if attachments:
        for attachment_config in attachments:
            mail.Attachments.Add(attachment_config)
    if just_show:
        mail.display()
    else:
        mail.Send()


def send_mail_by_config_file(config_dir_path):
    
    config_dir_path = r"C:\Users\charlielin\Envs\pyol\pyol\AutoSendEmail_Pro\weekly_report"

    with open(config_dir_path + r"\subject_config.txt", encoding="utf-8") as file1:
        subject = file1.readlines()
    subject_config = subject[0]

    with open(config_dir_path + r"\to_config.txt", encoding="utf-8") as file2:
        to = file2.readlines()
    to_list = ";".join(to)
    to_config = to_list.replace("\n","")

    with open(config_dir_path + r"\body_config.txt", encoding="utf-8") as file3:
        body = file3.readlines()
    body_config = "".join(body)
    
        

def mail_job():
    send_mail_by_config_file(config_dir_path)

if schedule.every(10).seconds.do(mail_job)

"""
For more schedule example:
https://github.com/dbader/schedule
"""

while True:
    schedule.run_pending()
    time.sleep(10)
    
    
