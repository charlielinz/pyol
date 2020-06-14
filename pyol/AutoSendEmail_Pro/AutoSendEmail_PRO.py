from win32com import client
import schedule
import time
import os.path


"""Define file path"""
this_py_file_path = os.path.abspath(__file__)
this_py_file_dir_path = os.path.dirname(this_py_file_path)

config_dir_path = os.path.join(this_py_file_dir_path,"weekly_report")
subject_path = os.path.join(config_dir_path,"subject_config.txt")
to_path = os.path.join(config_dir_path,"to_config.txt")
body_path = os.path.join(config_dir_path,"body_config.txt")

#subject_config = ""
#to_config = ""
#body_config = ""
    
def send_mail(subject, body, to, cc='', bcc='', attachments=[], just_show=False):
   
    """The method of sending Email by outlook."""
    
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
        for attachment_config in attachments:
            mail.Attachments.Add(attachment_config)
    if just_show:
        mail.display()
    else:
        mail.Send()


def send_mail_by_config_file(config_dir_path):
    
    with open(subject_path, encoding="utf-8") as file1:
        subject = file1.readlines()
    subject_config = subject[0]

    with open(to_path, encoding="utf-8") as file2:
        to = file2.readlines()
    to_list = ";".join(to)
    to_config = to_list.replace("\n","")

    with open(body_path, encoding="utf-8") as file3:
        body = file3.readlines()
    body_config = "".join(body)

    send_mail(subject=subject_config,body=body_config,to=to_config)
        
def mail_job():
    send_mail_by_config_file(config_dir_path=config_dir_path)

schedule.every(10).seconds.do(mail_job)

"""
For more schedule example:
https://github.com/dbader/schedule
"""

while True:
    schedule.run_pending()
    time.sleep(10)
    
    
