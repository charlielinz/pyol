from win32com import client
import schedule
import time
import os


"""
Define file path
"""

this_py_file_path = os.path.abspath(__file__)
this_py_file_dir_path = os.path.dirname(this_py_file_path)

config_folder_dir_path = os.path.join(this_py_file_dir_path, "config_folder")
folder_paths_list = os.listdir(config_folder_dir_path)


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
        for attachment_config in attachments:
            mail.Attachments.Add(attachment_config)
    if just_show:
        mail.display()
    else:
        mail.Send()


for folder_path in folder_paths_list:
     
    config_dir_path = os.path.join(config_folder_dir_path, folder_path)
    status_config_path = os.path.join(config_dir_path, "status_config.txt")

    with open(status_config_path, encoding="utf-8") as status_config:
        status_list = status_config.readlines()
        status = status_list[0].replace("\n", "")
        period = int(status_list[1])
    if status == "on":
        def send_mail_by_config_file(config_dir_path):
              
            subject_path = os.path.join(config_dir_path, "subject_config.txt")
            to_path = os.path.join(config_dir_path, "to_config.txt")
            body_path = os.path.join(config_dir_path, "body_config.txt")


            with open(subject_path, encoding="utf-8") as subject_config:
                subject_list = subject_config.readlines()
            subject = subject_list[0]

            with open(to_path, encoding="utf-8") as to_config:
                to_list = to_config.readlines()
            to_list_str = ";".join(to_list)
            to = to_list_str.replace("\n", "")

            with open(body_path, encoding="utf-8") as body_config:
                body_list = body_config.readlines()
            body = "".join(body_list)

            send_mail(subject=subject, body=body, to=to)

        schedule.every(period).seconds.do(send_mail_by_config_file,config_dir_path=os.path.join(config_folder_dir_path, folder_path))


while True:
    schedule.run_pending()
    time.sleep(10)



"""
For more schedule example:
https://github.com/dbader/schedule
"""
