from win32com import client
import schedule
import time
import os
import json


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
        for attachment in attachments:
            mail.Attachments.Add(attachment)
    if just_show:
        mail.display()
    else:
        mail.Send()


def read_file_as_long_string(config_path):
    with open(config_path, encoding="utf-8") as config_dct:
        lines = config_dct.readlines()
        lines_fixed = []
        for line in lines:
            if line[-1] == '\n':
                lines_fixed += [line[0:-1]]
            else:
                lines_fixed += [line]
    file_as_long_string = "".join(lines_fixed)
    return file_as_long_string


def send_mail_by_config_file(config_dir_path):

    config_path = os.path.join(config_dir_path, "config.txt")
    body_path = os.path.join(config_dir_path, "body_config.txt")
    attachment_folder_path = os.path.join(config_dir_path, "attachment")
    attachment_list = os.listdir(attachment_folder_path)
    config_json = read_file_as_long_string(config_path=config_path)
    config = json.loads(config_json)

    with open(config_path, encoding="utf-8") as config_dct:
        lines = config_dct.readlines()
        lines_fixed = []
        for line in lines:
            if line[-1] == '\n':
                lines_fixed += [line[0:-1]]
            else:
                lines_fixed += [line]

    with open(body_path, encoding="utf-8") as body_config:
        body_list = body_config.readlines()
    body = "".join(body_list)

    attachment_path_list = []
    for attachment in attachment_list:
        attachment_path = os.path.join(attachment_folder_path, attachment)
        attachment_path_list += [attachment_path]

    send_mail(subject=config["subject"], body=body,
              to=config["to"], attachments=attachment_path_list)


for folder_path in folder_paths_list:

    config_dir_path = os.path.join(config_folder_dir_path, folder_path)
    status_config_path = os.path.join(config_dir_path, "status_config.txt")
    config_path = os.path.join(config_dir_path, "config.txt")
    config_json = read_file_as_long_string(config_path=config_path)
    config = json.loads(config_json)

    if config["status"] == "on":

        if config["cycle"] == "every_minute":
            if config["period"] == True:
                if config["time"] == True:
                    schedule.every(config["period"]).minutes.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every(config["period"]).minutes.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                if config["time"] == True:
                    schedule.every().minute.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every().minute.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_hour":
            if config["period"] == True:
                if config["time"] == True:
                    schedule.every(config["period"]).hours.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every(config["period"]).hours.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                if config["time"] == True:
                    schedule.every().hour.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every().hour.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_day":
            if config["period"] == True:
                if config["time"] == True:
                    schedule.every(config["period"]).days.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every(config["period"]).days.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                if config["time"] == True:
                    schedule.every().day.at(config["time"]).do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)
                else:
                    schedule.every().day.do(
                        send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_monday":
            if config["time"] == True:
                schedule.every().monday.at(config["time"]).do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                schedule.every().monday.do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_tuesday":
            if config["time"] == True:
                schedule.every().monday.at(config["time"]).do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                schedule.every().monday.do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_wednesday":
            if config["time"] == True:
                schedule.every().monday.at(config["time"]).do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                schedule.every().monday.do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_thursday":
            if config["time"] == True:
                schedule.every().monday.at(config["time"]).do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                schedule.every().monday.do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)

        elif config["cycle"] == "every_friday":
            if config["time"] == True:
                schedule.every().monday.at(config["time"]).do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)
            else:
                schedule.every().monday.do(
                    send_mail_by_config_file, config_dir_path=config_dir_path)


while True:
    schedule.run_pending()
    time.sleep(10)


"""
For more schedule example:
https://github.com/dbader/schedule
"""
