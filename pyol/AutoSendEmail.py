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



body = """Dear All,

附件是我的本週週報，祝兩位週末愉快！

          
Sincerely, 
  
Charlie 
  
  
林定垣 Charlie Lin 
財團法人資訊工業策進會 
地方創生服務處/北區產業中心 
DIGI+數位經濟產業推動辦公室 
臺北市大同區承德路三段287號C棟3樓 
---------------------------------------------------------------------- 
Institue for Information Industry 
Regional Industrial Service Division 
A: 3F., Building C., No. 287, Sec, 3, Chengde Rd., Datong Dist., Taipei City 193, Taiwan 
T: 886-2592-2681 ext. 142 F: 886-2591-5876 
E: charlielin@iii.org.tw


"""
path = r"C:\Users\charlielin\Desktop\檔案\工作週報"
a = os.listdir(path)
location = path + "\\" + a[-1]


def mail_job():
    """
    Customize your mail job here.
    Here comes an exmaple.
    Note that when you use this on a windows PC, 
    the format of attachments should be a list containing r-string like
    [r'path1', r'path2'].
    """
    send_mail(
        subject='本週週報_林定垣',
        body= body ,
        to='sandycclin@iii.org.tw;sylviahuang@iii.org.tw',
        #to="charlielin@iii.org.tw",
        #cc='',
        attachments=[location]
    )

schedule.every().friday.at("17:50").do(mail_job)
#schedule.every(10).seconds.do(mail_job)

"""
For more schedule example:
https://github.com/dbader/schedule
"""

while True:
    schedule.run_pending()
    time.sleep(10)
    with open(r"C:\Users\charlielin\Envs\pyol\pyol\readingfile.txt", encoding="utf-8") as file1:
        lines = file1.readlines()
    if lines[0] != "0":
        break
    
