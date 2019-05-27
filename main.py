from gmail.gmail_util import *
from excel.excel_util import *
import os
import shutil
import settings


GAMFED_INVITE_PREFIX = settings.renamed_pdf_prefix

invite_list_file = settings.invite_excel
invite_pdf_folder = settings.invite_pdf_folder
invite_pdf_folder_sent = settings.invite_pdf_folder_sent

receiver = "<RECEIVER>"
content = "Merhaba " + receiver + """,

Bu yıl 2 Haziran 2019 tarihinde 3. kez düzenlenecek Gamification Meetup için oluşturulmuş 1 adet davetiyeyi ekte bulabilirsin.

Etkinliğimizin daha geniş kitlelere ulaşması için sosyal medyada #oyunagel ve #gamificationmeetup hashtagleri ile paylaşım yaparak sen de yolculuğumuzun bir parçası olabilirsin.

Senin gibi 'oyunu ciddiye alan' birini etkinliğinizde görmek bizi çok mutlu edecek.

Oyun dolu bir gün dileriz.

Rutkay
Gamfed Türkiye Ekibi
http://www.gamificationmeetup.com/
"""

title = "Gamfed Türkiye 3. Gamification Meetup Davetiyesi"


excel_util = ExcelUtil(invite_list_file)
list = excel_util.get_invite_list()

list_dir = os.listdir(invite_pdf_folder)
list_dir = [f.lower() for f in list_dir]
pdf_list = sorted(list_dir)


mail_sender = GmailSender(settings.login_username, settings.login_password)
mail_sender.connect()
for t in range(0, list.__len__()):
    i = list[t]

    content_temp = str(content)
    content_temp = content_temp.replace(receiver,i.name)

    email = Mail()
    email.set_title(title)
    email.set_content(content_temp)
    email.set_sender(settings.email_sender)
    email.set_receiver(i.mail)
    email.set_cc(settings.email_cc)
    email.set_bcc(settings.email_bcc)

    file = os.path.join(invite_pdf_folder, pdf_list[t])
    renamed_file = os.path.join(invite_pdf_folder,GAMFED_INVITE_PREFIX + str(t) + "_" + i.file_name + ".pdf")

    are_you_sure = input("Are you sure to send this mail to \n" + i.mail + "\n as \n" + str(content_temp)[0:40] + "\n With file: " + renamed_file + "\n")
    if not str.__eq__(are_you_sure, "y"):
        continue

    os.rename(file, renamed_file)

    email.add_attachment(renamed_file, GAMFED_INVITE_PREFIX + str(t) + "_" + i.file_name + ".pdf")

    res = mail_sender.send(email)

    if not res:
        os.rename(renamed_file, file)
    else:
        renamed_file_move = os.path.join(invite_pdf_folder_sent,GAMFED_INVITE_PREFIX + str(t) + "_" + i.file_name + ".pdf")
        shutil.move(renamed_file, renamed_file_move)
