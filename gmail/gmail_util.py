import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase


class Mail:
    def __init__(self):
        self.__message = MIMEMultipart("alternative")

    def add_attachment(self, filename, file_name):
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {file_name}",
        )

        # Add attachment to message and convert message to string
        self.__message.attach(part)

    def set_title(self, title):
        self.__message["Subject"] = title

    def set_content(self, content):
        plain_content = MIMEText(content, "plain")
        self.__message.attach(plain_content)

    def set_sender(self, sender):
        self.__message["From"] = sender

    def get_sender(self):
        return self.__message["From"]

    def set_receiver(self, receiver):
        self.__message["To"] = receiver

    def get_receiver(self):
        return self.__message["To"]

    def set_bcc(self, bcc):
        self.__message["Bcc"] = bcc

    def get_bcc(self):
        return self.__message["Bcc"]

    def set_cc(self, cc):
        self.__message["Cc"] = cc

    def get_cc(self):
        return self.__message["Cc"]

    def get_message(self):
        return self.__message

    def __str__(self):
        return self.__message.as_string()


class GmailSender:
    __username = ""
    __password = ""
    __smtp_server = None

    def __init__(self, username, password):
        self.__username = username
        self.__password = password

    def connect(self):
        is_success = False
        smtp_server = "smtp.gmail.com"
        port = 587  # For SSL

        try:
            self.__smtp_server = smtplib.SMTP(smtp_server, port)
            # Secure the connection
            self.__smtp_server.starttls()
            self.__smtp_server.login(self.__username, self.__password)
            is_success = True
            print("Successfully logged in")
        except Exception as e:
            # Print any error messages to stdout
            print(e)
            self.__smtp_server.quit()

        return is_success

    def disconnect(self):
        self.__smtp_server.quit()

    def send(self, mail):
        res = False

        try:
            self.__smtp_server.sendmail(
                    mail.get_sender(), [mail.get_receiver(), mail.get_cc(), mail.get_bcc()], mail.__str__()
                )

            print("SUCCESSFULLY sent to " + mail.get_receiver())
            res = True
        except Exception as e:
            print("ERROR at " + mail.get_receiver() + ", " + e.__str__())

        return res
