import win32com.client
from data_structures import EmailInfo


class Email:
    def __init__(self, email_info: EmailInfo,
                 subject: str, body: str, attachment: str = None) -> None:
        self.email = email_info
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.subject, self.body = subject, body
        self.attachment = attachment if attachment else None

    def run(self) -> None:
        for mail_address in self.email.email_list:
            self.send_mail(mail_address)

    def send_mail(self, mail_address: str) -> None:
        mail = self.outlook.CreateItem(0)
        mail.To = mail_address
        mail.Subject = self.subject
        mail.Body = self.body
        if self.attachment:
            mail.Attachments.Add(self.attachment)
        mail.Send()
