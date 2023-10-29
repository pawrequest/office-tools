import webbrowser
from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path

from win32com.client import Dispatch
from win32com.universal import com_error


class EmailHandler(ABC):
    @abstractmethod
    def send_email(self, email: "Email") -> None:
        ...


@dataclass
class Email:
    to_address: str
    subject: str
    body: str
    attachment_path: Path or None = None

    def send(self, sender: EmailHandler) -> None:
        sender.send_email(self)


class OutlookSender(EmailHandler):
    def send_email(self, email: Email):
        try:
            outlook = Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = email.to_address
            mail.Subject = email.subject
            mail.Body = email.body
            if email.attachment_path:
                mail.Attachments.Add(str(email.attachment_path))
                print("Added attachment")
            print("Sending email [disabled - uncomment to enable]")
            # mail.Display()
            # mail = None
            # mail.Send()
        except com_error as e:
            msg = f"Outlook not installed - {e.args[0]}"
            print(msg)
            raise EmailError(msg)
        except Exception as e:
            msg = f"Failed to send email with error: {e.args[0]}"
            print(msg)
            raise EmailError(msg)


class EmailError(Exception):
    ...


class GmailSender(EmailHandler):
    def send_email(self, email: Email) -> None:
        compose_link = gmail_compose_link(email)
        webbrowser.open(compose_link)
        print(
            "Opened Gmail compose window, no attachment possible, please attach manually, implement better solution i.e. gmail oauth"
        )


def gmail_compose_link(email: Email) -> str:
    to_encoded = email.to_address.replace("@", "%40")
    subject_encoded = email.subject.replace(" ", "%20")
    body_encoded = email.body.replace(" ", "%20")
    compose_link = f"https://mail.google.com/mail/u/0/?view=cm&fs=1&to={to_encoded}&su={subject_encoded}&body={body_encoded}"
    return compose_link
