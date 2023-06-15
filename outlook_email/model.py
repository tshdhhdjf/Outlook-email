from typing import Any, List, Union, Optional
from pathlib import Path


class NewMail:
    def __init__(self, mail: Any) -> None:
        self.mail: Any = mail
        """Mail Object"""
        self.to: Union[List[str], str] = ""
        """Mail To Address"""
        self.cc: Optional[Union[List[str], str]] = None
        """Mail Carbon Copy Address"""
        self.bcc: Optional[Union[List[str], str]] = None
        """Mail Blind Carbon Copy Address"""
        self.subject: str = ""
        """Mail Subject"""
        self.bodyformat: int = 2
        """2: html format"""
        self.body: Optional[str] = None
        """Mail Body"""
        self.HTMLbody: Optional[str] = None
        """Mail HTML Body"""
        self.attchments: Optional[List[str]] = None
        """Accthments list"""

    def add_attchment(self, path: Union[Path, str]):
        if not self.attchments:
            self.attchments = []
        path_str = str(Path(path).absolute())
        self.attchments.append(path_str)
        self.mail.Attachments.Add(path_str)

    def send(self):
        if isinstance(self.to, List):
            self.to = ", ".join(self.to)
        self.mail.To = self.to

        if self.cc:
            if isinstance(self.cc, List):
                self.cc = "; ".join(self.cc)
            self.mail.CC = self.cc

        if self.bcc:
            if isinstance(self.bcc, List):
                self.bcc = "; ".join(self.bcc)
            self.mail.BCC = self.bcc

        self.mail.Subject = self.subject
        self.mail.BodyFormat = self.bodyformat
        if self.body:
            self.mail.Body = self.body
        if self.HTMLbody:
            self.mail.HTMLBody = self.HTMLbody
        self.mail.Send()


class Account:
    def __init__(self, account) -> None:
        self.account = account
        self.display_name = account.DisplayName
        self.user_name = account.UserName
        self.smtp_address = account.SmtpAddress


class Attachment:
    def __init__(self, attachment):
        self.attachment = attachment
        self.filename = attachment.FileName

    def save(self, path: Union[Path, str], filename: Optional[str] = None):
        if filename:
            self.attachment.SaveASFile(str(Path(path).absolute() / filename))
        else:
            self.attachment.SaveASFile(
                str(
                    Path(
                        path,
                    ).absolute()
                    / self.filename
                )
            )


class Attachments:
    def __init__(self, mail) -> None:
        self.mail = mail
        """Mail Object"""
        self.items = mail.Attachments

    def iter_attachments(self):
        for item in self.items:
            yield Attachment(item)

    def save_all_attachments(self, path: Union[Path, str]):
        for a in self.iter_attachments():
            a.save(path)

    @property
    def count(self) -> int:
        return self.mail.Attachments.Count


class Mail:
    def __init__(self, mail) -> None:
        self.mail: Any = mail
        """Mail Object"""
        self.to: str = mail.to
        """Mail To Address"""
        self.sender: str = mail.sender

        self.cc: Optional[str] = mail.CC or None
        """Mail Carbon Copy Address"""
        self.bcc: Optional[str] = mail.BCC or None
        """Mail Blind Carbon Copy Address"""
        self.subject: str = mail.subject
        """Mail Subject"""
        self.body: str = mail.body
        """Mail Body"""
        self.HTMLbody: str = mail.HTMLbody
        """Mail HTML Body"""

    @property
    def sender_address(self) -> str:
        mail = self.mail
        if mail.Sender.GetExchangeUser():
            return mail.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return mail.SenderEmailAddress

    @property
    def cc_address(self) -> str:
        # return self.
        pass

    @property
    def attachments(self) -> Attachments:
        return Attachments(self.mail)
