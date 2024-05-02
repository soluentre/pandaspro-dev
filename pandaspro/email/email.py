from pandaspro.user_config.email_path import config
from jinja2 import Template
from pandaspro.email.utlis import replace_with_dict
import win32com.client as win32


class email_template:
    """
    In the user_config.email_path, the email_root variable should be connecting to the user's email profile

    The root folder should have:
        1. a 'templates' folder, with html files in it
        2. a 'config' file noting the engines to parse each template
    """
    def __init__(
            self,
            template: str = None,
            *args,
            **kwargs
    ):
        myemail = config.general[template]['engine'](*args, **kwargs)
        self.email = myemail
        self.to = myemail.to
        self.cc = myemail.cc
        self.subject = myemail.subject
        self.template_raw = myemail.template_raw
        self.render_dict = myemail.render_dict
        self.input = myemail.input

        with open(self.template_raw, 'r') as file:
            self.html_raw = file.read()

        self.html = Template(self.html_raw).render(myemail.render_dict)
        self.html = replace_with_dict(self.html, myemail.input)


class email:
    def __init__(
            self,
            template: email_template = None,
            subject: str = '',
            to: str = 'swang12@worldbankgroup.org',
            cc: str = '',
            bcc: str = '',
    ):
        self.template = template
        self.to = self.template.to if self.template.to is not None else to
        self.cc = self.template.cc if self.template.cc is not None else cc
        self.bcc = bcc
        self.subject = self.template.subject if self.template.subject is not None else subject

        olApp = win32.Dispatch('Outlook.Application')
        olNS = olApp.GetNameSpace('MAPI')
        mailItem = olApp.CreateItem(0)
        mailItem.To = self.to
        mailItem.CC = self.cc
        mailItem.BCC = self.bcc
        mailItem.Subject = self.subject
        mailItem.HTMLBody = template.html

        self.olNS = olNS
        self.mail = mailItem

    def attach(self, attachments: list):
        for item in attachments:
            self.mail.Attachments.Add(item)
        return self

    @property
    def display(self):
        self.mail.Display()


if __name__ == '__main__':
    att = [
        r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Emails and Manuals\STA and DAIS\T200011_Short_Term_Assignment_Developmental_Assignment_Memorandum.pdf',
    ]
    mytemplate = email_template(
        'sta_dais_init',
        rec_mgr_upi=288695,
        staff_upi=562499,
        position=56160
    )
    e = email(mytemplate)
    e.attach(att).display
