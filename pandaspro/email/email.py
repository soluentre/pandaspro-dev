from pandaspro.user_config.email_path import email_root
from pandaspro.user_config.email_path import email_config
from jinja2 import Template
import os


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
            keys: dict = None,
    ):
        template_path = os.path.join(email_root + f'/templates/{template}.html')
        with open(template_path, 'r') as file:
            html_raw = file.read()


        self.html_raw = html_raw
        # self.html = html