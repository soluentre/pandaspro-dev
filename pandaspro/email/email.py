from pandaspro.user_config.email_path import config

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
        self.template = config