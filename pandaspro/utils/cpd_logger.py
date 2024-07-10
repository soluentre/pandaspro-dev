import logging
from datetime import datetime
from functools import wraps
import re


def cpd_logger(cls):
    class CustomFormatter(logging.Formatter):
        # ANSI escape sequences for text formatting
        RESET = "\033[0m"
        BOLD = "\033[1m"
        RED = "\033[31m"
        BLUE = "\033[34m"

        def format(self, record):
            msg = record.msg
            if msg == "\n":
                return "\n"

            # Apply bold and red formatting to text enclosed in **
            msg = re.sub(r"\*\*(.*?)\*\*", f"{self.BOLD}{self.BLUE}\\1{self.RESET}", msg)
            record.msg = msg

            return super().format(record)

    original_init = cls.__init__

    @wraps(original_init)
    def new_init(self, *args, **kwargs):
        # Ensure the debug and debug_file attributes are set before calling original __init__
        self.debug = kwargs.get('debug', 'critical').lower()
        self.debug_file = kwargs.get('debug_file', None)

        original_init(self, *args, **kwargs)

        self.logger = logging.getLogger(cls.__name__)
        self.configure_logger()

        # Adding log_section and end_log_section methods to the instance
        self.debug_section_lv1 = self._debug_section_lv1
        self.debug_section_lv2 = self._debug_section_lv2
        self.info_section_lv1 = self._info_section_lv1
        self.info_section_lv2 = self._info_section_lv2
        self.reconfigure_logger = self._reconfigure_logger

    def configure_logger(self):
        if self.logger.hasHandlers():
            self.logger.handlers.clear()

        # Console Handler
        console_handler = logging.StreamHandler()
        console_formatter = CustomFormatter('%(asctime)s %(levelname)s - %(message)s', datefmt='%H:%M:%S')
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)

        # Debug Level Settings
        if hasattr(self, 'debug') and self.debug == 'debug':
            self.logger.setLevel(logging.DEBUG)
        elif hasattr(self, 'debug') and self.debug == 'info':
            self.logger.setLevel(logging.INFO)
        else:
            self.logger.setLevel(logging.CRITICAL)

        # File Handler
        if hasattr(self, 'debug_file') and self.debug_file:
            file_name = f"putxl_log_{self.debug}_{datetime.now().strftime('%Y%m%d')}.txt" if isinstance(self.debug_file,
                                                                                                        bool) else f"{self.debug_file}.txt"
            file_handler = logging.FileHandler(file_name)
            file_formatter = CustomFormatter('%(asctime)s %(levelname)s - %(message)s', datefmt='%H:%M:%S')
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)

    def _debug_section_lv1(self, section_name):
        self.logger.debug("")
        self.logger.debug("=" * 30)
        self.logger.debug(f"{section_name} START")
        self.logger.debug("=" * 30)

    def _debug_section_lv2(self, section_name):
        self.logger.debug("")
        self.logger.debug(">" * 3 + f"{section_name} SUB-SECTION START")
        self.logger.debug(f"{section_name} END")

    def _info_section_lv1(self, section_name):
        self.logger.info("")
        self.logger.info("=" * 30)
        self.logger.info(f"{section_name} START")
        self.logger.info("=" * 30)

    def _info_section_lv2(self, section_name):
        self.logger.info("")
        self.logger.info(">" * 3 + f"{section_name} SUB-SECTION START")
        self.logger.info(f"{section_name} END")

    def _reconfigure_logger(self, debug='critical', debug_file=None):
        self.debug = debug.lower()
        self.debug_file = debug_file
        self.configure_logger()

    cls.__init__ = new_init
    cls.configure_logger = configure_logger
    cls._debug_section_lv1 = _debug_section_lv1
    cls._debug_section_lv2 = _debug_section_lv2
    cls._info_section_lv1 = _info_section_lv1
    cls._info_section_lv2 = _info_section_lv2
    cls._reconfigure_logger = _reconfigure_logger

    return cls
