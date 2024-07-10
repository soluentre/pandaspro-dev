import logging
from datetime import datetime
from functools import wraps


def cpd_logger(cls):
    class CustomFormatter(logging.Formatter):
        def format(self, record):
            if record.msg == "\n":
                return "\n"  # 纯空行直接返回换行符
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
        self.log_section = self._log_section
        self.end_log_section = self._end_log_section
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

    def _log_section(self, section_name):
        self.logger.debug("\n")
        self.logger.debug("=" * 30)
        self.logger.debug(f"{section_name} START")
        self.logger.debug("=" * 30)

    def _end_log_section(self, section_name):
        self.logger.debug("=" * 30)
        self.logger.debug(f"{section_name} END")
        self.logger.debug("=" * 30)
        self.logger.debug("\n")

    def _reconfigure_logger(self, debug='critical', debug_file=None):
        self.debug = debug.lower()
        self.debug_file = debug_file
        self.configure_logger()

    cls.__init__ = new_init
    cls.configure_logger = configure_logger
    cls._log_section = _log_section
    cls._end_log_section = _end_log_section
    cls._reconfigure_logger = _reconfigure_logger

    return cls
