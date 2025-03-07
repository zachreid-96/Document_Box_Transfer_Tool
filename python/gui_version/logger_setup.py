import logging
import os
import sys
from datetime import datetime


def setup():
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    log_file = os.path.join(exe_dir, datetime.now().strftime("%m-%d-%Y_%H-%M_runtime.log"))

    logger = logging.getLogger("document_box_logger")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
        ))
        logger.addHandler(file_handler)

    return logger
