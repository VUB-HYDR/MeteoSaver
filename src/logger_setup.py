import logging
import os

def setup_logger(log_dir, log_filename='meteosaver.log'):
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, log_filename)

    logger = logging.getLogger('MeteoSaver')
    logger.setLevel(logging.INFO)

    # Formatter
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')

    # Console handler
    ch = logging.StreamHandler()
    ch.setFormatter(formatter)

    # File handler
    fh = logging.FileHandler(log_path)
    fh.setFormatter(formatter)

    # Avoid duplicated logs
    if not logger.handlers:
        logger.addHandler(ch)
        logger.addHandler(fh)

    return logger
