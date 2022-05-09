'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.06

Library of tools used to manage logging
'''

import logging
import logging.config

import os
import sys
from logging.handlers import TimedRotatingFileHandler

FORMATTER = logging.Formatter("%(asctime)s - %(levelname)s - %(name)s:%(lineno)d - %(message)s")
LOG_FILE = "my_app.log"


# Add to your code
# my_logger = get_logger("my module name")
# my_logger.debug("a debug message")

def get_console_handler():
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(FORMATTER)
    return console_handler


def get_file_handler(logfile=LOG_FILE):
    file_handler = TimedRotatingFileHandler(LOG_FILE, when='midnight')
    file_handler.setFormatter(FORMATTER)
    return file_handler


def get_logger(logger_name, logfile=LOG_FILE, loggerlevel=None):
    logger = logging.getLogger(logger_name)

    if loggerlevel is not None:
        logger.setLevel(loggerlevel)  # better to have too much log than not enough

    logger.addHandler(get_console_handler())
    logger.addHandler(get_file_handler(logfile))

    # with this pattern, it's rarely necessary to propagate the error up to parent
    logger.propagate = False

    return logger


def get_config(log_path=LOG_FILE,
               fhandler='logging.handlers.RotatingFileHandler',
               loggerlevel=None,
               maxBytes=None):
    if maxBytes is None:
        maxBytes = 1024 * 1000 * 100

    if loggerlevel is None:
        loggerlevel = 'DEBUG'

    config = {
        'disable_existing_loggers': False,
        'version': 1,
        'formatters': {
            'default': {
                'format': '%(asctime)s %(levelname)s %(name)s:%(lineno)d %(funcName)s %(message)s',
            },
            'short': {
                'format': '%(asctime)s %(levelname)s %(name)s %(levelname)s:%(lineno)d: %(message)s'
            },
        },
        'handlers': {
            'console': {
                'level': 'INFO',
                'formatter': 'default',
                'class': 'logging.StreamHandler',
                'stream': 'ext://sys.stdout'
            },
            'file': {
                'level': loggerlevel,
                'class': fhandler,
                'formatter': 'default',
                'filename': log_path,
                'encoding': 'UTF-8',
                'maxBytes': maxBytes,
                'backupCount': 3
            }
        },
        'loggers': {
            '': {
                'handlers': ['console', 'file'],
                'level': loggerlevel,
            },
        },
    }
    if fhandler == 'logging.handlers.TimedRotatingFileHandler':
        config['handlers']['file']['when'] = 'midnight'
        # config['handlers']['file']['interval'] = 1
        config['handlers']['file']['backupCount'] = 31
        del config['handlers']['file']['maxBytes']
    elif fhandler == 'logging.FileHandler':
        # config['handlers']['file']['mode'] = 'a'
        # config['handlers']['file']['delay'] = False
        del config['handlers']['file']['maxBytes']
        del config['handlers']['file']['backupCount']

    return config


def setHandlerLevel(dictConfig, handlerType, level):
    dictConfig['handlers'][handlerType] = level


def dictConfig(config):
    logging.config.dictConfig(config)


def getLogger(name):
    return logging.getLogger(name)


def clear_logs(config, logger):
    log_file = config['handlers']['file']['filename']
    logging.shutdown()
    os.remove(log_file)
    dictConfig(config)
    logger.info('Logs cleared at startup: %s', log_file)


"""
Commented out this code block

# Capturing Traceback informatoin in your logs and JSON payload logging pointers
# https://www.datadoghq.com/blog/python-logging-best-practices/

# added logging feature to capture and log unhandled exceptions
def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception

"""
