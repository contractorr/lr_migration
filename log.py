import logging 

def get_logger(name):    
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)   
    return logger   

def update_handler(logger, file_name):
    while logger.hasHandlers():
        logger.removeHandler(logger.handlers[0])

    # decide on format
    formatter = logging.Formatter('L: %(lineno)d F: %(funcName)s \t\t %(message)s')

    # add file handler 
    file_handler = logging.FileHandler(f'{file_name}.log')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


