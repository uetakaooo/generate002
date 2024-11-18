import logging
import os

logger = logging.getLogger('my_log')
logger.setLevel(logging.DEBUG)

console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console.setFormatter(format)
logger.addHandler(console)

error_handler = logging.FileHandler('error.log',mode='w',encoding='utf-8')
error_handler.setLevel(logging.ERROR)
error_handler.setFormatter(format)
logger.addHandler(error_handler)