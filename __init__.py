import logging
import locale

locale.setlocale(locale.LC_ALL, 'en_US.UTF8')

logging.basicConfig(
    filename='despachos.log',
    format='%(asctime)s - %(message)s',
    datefmt='%d-%b-%y %H:%M:%S'
)