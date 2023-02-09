import sys
import logging
import locale

locale.setlocale(locale.LC_ALL, 'en_US.UTF8')

# sys.path.append(r'C:\Users\juane\Desktop\Emes despachos\server')
sys.path.append(r'../server')

logging.basicConfig(
    filename='despachos.log',
    format='%(asctime)s - %(message)s',
    datefmt='%d-%b-%y %H:%M:%S'
)
