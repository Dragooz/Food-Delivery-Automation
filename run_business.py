#import packages
import logging
from datetime import date

#mine
from mypackage.helpers import Helper
from mypackage.constants import *

#Setup
logging.getLogger().setLevel(logging.INFO)

# today_str = date.today().strftime('%d_%m_%Y')
today_str = date.today().strftime('31_3_2022')

def run():
    # print("SORTER: ", SORTER)
    Helper.create_folders(today_str) ##create folder structure

run()
