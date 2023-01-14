import os
import logging
import pandas as pd
from datetime import date, timedelta
import os
import dataframe_image as dfi
import logging
import win32com.client
from PIL import ImageGrab

from datetime import datetime, timedelta
import io
from apiclient import errors
# %%
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from httplib2 import Http
import pandas as pd
from dotenv import load_dotenv

from mypackage.helpers import Helper
from mypackage.constants import SHOP_NAME

from openpyxl import load_workbook, Workbook
import logging
logging.basicConfig(level=logging.INFO)


# ========================================================================================================================================================
# ========================================================================================================================================================
# ========================================================================================================================================================
#declaring constants =====================================================================================================================================

SHOP_NAME = {
    "KS": "kimseng",
    "HYK": "howyekee",
    "KSHYK2": "KSHYK2"
}

SORTER = [
    "Pacific Star - BETA table",
    "Pacific Star - ETA table",
    "Pacific Star - DELTA table",
    "Pacific Star - CAPELLA table",
    "Ryan&Miho - RYAN drop off chairs",
    "Ryan&Miho - MIHO drop off chairs",
    "KK9 - Hall Table",
    "KK1 - Drop Off Table",
    "Dental Fac - Balai Ungku Aziz tables",
    "KK6 - Lobby White Table",
    "Faculty Alam Bina - Foyer",
    "KK2 - Table in front of Mini Mart",
    "Engine Fac - BME Foyer",
    "Sc Fac - New Chemistry Building (Ground Floor)",
    "Sc Fac - Study Room in DK4 (1st floor)",   
    "Business Fac - Foyer of Za'ba Memorial Library",
    "FASS - Entrance Benches",
    "KK3 - Drop Off Table",
    "KK7 - Seats at dewan entrance",
    "KK4 - Rack outside Administration Office",
    "FSKTM - Benches at the Entrance",
    "KK8 - Dewan Seri Mutiara",
    "KK10 - Chairs at the entrance",
    "KK12 - Block D Benches (near Bank Islam ATM machine)",
    "KK12 - Block A Benches (near 2nd entrance gate)",
    "KK5 - Hall Lobby Table (near volleyball court)",
    "SouthView Apartment - Food Drop Off Point",
    "SouthLink Apartment- Food Drop Off Table",
    "Pantai Panorama Block 2 - Drop Off Tables"
]

DRIVER_A = [
    "Pacific Star - BETA table",
    "Pacific Star - ETA table",
    "Pacific Star - DELTA table",
    "Pacific Star - CAPELLA table",
    "Ryan&Miho - RYAN drop off chairs",
    "Ryan&Miho - MIHO drop off chairs",
    "KK9 - Hall Table",
    "KK1 - Drop Off Table",
    "Dental Fac - Balai Ungku Aziz tables",
    "KK6 - Lobby White Table",
    "Faculty Alam Bina - Foyer",
    "KK2 - Table in front of Mini Mart",
    "Engine Fac - BME Foyer",
    "Sc Fac - New Chemistry Building (Ground Floor)",
    "Sc Fac - Study Room in DK4 (1st floor)"
]

DRIVER_B = [
    "Business Fac - Foyer of Za'ba Memorial Library",
    "FASS - Entrance Benches",
    "KK3 - Drop Off Table",
    "KK7 - Seats at dewan entrance",
    "KK4 - Rack outside Administration Office",
    "FSKTM - Benches at the Entrance",
    "KK8 - Dewan Seri Mutiara",
    "KK10 - Chairs at the entrance",
    "KK12 - Block D Benches (near Bank Islam ATM machine)",
    "KK12 - Block A Benches (near 2nd entrance gate)",
    "KK5 - Hall Lobby Table (near volleyball court)",
    "SouthView Apartment - Food Drop Off Point",
    "SouthLink Apartment- Food Drop Off Table",
    "Pantai Panorama Block 2 - Drop Off Tables"
]



#dictionary for name conversion for location list
KK_DICT = {
    "Pacific Star - BETA table"                         :"Pacific BETA",
    "Pacific Star - ETA table"                          :"Pacific ETA",    
    "Pacific Star - DELTA table"                        :"Pacific DELTA",    
    "Pacific Star - CAPELLA table"                      :"Pacific CAPELLA",
    "Ryan&Miho - RYAN drop off chairs"                  :"Ryan",
    "Ryan&Miho - MIHO drop off chairs"                  :"Miho",
    "FASS - Entrance Benches"                           :"FASS",
    "Business Fac - Foyer of Za'ba Memorial Library"    :"Business",
    "KK1 - Drop Off Table"                              :"KK1",
    "Dental Fac - Balai Ungku Aziz tables"              :"Dental",   
    "KK6 - Lobby White Table"                           :"KK6",
    "Faculty Alam Bina - Foyer"                         :"FAB",
    "KK2 - Table in front of Mini Mart"                 :"KK2",
    "Engine Fac - BME Foyer"                            :"Engine",        
    "KK9 - Hall Table"                                  :"KK9",
    "Sc Fac - New Chemistry Building (Ground Floor)"    :"New Chem",
    "Sc Fac - Study Room in DK4 (1st floor)"            :"Study Room",
    "KK3 - Drop Off Table"                              :"KK3", 
    "KK7 - Seats at dewan entrance"                     :"KK7",
    "KK4 - Rack outside Administration Office"          :"KK4",
    "FSKTM - Benches at the Entrance"                   :"FSKTM",
    "KK8 - Dewan Seri Mutiara"                          :"KK8",
    "KK10 - Chairs at the entrance"                     :"KK10",
    "KK12 - Block D Benches (near Bank Islam ATM machine)" :"KK12 D",
    "KK12 - Block A Benches (near 2nd entrance gate)"      :"KK12 A",
    "KK5 - Hall Lobby Table (near volleyball court)"       :"KK5",
    "SouthView Apartment - Food Drop Off Point"            :"sVIEW Apartment",
    "SouthLink Apartment- Food Drop Off Table"             :"sLINK Apartment",
    "Pantai Panorama Block 2 - Drop Off Tables"            :"Panorama Block 2"
}


#dictionary for name conversion for Set Shortform
SET_DICT = {
"Spicy S"            : "S",
"Not Spicy N"        : "N",
"Vegetarian V"       : "V",
"RM9 (F-FKT) Fried Kuew Teow"            : "F-FKT",
"RM9 (F-FKTM) Fried Kuew Teow Mee"       : "F-FKTM",
"RM9 (F-FM) Fried Mee"                   : "F-FM",
"RM9.5 (F-FR) Fried Rice"                : "F-FR",
"RM9.5 (F-SFR) Sambal Fried Rice"        : "F-SFR",
"RM9.5 (P-PMD) Pan Mee Dry"              : "P-PMD",
"RM10 (P-PCPM) Dry Chilli pan mee"       : "P-PCPM",
"RM9.5 (P-YMD) Yee Mee dry"              : "P-YMD",
"RM9 (C-CSCR) Char Siew Chicken Rice"    : "C-CSCR",
"RM9 (C-RPCR) Roasted Pork Chicken Rice" : "C-RPCR",
"RM8 (C-CR) Chicken Rice"                : "C-CR",
"RM3 (B-CS) Char siew Bao"                   : "B-CS",
"RM3 (B-SY) Shang Yoke Bao"                  : "B-SY",
"RM3 (B-RB) Red Bean Bao"                    : "B-RB",
"RM3 (B-L) Lotus Bao"                        : "B-L",
"RM3 (B-K) Kaya Bao"                         : "B-K",
"RM5.5 (B-BB) Big Bao"                   : "B-BB",
"RM5.5 (B-LMK) Lo Mai Kai"               : "B-LMK",
"RM5.5 (B-SM) Siew Mai"                  : "B-SM"
}



SET_SERIES = {
    "S",
    "N",
    "V"
}

F_SERIES = {
   "F-FKT",
   "F-FKTM",
   "F-FM",
   "F-FR",
   "F-SFR"
}

P_SERIES = {
   "P-PMD",
   "P-PCPM",
   "P-YMD"
}

C_SERIES = {
    "C-CSCR",
    "C-RPCR",
    "C-CR"
}

B_SERIES = {
    "B-CS",
    "B-SY",
    "B-RB",
    "B-L",
    "B-K",
    "B-BB",
    "B-LMK",
    "B-SM"
}

ALL_SERIES = {*SET_SERIES, *F_SERIES, *P_SERIES, *C_SERIES, *B_SERIES}

COLOR_DICT = {
    'Pacific Star - BETA table'                       :'#fca503',
    'Pacific Star - ETA table'                        :'#fcc544',
    'Pacific Star - DELTA table'                      :'#92b9f7',
    'Pacific Star - CAPELLA table'                    :'#4388f7',
    'Ryan&Miho - RYAN drop off chairs'                :'#03fcc6',
    'Ryan&Miho - MIHO drop off chairs'                :'#cef2d8',
    'FASS - Entrance Benches'                         :'#fa5ca3',
    'Business Fac - Foyer of Za\'ba Memorial Library' :'#dda1f7',
    'KK1 - Drop Off Table'                            :'#88B04B',
    'Dental Fac - Balai Ungku Aziz tables'            :'#f5e1b0',     
    'KK6 - Lobby White Table'                         :'#F7CAC9',
    'Faculty Alam Bina - Foyer'                       :'#85f2a2',
    'KK2 - Table in front of Mini Mart'               :'#92A8D1',
    'Engine Fac - BME Foyer'                               :'#f7a1a1',
    'KK9 - Hall Table'                                     :'#FF6F61',
    'Sc Fac - New Chemistry Building (Ground Floor)'       :'#f4a1f7',
    'Sc Fac - Study Room in DK4 (1st floor)'               :'#009B77',
    'KK3 - Drop Off Table'                                 :'#f7a1b4',
    'KK7 - Seats at dewan entrance'                        :'#a1f7ae',
    'KK4 - Rack outside Administration Office'             :'#45B8AC',
    'FSKTM - Benches at the Entrance'                      :'#bc95f5',
    'KK8 - Dewan Seri Mutiara'                             :'#EFC050',
    'KK10 - Chairs at the entrance'                        :'#91a6eb',
    'KK12 - Block D Benches (near Bank Islam ATM machine)' :'#DFCFBE',
    'KK12 - Block A Benches (near 2nd entrance gate)'      :'#f6fa8c',
    'KK5 - Hall Lobby Table (near volleyball court)'       :'#f7a1cb',
    'SouthView Apartment - Food Drop Off Point'            :'#f5db7d',
    'SouthLink Apartment- Food Drop Off Table'             :'#a7fc86',
    'Pantai Panorama Block 2 - Drop Off Tables'            :'#65e0ad',
    'default' : '#BAE1FF' # Blue
}

# ========================================================================================================================================================
# ========================================================================================================================================================
# ========================================================================================================================================================

#declaring helpers
def create_folders(shop_name, today_str):
    '''
    Create necessary folder structures.
    '''
    # -Code 
    # -Kimseng
    #  `-Input 
    #    `-Excel files (data)
    #  `-Output
    #    `-11/11/2022... 18/11/2022
    #    `-Image, Order Text
    # -JDinner
    #  `-Input 
    #    `-Excel files (data)
    #  `-Output
    #    `-11/11/2022... 18/11/2022
    #    `-Image, Order Text

    logging.info("Creating Folders if doesn't exist...")
    if not os.path.exists(shop_name):
        os.makedirs(shop_name)

    path_input = os.path.join(shop_name, 'input')
    if not os.path.exists(path_input):
        os.makedirs(path_input)
    
    path_output = os.path.join(shop_name, 'output')
    if not os.path.exists(path_output):
        os.makedirs(path_output)

    path_output_date = os.path.join(path_output, today_str)
    if not os.path.exists(path_output_date):
        os.makedirs(path_output_date)

    logging.info(f"Finish in creating folders for {shop_name}!")

def highlight_rows_js(row):
    '''
    return row with styling format
    '''
    value = row.loc['DINNER Food Pick Up Point']
    color = COLOR_DICT[value]
    
    return ['background-color: {}'.format(color) for _ in row]

def highlight_rows_hyk(row):
    '''
    return row with styling format
    '''
    value = row.loc['HYK Food Pick Up Point']
    color = COLOR_DICT[value]
    
    return ['background-color: {}'.format(color) for _ in row]
    

def highlight_rows_ks(row):
    '''
    return row with styling format
    '''
    value = row.loc['KimSeng Food Pick Up Point']
    color = COLOR_DICT[value]
    
    return ['background-color: {}'.format(color) for _ in row]

class Helper:

    class KS_HYK:
        def output_text_kshyk(self, df_ks, df_hyk, text_name_list):
            pick_up_point_name_ks = "KimSeng Food Pick Up Point"
            pick_up_point_name_hyk = "HYK Food Pick Up Point"
            #create list of locations available today for ks and hyk
            ks_unique_loc = df_ks["KimSeng Food Pick Up Point"].unique()
            hyk_unique_loc = df_hyk[pick_up_point_name_hyk].unique()
            unique_sets = df_hyk['Set'].unique()
            
            #create a combined list of unique locations for both ks and hyk
            kshyk_unique_list = []

            for x in SORTER:
                if (x in hyk_unique_loc) or (x in ks_unique_loc):
                    kshyk_unique_list.append(x)
                    
            #kshyk combined locations output         
            with open(text_name_list[0], 'w', encoding="utf-8") as message_list:    
                for x in kshyk_unique_list:
                    message_list.write(f'{KK_DICT[x]}' + " = " )
                    if x in ks_unique_loc:
                        if df_ks[df_ks[pick_up_point_name_ks]== x]['Packet No.'].min() == df_ks[df_ks[pick_up_point_name_ks]== x]['Packet No.'].max():
                            message_list.write(f"{df_ks[df_ks[pick_up_point_name_ks]== x]['Packet No.'].min()}")
                        else:
                            message_list.write(f"{df_ks[df_ks[pick_up_point_name_ks]== x]['Packet No.'].min()}-{df_ks[df_ks[pick_up_point_name_ks]== x]['Packet No.'].max()}") 

                    if (x in ks_unique_loc) and (x in hyk_unique_loc):
                        message_list.write(" + ")

                    if x in hyk_unique_loc:
                        for set in SET_SERIES:
                            if (((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk['Set'] == set )).sum()) != 0:
                                message_list.write(f'{set}{((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk["Set"] == set)).sum()}' + ' ')
                        
                        for set in F_SERIES:
                            if (((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk['Set'] == set )).sum()) != 0:
                                message_list.write(f'{set}{((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk["Set"] == set)).sum()}' + ' ')

                        for set in P_SERIES:
                            if (((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk['Set'] == set )).sum()) != 0:
                                message_list.write(f'{set}{((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk["Set"] == set)).sum()}' + ' ')

                        for set in C_SERIES:
                            if (((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk['Set'] == set )).sum()) != 0:
                                message_list.write(f'{set}{((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk["Set"] == set)).sum()}' + ' ')
                        
                        for set in B_SERIES:
                            if (((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk['Set'] == set )).sum()) != 0:
                                message_list.write(f'{set}{((df_hyk[pick_up_point_name_hyk] == x)&(df_hyk["Set"] == set)).sum()}' + ' ')
                        
                    message_list.write("\n")
                    
            #part timers Salary
            with open(text_name_list[1], 'w', encoding="utf-8") as part_timer_salary:    

                duo_basic = 21
                lone_basic = 25
                incentive_per_pack = 0.05
                part_timer_salary.write('duo_basic = ' + f'{duo_basic}, ' + 'lone_basic = ' + f'{lone_basic}, ' + 'incentive_per_pack = ' + f'{incentive_per_pack}')
                part_timer_salary.write('\n\n')
                part_timer_salary.write('DRIVER A' + '\n')

                driver_A_ks_count = 0
                for x in DRIVER_A:
                    if (df_ks["KimSeng Food Pick Up Point"] == x).sum() != 0:
                        driver_A_ks_count = driver_A_ks_count + (int((df_ks["KimSeng Food Pick Up Point"] == x).sum()))
                part_timer_salary.write('KS packets = ' + f'{driver_A_ks_count}' + '\n') 

                driver_A_hyk_count = 0
                for x in DRIVER_A:
                    if (df_hyk["HYK Food Pick Up Point"] == x).sum() != 0:
                        driver_A_hyk_count = driver_A_hyk_count + (int((df_hyk["HYK Food Pick Up Point"] == x).sum()))
                part_timer_salary.write('HYK packets = ' + f'{driver_A_hyk_count}' + '\n')   
                driver_A_total_packets = driver_A_ks_count + driver_A_hyk_count
                part_timer_salary.write("TOTAL PACKETS DELIVERED = " + f'{driver_A_total_packets}' + '\n')
                part_timer_salary.write("SALARY = RM" + f'{duo_basic+incentive_per_pack*driver_A_total_packets}' + '\n')
                part_timer_salary.write('\n\n')
                part_timer_salary.write('DRIVER B' + '\n')

                driver_B_ks_count = 0
                for x in DRIVER_B:
                    if (df_ks["KimSeng Food Pick Up Point"] == x).sum() != 0:
                        driver_B_ks_count = driver_B_ks_count + (int((df_ks["KimSeng Food Pick Up Point"] == x).sum()))
                part_timer_salary.write('KS packets = ' + f'{driver_B_ks_count}' + '\n') 
                driver_B_hyk_count = 0
                for x in DRIVER_B:
                    if (df_hyk["HYK Food Pick Up Point"] == x).sum() != 0:
                        driver_B_hyk_count = driver_B_hyk_count + (int((df_hyk["HYK Food Pick Up Point"] == x).sum()))
                part_timer_salary.write('HYK packets = ' + f'{driver_B_hyk_count}' + '\n')   
                driver_B_total_packets = driver_B_ks_count + driver_B_hyk_count
                part_timer_salary.write("TOTAL PACKETS DELIVERED = " + f'{driver_B_ks_count+driver_B_hyk_count}' + '\n')
                part_timer_salary.write("SALARY = RM" + f'{duo_basic+incentive_per_pack*driver_B_total_packets}' + '\n')

                total_packets = driver_A_total_packets + driver_B_total_packets
                part_timer_salary.write('\n\n')
                part_timer_salary.write('LONE DRIVER SALARY = ' + f'{lone_basic + total_packets * incentive_per_pack}' + '\n')
    class KimSeng:
        def process_input_ks(self, df, pick_up_point_name):
            '''
            clean df
            '''
            #https://stackoverflow.com/questions/23482668/sorting-by-a-custom-list-in-pandas/27255567
            # Create the dictionary that defines the order for sorting
            sorter_index = dict(zip(SORTER, range(len(SORTER))))

            # Generate a rank column that will be used to sort the dataframe numerically
            df['Point_Rank'] = df[pick_up_point_name].map(sorter_index)
            
            # Here is the result asked with the lexicographic sort
            # Result may be hard to analyze, so a second sorting is proposed next
            df.sort_values('Point_Rank', ascending = True, inplace = True)
            df.drop('Point_Rank', 1, inplace = True)
            #insert a column of packet numbering
            df.insert(3, "Packet No.", range(1,len(df['Email Address'])+1))  
            #Generate a column for location name conversion
            df['Location'] = df[pick_up_point_name].map(KK_DICT)
            return df

        def output_text_ks(self, txt_name_list, df, pick_up_point_name,):
            # order list
            dishes = df[df.columns[6]].tolist()
            customizes = df[df.columns[7]].fillna(' ').tolist()
            unique_points = df[pick_up_point_name].unique()
            with open(txt_name_list[0], 'w', encoding="utf-8") as order_list:
                # order_list.write("Order List: \n")
                for index, (d, c) in enumerate(zip(dishes, customizes)):
                    order_list.write(f'{index+1}. \n')
                    for dish in d.split(','):
                        dish = dish.strip()
                        order_list.write(dish)
                        order_list.write('\n')

                    for customize in c.split(','):
                        customize = customize.strip()
                        order_list.write(customize + '\n')
                    order_list.write('\n')
                
            with open(txt_name_list[1], 'w', encoding="utf-8") as order_location:
                # order_location.write('-'*50 + '\n')
                # order_location.write("Order Location: \n")
                #for loop min() and max() of index of each location into txt file.    
                for point in unique_points:
                    if df[df[pick_up_point_name]== point]['Packet No.'].min() == df[df[pick_up_point_name]== point]['Packet No.'].max():
                        order_location.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()}={KK_DICT[point]}"+"\n")
                    else:
                        order_location.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()}-{df[df[pick_up_point_name]== point]['Packet No.'].max()}={KK_DICT[point]}"+"\n")
                
        def generate_pickup_image_ks(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            
            df_style = df.style.apply(highlight_rows_ks, axis=1)

            #Export as excel
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            #format for the column width, .set_column(index_start, index_end, width)
            # writer.sheets['Sheet1'].set_column(3, 3, 5)
            # writer.sheets['Sheet1'].set_column(4, 4, 30)
            # writer.sheets['Sheet1'].set_column(5, 5, 40)
            # writer.sheets['Sheet1'].set_column(6, 6, 45)

            writer.close()

            o = win32com.client.Dispatch('Excel.Application')
            wb = o.Workbooks.Open(os.path.join(os.getcwd(), excel_output_path))
            ws = wb.Worksheets['Sheet1']

            ws.Range(ws.Cells(1,1),ws.Cells(df.shape[0],df.shape[1])).CopyPicture(Format=2)

            img = ImageGrab.grabclipboard()
            img.save(image_output_path)
            wb.Close(True)

            logging.info(f"Image generated in {image_output_path}!")

        def run_ks(self, df_input, date_str):
            shop_name = SHOP_NAME['KS'] #or 'JS'
            pick_up_point_name = "KimSeng Food Pick Up Point"
            excel_output_path = f"{date_str} orders.xlsx"
            image_output_path = f"{date_str} image.jpg"
            txt_name_list = ['KS Order List.txt', 'KS Order Location.txt']
            df = self.process_input_ks(df_input, pick_up_point_name)
            # display(df_input)
            self.output_text_ks(txt_name_list, df, pick_up_point_name) ##print text
            self.generate_pickup_image_ks(df, excel_output_path, image_output_path)

            return excel_output_path, image_output_path, txt_name_list

    class HYK:
        def process_input_hyk(self, df, pick_up_point_name_hyk):
            '''
            clean df
            '''
            #https://stackoverflow.com/questions/23482668/sorting-by-a-custom-list-in-pandas/27255567
            # Create the dictionary that defines the order for sorting
            sorter_index = dict(zip(SORTER, range(len(SORTER))))
            menu_name = f'Menu (Select One)'

            # Generate a rank column that will be used to sort the dataframe numerically
            df['Point_Rank'] = df[pick_up_point_name_hyk].map(sorter_index)
            
            # Here is the result asked with the lexicographic sort
            # Result may be hard to analyze, so a second sorting is proposed next
            df.sort_values(['Point_Rank','Menu (Select One)'], 
                        ascending = [True, True], inplace = True)
            df.drop('Point_Rank', 1, inplace = True)
            
            #insert a column of packet numbering
            df.insert(3, "No.", range(1,len(df['Name'])+1))  
            
            #create a list of all unique locations.
            unique_points_hyk = df['HYK Food Pick Up Point'].unique()
            
            #change Menu name containing substring to short form
            #https://stackoverflow.com/questions/39768547/replace-whole-string-if-it-contains-substring-in-pandas
            df.loc[df['Menu (Select One)'].str.contains('Spicy Set S'), 'Menu (Select One)'] = 'Spicy S'
            df.loc[df['Menu (Select One)'].str.contains('Not Spicy Set N'), 'Menu (Select One)'] = 'Not Spicy N'
            df.loc[df['Menu (Select One)'].str.contains('Vegetarian Set V'), 'Menu (Select One)'] = 'Vegetarian V'
            
            # Generate a column for location name conversion
            df['Location'] = df['HYK Food Pick Up Point'].map(KK_DICT)
            # Generate a column for shortform of each set conversion
            df['Set'] = df['Menu (Select One)'].map(SET_DICT)
            return df
    

        def output_text_hyk(self, df, pick_up_point_name, txt_name_list):
            #create a list of all unique locations.
            unique_locations_hyk = df['Location'].unique()
            #create a list of all unique Menu.
            unique_flavours = df['Menu (Select One)'].unique()
            #create a list of all unique set shortform
            unique_sets = df['Set'].unique()


            with open(txt_name_list[0], 'w') as hyk_chef_note:
                # hyk_chef_note.write('HYK Chef - Total Food To Prepare.txt' + '\n')
                for ss in SET_SERIES:
                    if ss in list(df['Set']):
                        hyk_chef_note.write(f'{ss}'.ljust(8) + f'{(df["Set"] == ss).sum()}' + '\n')
                
                if len(F_SERIES.intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'F_SERIES' + '\n')
                
                for fs in F_SERIES:
                    if fs in list(df['Set']):
                        hyk_chef_note.write(f'{fs}'.ljust(8) + f'{(df["Set"] == fs).sum()}' + '\n')
                            
                if len(P_SERIES.intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'P_SERIES' + '\n')
                
                for ps in P_SERIES:
                    if ps in list(df['Set']):
                        hyk_chef_note.write(f'{ps}'.ljust(8) + f'{(df["Set"] == ps).sum()}' + '\n')
                            
                if len(C_SERIES.intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'C_SERIES' + '\n')
                
                for cs in C_SERIES:
                    if cs in list(df['Set']):
                        hyk_chef_note.write(f'{cs}'.ljust(8) + f'{(df["Set"] == cs).sum()}' + '\n')
                                    
                if len(B_SERIES.intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'B_SERIES' + '\n')
                
                for bs in B_SERIES:
                    if bs in list(df['Set']):
                        hyk_chef_note.write(f'{bs}'.ljust(8) + f'{(df["Set"] == bs).sum()}' + '\n')

            with open(txt_name_list[1], 'w') as hyk_packaging_list:
                index = 0
                # hyk_packaging_list.write('HYK Packaging List.txt' + '\n')
                for location in unique_locations_hyk:
                    index += 1
                    hyk_packaging_list.write(f'{index}.'.ljust(4))
                    for ss in SET_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == ss )).sum()) != 0:
                            hyk_packaging_list.write(f'{ss}{((df["Location"] == location)&(df["Set"] == ss)).sum()}'+"  ")
                            
                    for fs in F_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == fs )).sum()) != 0:
                            hyk_packaging_list.write(f'{fs}{((df["Location"] == location)&(df["Set"] == fs)).sum()}'+"  ")
                            
                    for ps in P_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == ps )).sum()) != 0:
                            hyk_packaging_list.write(f'{ps}{((df["Location"] == location)&(df["Set"] == ps)).sum()}'+"  ")
                            
                    for cs in C_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == cs )).sum()) != 0:
                            hyk_packaging_list.write(f'{cs}{((df["Location"] == location)&(df["Set"] == cs)).sum()}'+"  ")
                            
                    for bs in B_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == bs )).sum()) != 0:
                            hyk_packaging_list.write(f'{bs}{((df["Location"] == location)&(df["Set"] == bs)).sum()}'+"  ")
                    
                    hyk_packaging_list.write('\n\n')

            with open(txt_name_list[2], 'w') as backup_hyk_runner_location_list:
                # backup_hyk_runner_location_list.write('backup HYK Runner - Location List.txt' + '\n')
                for location in unique_locations_hyk:
                    backup_hyk_runner_location_list.write(f'{location} = ')
                    for ss in SET_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == ss )).sum()) != 0:
                            backup_hyk_runner_location_list.write(f'{ss}{((df["Location"] == location)&(df["Set"] == ss)).sum()}'+"  ")
                            
                    for fs in F_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == fs )).sum()) != 0:
                            backup_hyk_runner_location_list.write(f'{fs}{((df["Location"] == location)&(df["Set"] == fs)).sum()}'+"  ")
                            
                    for ps in P_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == ps )).sum()) != 0:
                            backup_hyk_runner_location_list.write(f'{ps}{((df["Location"] == location)&(df["Set"] == ps)).sum()}'+"  ")
                            
                    for cs in C_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == cs )).sum()) != 0:
                            backup_hyk_runner_location_list.write(f'{cs}{((df["Location"] == location)&(df["Set"] == cs)).sum()}'+"  ")
                            
                    for bs in B_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == bs )).sum()) != 0:
                            backup_hyk_runner_location_list.write(f'{bs}{((df["Location"] == location)&(df["Set"] == bs)).sum()}'+"  ")
                            
                    backup_hyk_runner_location_list.write('\n\n')

        def generate_pickup_image_ks(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            
            df_style = df.style.apply(highlight_rows_ks, axis=1)

            #Export as excel
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            writer.close()

            o = win32com.client.Dispatch('Excel.Application')
            wb = o.Workbooks.Open(os.path.join(os.getcwd(), excel_output_path))
            ws = wb.Worksheets['Sheet1']

            ws.Range(ws.Cells(1,1),ws.Cells(df.shape[0],df.shape[1])).CopyPicture(Format=2)

            img = ImageGrab.grabclipboard()
            img.save(image_output_path)
            wb.Close(True)

            logging.info(f"Image generated in {image_output_path}!")
                
        def generate_pickup_image_hyk(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            
            #for excel
            df_style = df.style.apply(highlight_rows_hyk, axis=1)
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            writer.close()

            o = win32com.client.Dispatch('Excel.Application')
            wb = o.Workbooks.Open(os.path.join(os.getcwd(), excel_output_path))
            ws = wb.Worksheets['Sheet1']

            ws.Range(ws.Cells(1,2),ws.Cells(df.shape[0]+1,df.shape[1])).CopyPicture(Format=2)

            img = ImageGrab.grabclipboard()
            img.save(image_output_path)
            wb.Close(True)

            logging.info(f"Image generated in {image_output_path}!")

        def run_hyk(self, df_input, date_str):
            shop_name = SHOP_NAME['KSHYK2'] #or 'HYK' or 'KS'
            pick_up_point_name = "HYK Food Pick Up Point"
            excel_output_path = f"{date_str} orders.xlsx"
            image_output_path = f"{date_str} image.jpg"
            txt_name_list = ['HYK Total Food To Prepare.txt','HYK Packaging List.txt', 'backup HYK Runner - Location List.txt']

            df = self.process_input_hyk(df_input, pick_up_point_name)
            # display(df)
            self.output_text_hyk(df, pick_up_point_name, txt_name_list) ##print text
            self.generate_pickup_image_hyk(df, excel_output_path, image_output_path)
            return excel_output_path, image_output_path, txt_name_list

# ========================================================================================================================================================
# ========================================================================================================================================================
# ========================================================================================================================================================

def create_excel_input(service_form, service_sheet, formId):
    '''
    Retrieve google sheet and create csv locally
    '''
    # Retrieve google sheet
    form = service_form.forms().get(formId=formId).execute()
    sheet_res = service_sheet.spreadsheets().get(spreadsheetId=form['linkedSheetId']).execute()
    dataset = service_sheet.spreadsheets().values().get(
        spreadsheetId= sheet_res['spreadsheetId'],
        range=sheet_res['sheets'][0]['properties']['title'],
        majorDimension= 'ROWS',
    ).execute()
    df = pd.DataFrame(dataset['values'])
    df = df.rename(columns=df.iloc[0]).drop(df.index[0])
    columns_name = list(df.columns)
    column_d, column_e, column_f = columns_name[3], columns_name[4], columns_name[5]
    if df.duplicated(subset=[column_d, column_e, column_f], keep=False).any() == True:
        df_duplicated = df.loc[df.duplicated()]
        file_name_duplicated = f"{sheet_res['properties']['title'].replace('/', '-')}_duplicated.csv"
        df_duplicated.to_csv(file_name_duplicated, index=False)
        df.drop_duplicates(subset=[column_d, column_e, column_f], keep='first', inplace=True)
    else:
        file_name_duplicated="None"

    file_name_real = f"{sheet_res['properties']['title'].replace('/', '-')}_Generated.csv"
    df.to_csv(file_name_real, index=False)
    if file_name_duplicated != "None":
        return df, [file_name_real, file_name_duplicated]
    return df, [file_name_real]

def create_all_output(kimseng_helper_obj, hyk_helper_obj, today_str, df_input, is_kimseng):
    if is_kimseng:
        excel_output_path, image_output_path, txt_name_list = kimseng_helper_obj.run_ks(df_input, today_str)
        return excel_output_path, image_output_path, txt_name_list
    else:
        excel_output_path, image_output_path, txt_name_list = hyk_helper_obj.run_hyk(df_input, today_str)
        return excel_output_path, image_output_path, txt_name_list

def get_or_create_folder_id(service_drive, file_name, parents=None, is_root=False):
    '''
    file_name: File's Name
    parents: [File Id]
    '''
    file_id = ''
    # if is_root:
    files = service_drive.files().list( q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{parents[0]}' " ,
                                    spaces='drive').execute()['files']
    # else:
    #   files = service_drive.files().list(q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false" ,
    #                                    spaces='drive').execute()['files']
    if len(files) > 0: #exist
        print(f"{file_name} exist. ")
        file_id = files[0].get('id')
        return file_id
    #create folder
    file_metadata = {
            'name': f"{file_name}",
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': parents, 
        }

    res_file = service_drive.files().create(body=file_metadata, supportsAllDrives=True, fields='id').execute()
    file_id = res_file.get('id')
    # service_drive.permissions().create(fileId = file_id, body={"role": "reader", "type":"user", "emailAddress": "byichonggoh@gmail.com"}, supportsAllDrives=True, fields='id').execute()
    return file_id

def file_exist(service_drive, file_name, mimetype):
    files = service_drive.files().list( q=f"name = '{file_name}' and mimeType = '{mimetype}' and trashed = false" ,
                                       spaces='drive',).execute()['files']
    if len(files) > 0: #exist
        return files[0].get('id')
    return False

def upload_input_file(service_drive, file_name, parents):
    '''
    Input:
    file_name: File's Name
    parents: [File Id]
    '''
    file_metadata = {'name': file_name,'parents': parents,}
    file_media = MediaFileUpload(file_name, mimetype='text/csv')
    # if file_id := file_exist(service_drive, file_name, 'text/csv'):
    #     service_drive.files().update(fileId=file_id, body={'name': file_name}, media_body=file_media, fields='id').execute()['id']
    # else:
    file_id = service_drive.files().create(body=file_metadata, media_body=file_media, fields='id').execute()['id']
    file_media = None #To stop reading the file to allow delete

def upload_output_file(service_drive, parents, excel_name="", image_name="", txt_name_list=[]):
    '''
    Output
    file_name: File's Name
    parents: [File Id]
    '''

    if image_name != "":
        image_metadata = {'name': image_name, 'parents': parents,}
        image_media = MediaFileUpload(image_name, mimetype='image/png')
        # if image_file_id := file_exist(image_name, 'image/png'):
        #     service_drive.files().update(fileId= image_file_id, body={'name': image_name}, media_body=image_media, fields='id').execute()['id']
        # else:
        image_file_id = service_drive.files().create(body=image_metadata, media_body=image_media, fields='id').execute()['id']
        image_media=None
    
    if txt_name_list != []:
        for txt_name in txt_name_list:
            text_metadata = {'name': txt_name,'parents': parents,}
            text_media = MediaFileUpload(txt_name, mimetype='text/plain')
            # if text_file_id := file_exist(txt_name, 'text/plain'):
            #     service_drive.files().update(fileId=text_file_id, body={'name': txt_name}, media_body=text_media, fields='id').execute()['id']
            # else:
            text_file_id = service_drive.files().create(body=text_metadata, media_body=text_media, fields='id').execute()['id']
            text_media=None

    if excel_name != "":
        excel_metadata = {'name': excel_name,'parents': parents}
        excel_media = MediaFileUpload(excel_name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # if excel_file_id := file_exist(excel_name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
        #     service_drive.files().update(fileId=excel_file_id, body={'name': excel_name}, media_body=excel_media, fields='id').execute()['id']
        # else:
        excel_file_id = service_drive.files().create(body=excel_metadata, media_body=excel_media, fields='id').execute()['id']
        excel_media=None

def run_program(service_form, service_sheet, kimseng_helper_obj, hyk_helper_obj, today_str, service_drive, form_id, folder_input_id, folder_output_today_id, is_kimseng=False):
    df_input, excel_input_name_list = create_excel_input(service_form, service_sheet, form_id)
    logging.info(f"Created Excel Input file list!")
    if len(excel_input_name_list) > 1:
        #this is the duplicate file if exist
        logging.info(f"There's duplicated record. Creating a file for that...")
        upload_input_file(service_drive, excel_input_name_list[1], [folder_output_today_id])
        logging.info(f"Uploaded excel input duplicated file to output folder!")
        os.remove(excel_input_name_list[1])
    upload_input_file(service_drive, excel_input_name_list[0], [folder_input_id])
    logging.info(f"Uploaded excel input generated file to input folder!")
    os.remove(excel_input_name_list[0])
    
    excel_output_path, image_output_path, txt_name_list = create_all_output(kimseng_helper_obj, hyk_helper_obj, today_str, df_input, is_kimseng)
    upload_output_file(service_drive, [folder_output_today_id], excel_name=excel_output_path, image_name=image_output_path, txt_name_list=txt_name_list)
    logging.info(f"Uploaded all the output files!")
    os.remove(excel_output_path)
    os.remove(image_output_path)
    [os.remove(txt_name) for txt_name in txt_name_list]
    return df_input

def run():
    #setup
    logging.info("Loading local environment...")
    load_dotenv()

    SERVICE_ACCOUNT = os.getenv('SERVICE_ACCOUNT') #The service acc used to create files.
    GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") #the shared google sheet that user can access and input ID
    SHARED_PARENT_FOLDER_ID = os.getenv("SHARED_PARENT_FOLDER_ID") #the shared parent folder on personal acc

    SCOPES = ["https://www.googleapis.com/auth/forms.body", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/forms.responses.readonly", "https://www.googleapis.com/auth/spreadsheets.readonly"]

    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT, scopes=SCOPES)
    ks_hyk_helper_obj = Helper.KS_HYK()
    kimseng_helper_obj = Helper.KimSeng()
    hyk_helper_obj = Helper.HYK()

    logging.info("Building Google's Credentials...")
    #TODAY_STR
    today_str = (datetime.today() + timedelta(days=1)).strftime('%d-%m-%Y') #today_str = date that file should be run. Hence if today is 04-01-2023, then today str should be 05-01-2023 so that 05-01-2023 row will be read.
    today_str = "10-01-2023" #today_str = date that file should be run. Hence if today is 04-01-2023, then today str should be 05-01-2023 so that 05-01-2023 row will be read.
    service_form = build('forms', 'v1', credentials=credentials) #forms
    service_sheet = build('sheets', 'v4', credentials=credentials)
    service_drive = build('drive', 'v3', credentials=credentials)

    #create stuff
    #so the idea here is to export to files, upload the files, and delete the files. can be triggered through the form id
    logging.info("Creating/Getting all folder ids...")
    ks_j_hyk_folder_id = get_or_create_folder_id(service_drive, 'KS&J&HYK', [SHARED_PARENT_FOLDER_ID], is_root=True)
    kimseng_folder_id = get_or_create_folder_id(service_drive, 'KimSeng', [ks_j_hyk_folder_id])
    kimseng_folder_input_id = get_or_create_folder_id(service_drive, 'KimSengInput', [kimseng_folder_id])
    kimseng_folder_output_id = get_or_create_folder_id(service_drive, 'KimSengOutput', [kimseng_folder_id])
    kimseng_folder_output_today_id = get_or_create_folder_id(service_drive, f'{today_str} Koutput', [kimseng_folder_output_id])

    hyk_folder_id = get_or_create_folder_id(service_drive, 'HYK', [ks_j_hyk_folder_id])
    hyk_folder_input_id = get_or_create_folder_id(service_drive, 'HYKInput', [hyk_folder_id])
    hyk_folder_output_id = get_or_create_folder_id(service_drive, 'HYKOutput', [hyk_folder_id])
    hyk_folder_output_today_id = get_or_create_folder_id(service_drive, f'{today_str} HYKOutput', [hyk_folder_output_id])

    logging.info("Getting Google Sheet ID information...")
    #get sheet list
    sheet_res = service_sheet.spreadsheets().get(spreadsheetId=GOOGLE_SHEET_ID).execute()
    dataset = service_sheet.spreadsheets().values().get(
            spreadsheetId= sheet_res['spreadsheetId'],
            range=sheet_res['sheets'][0]['properties']['title'],
            majorDimension= 'ROWS',
        ).execute()
    google_sheet_df = pd.DataFrame(dataset['values'])
    google_sheet_df = google_sheet_df.rename(columns=google_sheet_df.iloc[0]).drop(google_sheet_df.index[0])
    google_sheet_df = google_sheet_df[google_sheet_df['Date'] == today_str]

    logging.info("Running Kim Seng program...")
    # Kimseng
    ks_df_input = hyk_df_input = pd.DataFrame()
    kimseng_form_id = google_sheet_df["KimSeng"].item()
    if kimseng_form_id != "None":
        ks_df_input = run_program(service_form, service_sheet, kimseng_helper_obj, hyk_helper_obj, today_str, service_drive, kimseng_form_id, kimseng_folder_input_id, kimseng_folder_output_today_id, is_kimseng=True)
    else:
        logging.info(f"Kim Seng Google Form ID is none!")

    logging.info("Running HYK Program...")
    #hyk
    hyk_form_id = google_sheet_df["HYK"].item()
    if hyk_form_id != "None":
        hyk_df_input = run_program(service_form, service_sheet, kimseng_helper_obj, hyk_helper_obj, today_str, service_drive, hyk_form_id, hyk_folder_input_id, hyk_folder_output_today_id, is_kimseng=False)
    else:
        logging.info(f"HYK Google Form ID is none!")

    if not ks_df_input.empty and not hyk_df_input.empty:
        #kshyk combined text
        logging.info("KS and HYK df are both not empty. Generating txt files...")
        kshyk_text_name_list=['KSHYK driver.txt', 'Parttimer Salary.txt']
        ks_hyk_helper_obj.output_text_kshyk(ks_df_input, hyk_df_input, kshyk_text_name_list)
        upload_output_file(service_drive, [kimseng_folder_output_today_id], txt_name_list=kshyk_text_name_list)
        upload_output_file(service_drive, [hyk_folder_output_today_id], txt_name_list=kshyk_text_name_list)
        [os.remove(txt_name) for txt_name in kshyk_text_name_list]

        #calculations
        FILE_NAME = 'Kimseng Calculator.xlsx'
        OUTPUT_FILE_NAME = f'{today_str} {FILE_NAME}'

        #get from drive
        logging.info(f"Calculating cost using Kimseng Calculator...")
        file_id = service_drive.files().list(q=f"name='KimSeng Calculator.xlsx' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'", spaces='drive',).execute()['files'][0]['id']
        request = service_drive.files().get_media(fileId=file_id)
        try:
            file = io.BytesIO()
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            logging.info(f'Download {int(status.progress() * 100)}.')
        except HttpError as error:
            logging(f'An error occurred: {error}')
            file = None
        
        if file != None:
            xlsx = io.BytesIO(file.getvalue())
            df_acc = pd.read_excel(xlsx, sheet_name='Sheet1')

            #start calculation
            hyk_quantities= df_acc.apply(lambda row: list(hyk_df_input['Set']).count(row['Code']) , axis=1)
            kimseng_quantity= ks_df_input['Email Address'].count()

            workbook_data_only = load_workbook(filename= FILE_NAME, data_only=False)
            ws_data = workbook_data_only.get_sheet_by_name('Sheet1') #initial data

            for index, quantity in enumerate(hyk_quantities):
                ws_data[f'B{index+2}'] = quantity

            ws_data[f'C2'] = kimseng_quantity
            workbook_data_only.save(OUTPUT_FILE_NAME)

            upload_output_file(service_drive, [kimseng_folder_output_today_id], excel_name=OUTPUT_FILE_NAME)
            upload_output_file(service_drive, [hyk_folder_output_today_id], excel_name=OUTPUT_FILE_NAME)
            os.remove(OUTPUT_FILE_NAME)
            logging.info(f"Calculation Done!")
        else:
            logging.info(f"No KimSeng Calculator found! exiting...")

if __name__ == '__main__':
    run()

