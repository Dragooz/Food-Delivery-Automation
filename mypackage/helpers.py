import os
import logging
from mypackage.constants import COLOR_DICT, SORTER, KK_DICT, SHOP_NAME, DRIVER_A, DRIVER_B, SET_DICT, SET_SERIES, F_SERIES, P_SERIES, C_SERIES, B_SERIES, DUO_BASIC,  LONE_BASIC, INCENTIVE_PER_PACK
#import packages
import pandas as pd
from datetime import date, timedelta
import os
import dataframe_image as dfi
import logging
import win32com.client
from PIL import ImageGrab

#Setup
pd.set_option('display.max_rows', 1000)

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

            
                part_timer_salary.write('DUO_BASIC = ' + f'{DUO_BASIC}, ' + 'LONE_BASIC = ' + f'{LONE_BASIC}, ' + 'INCENTIVE_PER_PACK = ' + f'{INCENTIVE_PER_PACK}')
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
                part_timer_salary.write("SALARY = RM" + f'{DUO_BASIC+INCENTIVE_PER_PACK*driver_A_total_packets}' + '\n')
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
                part_timer_salary.write("SALARY = RM" + f'{DUO_BASIC+INCENTIVE_PER_PACK*driver_B_total_packets}' + '\n')

                total_packets = driver_A_total_packets + driver_B_total_packets
                part_timer_salary.write('\n\n')
                part_timer_salary.write('LONE DRIVER SALARY = ' + f'{LONE_BASIC + total_packets * INCENTIVE_PER_PACK}' + '\n')
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

            with open(txt_name_list[2], 'w', encoding="utf-8") as order_packaging:
                for point in unique_points:
                    if df[df[pick_up_point_name]== point]['Packet No.'].min() == df[df[pick_up_point_name]== point]['Packet No.'].max():
                        order_packaging.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()} ({df[df[pick_up_point_name]== point]['Packet No.'].max()-df[df[pick_up_point_name]== point]['Packet No.'].min()+1} 包)"+"\n")
                    else:
                        order_packaging.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()}-{df[df[pick_up_point_name]== point]['Packet No.'].max()} ({df[df[pick_up_point_name]== point]['Packet No.'].max()-df[df[pick_up_point_name]== point]['Packet No.'].min()+1} 包)"+"\n")


        def generate_pickup_image_ks(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            df_copy = df.copy()
            df_copy = df_copy.iloc[:, 3:6]
            df_style = df_copy.style.apply(highlight_rows_ks, axis=1)

            #Export as excel
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            #format for the column width, .set_column(index_start, index_end, width)
            writer.sheets['Sheet1'].set_column("A:A", 10)
            writer.sheets['Sheet1'].set_column("B:B", 30)
            writer.sheets['Sheet1'].set_column("C:C", 40)
            # writer.sheets['Sheet1'].set_column(6, 6, 45)

            writer.close()

            o = win32com.client.Dispatch('Excel.Application')
            wb = o.Workbooks.Open(os.path.join(os.getcwd(), excel_output_path))
            ws = wb.Worksheets['Sheet1']

            ws.Range(ws.Cells(1,1),ws.Cells(df_copy.shape[0]+1,df_copy.shape[1])).CopyPicture(Format=2)

            img = ImageGrab.grabclipboard()
            img.save(image_output_path)
            wb.Close(True)

            logging.info(f"Image generated in {image_output_path}!")

        def run_ks(self, df_input, date_str):
            shop_name = SHOP_NAME['KS'] #or 'JS'
            pick_up_point_name = "KimSeng Food Pick Up Point"
            excel_output_path = f"KimSeng {date_str} orders.xlsx"
            image_output_path = f"KimSeng {date_str} image.jpg"
            txt_name_list = ['KimSeng Order List.txt', 'KimSeng backup Delivery Location.txt', 'KimSeng Packaging.txt']
            df = self.process_input_ks(df_input, pick_up_point_name)
            # display(df_input)
            self.output_text_ks(txt_name_list, df, pick_up_point_name) ##print text
            self.generate_pickup_image_ks(df, excel_output_path, image_output_path)

            return excel_output_path, image_output_path, txt_name_list

    # class Jackson:
    #     def process_input_js(self, df):
    #         '''
    #         clean df
    #         '''
    #         #https://stackoverflow.com/questions/23482668/sorting-by-a-custom-list-in-pandas/27255567
    #         # Create the dictionary that defines the order for sorting
    #         sorter_index = dict(zip(SORTER, range(len(SORTER))))
    #         pick_up_point_name = "DINNER Food Pick Up Point"

    #         # Generate a rank column that will be used to sort the dataframe numerically
    #         df['Point_Rank'] = df[pick_up_point_name].map(sorter_index)
            
    #         # Here is the result asked with the lexicographic sort
    #         # Result may be hard to analyze, so a second sorting is proposed next
    #         df.sort_values(['Point_Rank','Menu (Select One)'], ascending = [True, True], inplace = True)
    #         df.drop('Point_Rank', 1, inplace = True)
    #         #insert a column of packet numbering
    #         df.insert(3, "No.", range(1,len(df['Email Address'])+1))  
    #         # #create a list of all unique locations.
    #         # unique_points = df[pick_up_point_name].unique()
    #         #Generate a column for location name conversion
    #         df['Location'] = df[pick_up_point_name].map(KK_DICT)
    #         #Generate a column for Flavour name conversion
    #         df['Flavour'] = df['Menu (Select One)'].map(FLAVOUR_DICT)
    #         return df

    #     def output_text_js(self, txt_name_list, df):
    #         unique_locations = df['Location'].unique()
    #         unique_flavours = df['Flavour'].unique()

    #         with open(txt_name_list, 'w') as food_prepare_runner_assign:
    #             food_prepare_runner_assign.write('Chef - Total Food To Prepare: ' + '\n')
    #             food_prepare_runner_assign.write(str(df['Menu (Select One)'].value_counts().rename_axis('Menu').to_frame('counts')))
                
    #             food_prepare_runner_assign.write('Runner - Delivery List: ' + '\n')
    #             for index, flavour in enumerate(unique_flavours):
    #                 if index != 0:
    #                     food_prepare_runner_assign.write(f'\n')
    #                 food_prepare_runner_assign.write(f'*{flavour}*\n')
    #                 for location in unique_locations:
    #                     if (((df["Location"] == location)&(df["Flavour"] == flavour)).sum()) != 0:
    #                         food_prepare_runner_assign.write(f'{location}={((df["Location"] == location)&(df["Flavour"] == flavour)).sum()}\n')

    #             food_prepare_runner_assign.write('Chef - Assign Food List: ' + '\n')
    #             for index, location in enumerate(unique_locations):
    #                 if index != 0:
    #                     food_prepare_runner_assign.write(f'\n')
    #                 for flavour in unique_flavours:
    #                     if (((df["Location"] == location)&(df["Flavour"] == flavour)).sum()) != 0:
    #                         food_prepare_runner_assign.write(f'{location} {flavour}={((df["Location"] == location)&(df["Flavour"] == flavour)).sum()}\n')

    #     def generate_pickup_image_js(self, df, excel_output_path, image_output_path):
    #         '''
    #         Style df > Export it > Generate image from the excel
    #         '''
    #         logging.info(f"Generating pickup image...")
    #         df_style = df.style.apply(highlight_rows_js, axis=1)

    #         #Export as excel
    #         writer = pd.ExcelWriter(excel_output_path)
    #         df_style.to_excel(writer, sheet_name='Sheet1', index = False)

    #         #format for the column width, .set_column(index_start, index_end, width)
    #         # writer.sheets['Sheet1'].set_column(3, 3, 5)
    #         # writer.sheets['Sheet1'].set_column(4, 4, 30)
    #         # writer.sheets['Sheet1'].set_column(5, 5, 40)
    #         # writer.sheets['Sheet1'].set_column(6, 6, 45)

    #         writer.save()
    #         writer.close()

    #         #set No. as index
    #         df = df.set_index("No.")

    #         #Drop columns unwanted in ss
    #         df = df.drop(['Timestamp', 'Email Address','Contact No. (For purpose of food arrival notification)','Receipt of Payment','Location','Flavour'], axis=1)

    #         #highlight just as above
    #         df_style_ss = df.style.apply(highlight_rows_js, axis=1)

    #         #create image
    #         dfi.export(df_style_ss, f'{image_output_path}')

    #     def run_js(self, df_input, date_str):
    #         shop_name = SHOP_NAME['JS']
    #         excel_output_path = f"{date_str} orders.xlsx"
    #         image_output_path = f"{date_str} image.png"
    #         txt_name_list = 'Chef Prep + Runner + Chef Assign.txt'

    #         df = self.process_input_js(df_input)
    #         self.output_text_js(txt_name_list, df) ##print text
    #         self.generate_pickup_image_js(df, excel_output_path, image_output_path)

    #         return excel_output_path, image_output_path, txt_name_list

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
                
                if len(set(F_SERIES).intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'F_SERIES' + '\n')
                
                for fs in F_SERIES:
                    if fs in list(df['Set']):
                        hyk_chef_note.write(f'{fs}'.ljust(8) + f'{(df["Set"] == fs).sum()}' + '\n')
                            
                if len(set(P_SERIES).intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'P_SERIES' + '\n')
                
                for ps in P_SERIES:
                    if ps in list(df['Set']):
                        hyk_chef_note.write(f'{ps}'.ljust(8) + f'{(df["Set"] == ps).sum()}' + '\n')
                            
                if len(set(C_SERIES).intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'C_SERIES' + '\n')
                
                for cs in C_SERIES:
                    if cs in list(df['Set']):
                        hyk_chef_note.write(f'{cs}'.ljust(8) + f'{(df["Set"] == cs).sum()}' + '\n')
                                    
                if len(set(B_SERIES).intersection(set(list(df['Set'])))) != 0 :
                    hyk_chef_note.write('\n' + 'B_SERIES' + '\n')
                
                for bs in B_SERIES:
                    if bs in list(df['Set']):
                        hyk_chef_note.write(f'{bs}'.ljust(8) + f'{(df["Set"] == bs).sum()}' + '\n')

            with open(txt_name_list[1], 'w') as hyk_packaging_list:
                for location in unique_locations_hyk:
                    hyk_packaging_list.write(f'{location} = ')
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
                
        def generate_pickup_image_hyk(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            
            #for excel
            df_copy = df.copy()
            df_copy = df_copy.iloc[:, 3:7]
            df_style = df_copy.style.apply(highlight_rows_hyk, axis=1)
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            #format for the column width, .set_column(index_start, index_end, width)
            writer.sheets['Sheet1'].set_column("A:A", 15)
            writer.sheets['Sheet1'].set_column("B:B", 30)
            writer.sheets['Sheet1'].set_column("C:C", 40)
            writer.sheets['Sheet1'].set_column("D:D", 40)
            # writer.sheets['Sheet1'].set_column(3, 3, 5)
            # writer.sheets['Sheet1'].set_column(4, 4, 30)
            # writer.sheets['Sheet1'].set_column(5, 5, 40)
            # writer.sheets['Sheet1'].set_column(6, 6, 15)
            writer.close()

            o = win32com.client.Dispatch('Excel.Application')
            wb = o.Workbooks.Open(os.path.join(os.getcwd(), excel_output_path))
            ws = wb.Worksheets['Sheet1']

            ws.Range(ws.Cells(1,2),ws.Cells(df_copy.shape[0]+1,df_copy.shape[1])).CopyPicture(Format=2)

            img = ImageGrab.grabclipboard()
            img.save(image_output_path)
            wb.Close(True)

            logging.info(f"Image generated in {image_output_path}!")

        def run_hyk(self, df_input, date_str):
            shop_name = SHOP_NAME['KSHYK2'] #or 'HYK' or 'KS'
            pick_up_point_name = "HYK Food Pick Up Point"
            excel_output_path = f"HYK {date_str} orders.xlsx"
            image_output_path = f"HYK {date_str} image.jpg"
            txt_name_list = ['HYK Total Food To Prepare.txt','HYK Packaging List.txt', 'HYK backup HYK Runner - Location List.txt']

            df = self.process_input_hyk(df_input, pick_up_point_name)
            # display(df)
            self.output_text_hyk(df, pick_up_point_name, txt_name_list) ##print text
            self.generate_pickup_image_hyk(df, excel_output_path, image_output_path)
            return excel_output_path, image_output_path, txt_name_list

        

