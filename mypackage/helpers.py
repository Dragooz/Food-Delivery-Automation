import os
import logging
from mypackage.constants import COLOR_DICT, SORTER, KK_DICT, SHOP_NAME, DRIVER_A, DRIVER_B, SET_DICT, SET_SERIES, F_SERIES, P_SERIES, C_SERIES, B_SERIES 
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

        def output_text_ks(self, txt_name, df, pick_up_point_name,):
            # order list
            dishes = df[df.columns[6]].tolist()
            customizes = df[df.columns[7]].fillna(' ').tolist()
            unique_points = df[pick_up_point_name].unique()
            with open(txt_name, 'w', encoding="utf-8") as order_list_location:
                order_list_location.write("Order List: \n")
                for index, (d, c) in enumerate(zip(dishes, customizes)):
                    order_list_location.write(f'{index+1}. \n')
                    for dish in d.split(','):
                        dish = dish.strip()
                        order_list_location.write(dish)
                        order_list_location.write('\n')

                    for customize in c.split(','):
                        customize = customize.strip()
                        order_list_location.write(customize + '\n')
                    order_list_location.write('\n')
                
                order_list_location.write('-'*50 + '\n')
                order_list_location.write("Order Location: \n")
                #for loop min() and max() of index of each location into txt file.    
                for point in unique_points:
                    if df[df[pick_up_point_name]== point]['Packet No.'].min() == df[df[pick_up_point_name]== point]['Packet No.'].max():
                        order_list_location.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()}={KK_DICT[point]}"+"\n")
                    else:
                        order_list_location.write(f"{df[df[pick_up_point_name]== point]['Packet No.'].min()}-{df[df[pick_up_point_name]== point]['Packet No.'].max()}={KK_DICT[point]}"+"\n")
                
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
            txt_name = 'Order List + Order Location.txt'
            df = self.process_input_ks(df_input, pick_up_point_name)
            # display(df_input)
            self.output_text_ks(txt_name, df, pick_up_point_name) ##print text
            self.generate_pickup_image_ks(df, excel_output_path, image_output_path)

            return excel_output_path, image_output_path, txt_name

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

    #     def output_text_js(self, txt_name, df):
    #         unique_locations = df['Location'].unique()
    #         unique_flavours = df['Flavour'].unique()

    #         with open(txt_name, 'w') as food_prepare_runner_assign:
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
    #         txt_name = 'Chef Prep + Runner + Chef Assign.txt'

    #         df = self.process_input_js(df_input)
    #         self.output_text_js(txt_name, df) ##print text
    #         self.generate_pickup_image_js(df, excel_output_path, image_output_path)

    #         return excel_output_path, image_output_path, txt_name

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
    

        def output_text_hyk(self, df, pick_up_point_name, txt_name):
            #create a list of all unique locations.
            unique_locations_hyk = df['Location'].unique()

            with open(txt_name, 'w') as hyk_chef_note:
                
                hyk_chef_note.write('HYK Chef - Total Food To Prepare.txt' + '\n')
                for set in SET_SERIES:
                    if set in list(df['Set']):
                            hyk_chef_note.write(f'{set}'.ljust(8) + f'{(df["Set"] == set).sum()}' + '\n')
                
                hyk_chef_note.write('\n')
                hyk_chef_note.write('F_SERIES' + '\n')
                
                for set in F_SERIES:
                    if set in list(df['Set']):
                            hyk_chef_note.write(f'{set}'.ljust(8) + f'{(df["Set"] == set).sum()}' + '\n')
                            
                hyk_chef_note.write('\n')
                hyk_chef_note.write('P_SERIES' + '\n')
                
                for set in P_SERIES:
                    if set in list(df['Set']):
                            hyk_chef_note.write(f'{set}'.ljust(8) + f'{(df["Set"] == set).sum()}' + '\n')
                            
                hyk_chef_note.write('\n')
                hyk_chef_note.write('C_SERIES' + '\n')
                
                for set in C_SERIES:
                    if set in list(df['Set']):
                            hyk_chef_note.write(f'{set}'.ljust(8) + f'{(df["Set"] == set).sum()}' + '\n')
                                    
                hyk_chef_note.write('\n')
                hyk_chef_note.write('B_SERIES' + '\n')
                
                for set in B_SERIES:
                    if set in list(df['Set']):
                            hyk_chef_note.write(f'{set}'.ljust(8) + f'{(df["Set"] == set).sum()}' + '\n')

                index = 0
                hyk_chef_note.write('HYK Packaging List.txt' + '\n')
                for location in unique_locations_hyk:
                    index += 1
                    hyk_chef_note.write(f'{index}.'.ljust(4))
                    for set in SET_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in F_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in P_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in C_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in B_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                    
                    hyk_chef_note.write('\n\n')

                hyk_chef_note.write('backup HYK Runner - Location List.txt' + '\n')
                for location in unique_locations_hyk:
                    hyk_chef_note.write(f'{location} = ')
                    for set in SET_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in F_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in P_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in C_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    for set in B_SERIES:
                        if (((df["Location"] == location)&(df['Set'] == set )).sum()) != 0:
                            hyk_chef_note.write(f'{set}{((df["Location"] == location)&(df["Set"] == set)).sum()}'+"  ")
                            
                    hyk_chef_note.write('\n\n')

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
                
        def generate_pickup_image_hyk(self, df, excel_output_path, image_output_path):
            '''
            Style df > Export it > Generate image from the excel
            '''
            logging.info(f"Generating pickup image...")
            
            #for excel
            df_style = df.style.apply(highlight_rows_hyk, axis=1)
            writer = pd.ExcelWriter(excel_output_path)
            df_style.to_excel(writer, sheet_name='Sheet1', index = False)

            #format for the column width, .set_column(index_start, index_end, width)
            # writer.sheets['Sheet1'].set_column(3, 3, 5)
            # writer.sheets['Sheet1'].set_column(4, 4, 30)
            # writer.sheets['Sheet1'].set_column(5, 5, 40)
            # writer.sheets['Sheet1'].set_column(6, 6, 15)
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
            txt_name = 'Order List + Order Location.txt'

            df = self.process_input_hyk(df_input, pick_up_point_name)
            # display(df)
            self.output_text_hyk(df, pick_up_point_name, txt_name) ##print text
            self.generate_pickup_image_hyk(df, excel_output_path, image_output_path)
            return excel_output_path, image_output_path, txt_name

        

