# %%
from datetime import datetime, timedelta
from mypackage.helpers import Helper
from mypackage.constants import COLOR_DICT, SORTER, KK_DICT, SHOP_NAME, ALL_SERIES
import io
from apiclient import errors
import re

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

#create a logger
logging.basicConfig(level=logging.INFO)

#setup
load_dotenv()

SERVICE_ACCOUNT = os.getenv('SERVICE_ACCOUNT') #The service acc used to create files.'
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") #the shared google sheet that user can access and input ID
SHARED_PARENT_FOLDER_ID = os.getenv("SHARED_PARENT_FOLDER_ID") #the shared parent folder on personal acc

SCOPES = ["https://www.googleapis.com/auth/forms.body", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/forms.responses.readonly", "https://www.googleapis.com/auth/spreadsheets.readonly"]

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT, scopes=SCOPES)
# credentials = None

# if os.path.exists("token.json"):
#     credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
# if not credentials or not credentials.valid:
#     if credentials and credentials.expired and credentials.refresh_token:
#         credentials.refresh(Request())
#     else:
#         flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
#         credentials = flow.run_local_server(port=0)
#     with open('token.json', 'w') as token:
#         token.write(credentials.to_json())

ks_hyk_helper_obj = Helper.KS_HYK()
kimseng_helper_obj = Helper.KimSeng()
# jackson_helper_obj = Helper.Jackson()
hyk_helper_obj = Helper.HYK()

today_str = (datetime.today() + timedelta(days=1)).strftime('%d-%m-%Y') #today_str = date that file should be run. Hence if today is 04/01/2023, then today str should be 05/01/2023 so that 05/01/2023 row will be read.
today_str = '10-01-2023'

# files = service_drive.files().list( q=f"name='KS&J&HYK' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}' " ,
#                                        spaces='drive').execute()['files']

# %%
service_form = build('forms', 'v1', credentials=credentials) #forms
service_sheet = build('sheets', 'v4', credentials=credentials)
service_drive = build('drive', 'v3', credentials=credentials)
# GOOGLE_SHEET_ID = '1ey4vDFNm3qIogHpLeUVkABsgIbD0UskUAdTnvuentng'

def create_excel_input(formId):
    '''
    Retrieve google sheet and create csv locally
    '''
    # Retrieve google sheet
    form = service_form.forms().get(formId=formId).execute()
    print("form['linkedSheetId']", form['linkedSheetId'])
    sheet_res = service_sheet.spreadsheets().get(spreadsheetId=form['linkedSheetId']).execute()
    dataset = service_sheet.spreadsheets().values().get(
        spreadsheetId= sheet_res['spreadsheetId'],
        range=sheet_res['sheets'][0]['properties']['title'],
        majorDimension= 'ROWS',
    ).execute()
    df = pd.DataFrame(dataset['values'])
    df = df.rename(columns=df.iloc[0]).drop(df.index[0])
    # df = pd.concat([df, df])
    columns_name = list(df.columns)
    column_d, column_e, column_f = columns_name[3], columns_name[4], columns_name[5]
    if df.duplicated(subset=[column_d, column_e, column_f], keep=False).any() == True:
        df_duplicated = df.loc[df.duplicated()]
        file_name_duplicated = f"{sheet_res['properties']['title'].replace('/', '-')}_duplicated.csv"
        df_duplicated.to_csv(file_name_duplicated, index=False)
        df.drop_duplicates(subset=[column_d, column_e, column_f], keep='first', inplace=True)
    else:
        file_name_duplicated="None"
    
    # shop_name = SHOP_NAME['JS']
    # shop_input_path = os.path.join(shop_name, 'input')
    # final_path = os.path.join(shop_input_path, f"{sheet_res['properties']['title'].replace('/', '-')}_Generated.csv")
    file_name_real = f"{sheet_res['properties']['title'].replace('/', '-')}_Generated.csv"
    df.to_csv(file_name_real, index=False)
    if file_name_duplicated != "None":
        return df, [file_name_real, file_name_duplicated]
    return df, [file_name_real]

def create_all_output(df_input, is_kimseng):
    if is_kimseng:
        excel_output_path, image_output_path, txt_name_list = kimseng_helper_obj.run_ks(df_input, today_str)
        return excel_output_path, image_output_path, txt_name_list
    else:
        excel_output_path, image_output_path, txt_name_list = hyk_helper_obj.run_hyk(df_input, today_str)
        return excel_output_path, image_output_path, txt_name_list

# %%
# service_drive.permissions().list(fileId=ks_j_hyk_folder_id).execute()
# service_drive.permissions().delete(permissionId='03183278425046065265', fileId = ks_j_hyk_folder_id).execute()

# res = service_drive.permissions().list(fileId=kimseng_folder_output_today_id).execute()

# permissionId = ""
# if res['permissions'].length > 0:
#     permissionId = res['permissions'][0].get('id')

# if permissionId != "":
    
# print(len(res['permissions'])

# %%
# service_drive.files().delete(fileId=ks_j_hyk_folder_id).execute()
# service_drive.permissions().create(fileId=ks_j_hyk_folder_id, body={
#       'emailAddress': 'byichonggoh@gmail.com',
#       'type': 'user',
#       'role': 'reader',
#   }).execute()

# file_metadata = {
#     'name': f"testing folder",
#     'mimeType': 'application/vnd.google-apps.folder',
# }

# res_file = service_drive.files().create(body=file_metadata, supportsAllDrives=True, fields='id').execute()
# res_file
# new_permission = {
#       'type': 'user',
#       'role': 'reader',
#       'emailAddress': 'byichonggoh@gmail.com',
#   }
# service_drive.permissions().create(fileId='1OgSojD3wHJxjYiLrd8tjKbhB1RoBsipr', body=new_permission, fields='id').execute()

# service_drive.files().list(q=f"mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'").execute()
# service_drive.files().list( q=f"name='KS&J&HYK' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'", spaces='drive',).execute()['files']
# res = service_drive.files().list().execute()['files']
# for r in res:
#     file_id = r['id']
#     print(r)
#     try:
#         service_drive.files().delete(fileId=file_id).execute()
#     except HttpError as e:
#         print(e)


# %%
# service_drive.files().list( q=f"name='KS&J&HYK' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'", spaces='drive',).execute()['files']
# res = service_drive.files().list().execute()['files']
# for r in res:
#     file_id = r['id']
#     print(r)
#     try:
#         service_drive.files().delete(fileId=file_id).execute()
#     except HttpError as e:
#         print(e)


# %%
# service_drive.files().list().execute()['files']

# %%
#Search or Create folders in google drive
def get_or_create_folder_id(file_name, parents=None, is_root=False):
    '''
    file_name: File's Name
    parents: [File Id]
    '''
    file_id = ''
    # if is_root:
    files = service_drive.files().list( q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{parents[0]}' " ,
                                    spaces='drive').execute()['files']
    # else:
    #   files = service_drive.files().list( q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false" ,
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

#create stuff
#so the idea here is to export to files, upload the files, and delete the files. can be triggered through the form id
logging.info("Creating Folders if doesn't exist...")
ks_j_hyk_folder_id = get_or_create_folder_id('KS&J&HYK', [SHARED_PARENT_FOLDER_ID], is_root=True)
input_folder_id = get_or_create_folder_id('Input', [ks_j_hyk_folder_id])
output_folder_id = get_or_create_folder_id('Output', [ks_j_hyk_folder_id])
output_folder_today_id = get_or_create_folder_id(f'{today_str} output', [output_folder_id])

# kimseng_folder_id = get_or_create_folder_id('KimSeng', [ks_j_hyk_folder_id])
# kimseng_folder_input_id = get_or_create_folder_id('KimSengInput', [kimseng_folder_id])
# kimseng_folder_output_id = get_or_create_folder_id('KimSengOutput', [kimseng_folder_id])
# kimseng_folder_output_today_id = get_or_create_folder_id(f'{today_str} Koutput', [kimseng_folder_output_id])

# jackson_folder_id = get_or_create_folder_id('JDinner', [ks_j_hyk_folder_id])
# jackson_folder_input_id = get_or_create_folder_id('JDinnerInput', [jackson_folder_id])
# jackson_folder_output_id = get_or_create_folder_id('JDinnerOutput', [jackson_folder_id])
# jackson_folder_output_today_id = get_or_create_folder_id(f'{today_str} Joutput', [jackson_folder_output_id])

# hyk_folder_id = get_or_create_folder_id('HYK', [ks_j_hyk_folder_id])
# hyk_folder_input_id = get_or_create_folder_id('HYKInput', [hyk_folder_id])
# hyk_folder_output_id = get_or_create_folder_id('HYKOutput', [hyk_folder_id])
# hyk_folder_output_today_id = get_or_create_folder_id(f'{today_str} HYKOutput', [hyk_folder_output_id])

#put csv into input
# f"{sheet_res['properties']['title'].replace('/', '-')}_Generated"

# %%
def file_exist(file_name, mimetype):
    files = service_drive.files().list( q=f"name = '{file_name}' and mimeType = '{mimetype}' and trashed = false" ,
                                       spaces='drive',).execute()['files']
    if len(files) > 0: #exist
        return files[0].get('id')
    return False

def upload_input_file(file_name, parents):
    '''
    Input:
    file_name: File's Name
    parents: [File Id]
    '''
    file_metadata = {'name': file_name,'parents': parents,}
    file_media = MediaFileUpload(file_name, mimetype='text/csv')
    # if file_id := file_exist(file_name, 'text/csv'):
    #     service_drive.files().update(fileId=file_id, body={'name': file_name}, media_body=file_media, fields='id').execute()['id']
    # else:
    file_id = service_drive.files().create(body=file_metadata, media_body=file_media, fields='id').execute()['id']
    file_media = None #To stop reading the file to allow delete

def upload_output_file(parents, excel_name="", image_name="", txt_name_list=[]):
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

def run_program(form_id, folder_input_id, folder_output_today_id, is_kimseng=False):
    df_input, excel_input_name_list = create_excel_input(form_id)
    if len(excel_input_name_list) > 1:
        #this is the duplicate file if exist
        upload_input_file(excel_input_name_list[1], [folder_output_today_id])
        os.remove(excel_input_name_list[1])
    upload_input_file(excel_input_name_list[0], [folder_input_id])
    os.remove(excel_input_name_list[0])
    
    excel_output_path, image_output_path, txt_name_list = create_all_output(df_input, is_kimseng)
    upload_output_file([folder_output_today_id], excel_name=excel_output_path, image_name=image_output_path, txt_name_list=txt_name_list)
    os.remove(excel_output_path)
    os.remove(image_output_path)
    [os.remove(txt_name) for txt_name in txt_name_list]
    return df_input

def clean_text(text):
    '''
    from https://docs.google.com/forms/d/1q9p8iXFZl-aSsxpeci4FhgWE2GIn_-xLuNI_yzhy0lY/edit?usp=drive_web to 1q9p8iXFZl-aSsxpeci4FhgWE2GIn_-xLuNI_yzhy0lY
    '''
    google_form_id = re.search('/d/(.*)/edit', text).group(1)
    print("text: ", text)
    print("google_form_id: ", google_form_id)
    return google_form_id

# Kimseng
ks_df_input = hyk_df_input = pd.DataFrame()
kimseng_form_id = clean_text(google_sheet_df["KimSeng"].item())
if kimseng_form_id != "None":
    ks_df_input = run_program(kimseng_form_id, input_folder_id, output_folder_today_id, is_kimseng=True)

#hyk
hyk_form_id = clean_text(google_sheet_df["HYK"].item())
if hyk_form_id != "None":
    hyk_df_input = run_program(hyk_form_id, input_folder_id, output_folder_today_id, is_kimseng=False)

if not ks_df_input.empty and not hyk_df_input.empty:
    #kshyk combined text
    kshyk_text_name_list=['KSHYK driver.txt', 'Parttimer Salary.txt']
    ks_hyk_helper_obj.output_text_kshyk(ks_df_input, hyk_df_input, kshyk_text_name_list)
    upload_output_file([output_folder_today_id], txt_name_list=kshyk_text_name_list)
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
        upload_output_file([output_folder_today_id], excel_name=OUTPUT_FILE_NAME)
        os.remove(OUTPUT_FILE_NAME)
        logging.info(f"Calculation Done!")
    else:
        logging.info(f"No KimSeng Calculator found! exiting...")

# %%
# service_drive.files().get(fileId="1wqGRN9e-iMY3okNl0KUEnT7gwqjqtiv9WGIvo18fvKc?").execute()
# file_id = service_drive.files().list(q=f"name='KimSeng Calculator.xlsx' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'", spaces='drive',).execute()['files'][0]['id']
# file_id = service_drive.files().list(q=f"name='KimSeng Calculator.xlsx' and trashed=false and parents in '{SHARED_PARENT_FOLDER_ID}'", spaces='drive',).execute()['files'][0]['id']
# request = service_drive.files().get_media(fileId=file_id)
# file = io.BytesIO()
# downloader = MediaIoBaseDownload(file, request)
# done = False
# while done is False:
#     status, done = downloader.next_chunk()
#     print(F'Download {int(status.progress() * 100)}.')

# %%
# # Request body for creating a form
# NEW_FORM = {
#     "info": {
#         "title": "Quickstart form",
#     }
# }

# # Request body to add a multiple-choice question
# NEW_QUESTION = {
#     "requests": [{
#         "createItem": {
#             "item": {
#                 "title": "In what year did the United States land a mission on the moon?",
#                 "questionItem": {
#                     "question": {
#                         "required": True,
#                         "choiceQuestion": {
#                             "type": "RADIO",
#                             "options": [
#                                 {"value": "1965"},
#                                 {"value": "1967"},
#                                 {"value": "1969"},
#                                 {"value": "1971"}
#                             ],
#                             "shuffle": True
#                         }
#                     }
#                 },
#             },
#             "location": {
#                 "index": 0
#             }
#         }
#     }]
# }

# # Creates the initial form
# # result = form_service.forms().create(body=NEW_FORM).execute()

# # Adds the question to the form
# # question_setting = form_service.forms().batchUpdate(formId=result["formId"], body=NEW_QUESTION).execute()

# # Prints the result to show the question has been added
# get_result = form_service.forms().get(formId="1q9p8iXFZl-aSsxpeci4FhgWE2GIn_-xLuNI_yzhy0lY").execute()
# print(get_result)


