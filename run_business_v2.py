from datetime import datetime
from mypackage.helpers import Helper
from mypackage.constants import COLOR_DICT, SORTER, KK_DICT, SHOP_NAME
import io
from apiclient import errors

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

def create_excel_input(formId):
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

    # shop_name = SHOP_NAME['JS']
    # shop_input_path = os.path.join(shop_name, 'input')
    # final_path = os.path.join(shop_input_path, f"{sheet_res['properties']['title'].replace('/', '-')}_Generated.csv")
    file_name = f"{sheet_res['properties']['title'].replace('/', '-')}_Generated.csv"
    df.to_csv(file_name, index=False)
    return df, file_name

def create_all_output(df_input, is_kimseng):
    if is_kimseng:
        excel_output_path, image_output_path, txt_name = kimseng_helper_obj.run_ks(df_input, today_str)
        return excel_output_path, image_output_path, txt_name
    else:
        excel_output_path, image_output_path, txt_name = hyk_helper_obj.run_hyk(df_input, today_str)
        return excel_output_path, image_output_path, txt_name

#Search or Create folders in google drive
def get_or_create_folder_id(file_name, parents=None, is_root=False):
    '''
    file_name: File's Name
    parents: [File Id]
    '''
    file_id = ''
    if is_root:
      files = service_drive.files().list( q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false and parents in '{parents[0]}' " ,
                                       spaces='drive').execute()['files']
    else:
      files = service_drive.files().list( q=f"name='{file_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false" ,
                                       spaces='drive').execute()['files']
    if len(files) > 0: #exist
        print(f"{file_name} exist. ")
        file_id = files[0].get('id')
    
    #create folder
    file_metadata = {
            'name': f"{file_name}",
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': parents, 
        }

    res_file = service_drive.files().create(body=file_metadata, supportsAllDrives=True, fields='id').execute()
    file_id = res_file.get('id')
    return file_id

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
    if file_id := file_exist(file_name, 'text/csv'):
        service_drive.files().update(fileId=file_id, body={'name': file_name}, media_body=file_media, fields='id').execute()['id']
    else:
        file_id = service_drive.files().create(body=file_metadata, media_body=file_media, fields='id').execute()['id']
    file_media = None #To stop reading the file to allow delete

def upload_output_file(excel_name, image_name, text_name, parents):
    '''
    Output
    file_name: File's Name
    parents: [File Id]
    '''


    image_metadata = {'name': image_name, 'parents': parents,}
    image_media = MediaFileUpload(image_name, mimetype='image/png')
    if image_file_id := file_exist(image_name, 'image/png'):
        service_drive.files().update(fileId= image_file_id, body={'name': image_name}, media_body=image_media, fields='id').execute()['id']
    else:
        image_file_id = service_drive.files().create(body=image_metadata, media_body=image_media, fields='id').execute()['id']
    image_media=None

    text_metadata = {'name': text_name,'parents': parents,}
    text_media = MediaFileUpload(text_name, mimetype='text/plain')
    if text_file_id := file_exist(text_name, 'text/plain'):
        service_drive.files().update(fileId=text_file_id, body={'name': text_name}, media_body=text_media, fields='id').execute()['id']
    else:
        text_file_id = service_drive.files().create(body=text_metadata, media_body=text_media, fields='id').execute()['id']
    text_media=None

    excel_metadata = {'name': excel_name,'parents': parents}
    excel_media = MediaFileUpload(excel_name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if excel_file_id := file_exist(excel_name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
        service_drive.files().update(fileId=excel_file_id, body={'name': excel_name}, media_body=excel_media, fields='id').execute()['id']
    else:
        excel_file_id = service_drive.files().create(body=excel_metadata, media_body=excel_media, fields='id').execute()['id']
    excel_media=None

    
#setup
load_dotenv()

SERVICE_ACCOUNT = os.getenv('SERVICE_ACCOUNT') #The service acc used to create files.
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID") #the shared google sheet that user can access and input ID
SHARED_PARENT_FOLDER_ID = os.getenv("SHARED_PARENT_FOLDER_ID") #the shared parent folder on personal acc

SCOPES = ["https://www.googleapis.com/auth/forms.body", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/forms.responses.readonly", "https://www.googleapis.com/auth/spreadsheets.readonly"]

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT, scopes=SCOPES)
kimseng_helper_obj = Helper.KimSeng()
# jackson_helper_obj = Helper.Jackson()
hyk_helper_obj = Helper.HYK()

today_str = datetime.today().strftime('%d-%m-%Y')

# %%
service_form = build('forms', 'v1', credentials=credentials) #forms
service_sheet = build('sheets', 'v4', credentials=credentials)
service_drive = build('drive', 'v3', credentials=credentials)
#create stuff
#so the idea here is to export to files, upload the files, and delete the files. can be triggered through the form id

ks_j_hyk_folder_id = get_or_create_folder_id('KS&J&HYK', [SHARED_PARENT_FOLDER_ID], is_root=True)
kimseng_folder_id = get_or_create_folder_id('KimSeng', [ks_j_hyk_folder_id])
kimseng_folder_input_id = get_or_create_folder_id('KimSengInput', [kimseng_folder_id])
kimseng_folder_output_id = get_or_create_folder_id('KimSengOutput', [kimseng_folder_id])
kimseng_folder_output_today_id = get_or_create_folder_id(f'{today_str} Koutput', [kimseng_folder_output_id])

hyk_folder_id = get_or_create_folder_id('HYK', [ks_j_hyk_folder_id])
hyk_folder_input_id = get_or_create_folder_id('HYKInput', [hyk_folder_id])
hyk_folder_output_id = get_or_create_folder_id('HYKOutput', [hyk_folder_id])
hyk_folder_output_today_id = get_or_create_folder_id(f'{today_str} HYKOutput', [hyk_folder_output_id])

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


#Kimseng
kimseng_sheet_id = google_sheet_df["KimSeng"].item()
if kimseng_sheet_id != "None":
    df_input, excel_input_name = create_excel_input(kimseng_sheet_id)
    excel_output_path, image_output_path, txt_name = create_all_output(df_input, True)
    upload_input_file(excel_input_name, [kimseng_folder_input_id])
    os.remove(excel_input_name)

    upload_output_file(excel_output_path, image_output_path, txt_name, [kimseng_folder_output_today_id])
    os.remove(excel_output_path)
    os.remove(image_output_path)
    os.remove(txt_name)
else:
    upload_input_file('test.txt', ['1KmT34PACuq4avrlrVDnR6_rAoi3c1hGV'])

#hyk
hyk_sheet_id = google_sheet_df["HYK"].item()
if hyk_sheet_id != "None":
    df_input, excel_input_name = create_excel_input(hyk_sheet_id)
    excel_output_path, image_output_path, txt_name = create_all_output(df_input, True)
    upload_input_file(excel_input_name, [hyk_folder_input_id])
    os.remove(excel_input_name)

    upload_output_file(excel_output_path, image_output_path, txt_name, [hyk_folder_output_today_id])
    os.remove(excel_output_path)
    os.remove(image_output_path)
    os.remove(txt_name)

