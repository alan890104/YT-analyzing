# -*- coding: utf-8 -*-

import os
import openpyxl
import http.client
import google.oauth2.credentials
import google_auth_oauthlib.flow
import oauth2client
from oauth2client.file import Storage
from oauth2client import client
from oauth2client import tools
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/yt-analytics.readonly']

API_SERVICE_NAME = 'youtubeAnalytics'
API_VERSION = 'v2'
client_secrets_file = 'CS.json'
def get_service():
    credential_path = os.path.join('./', 'credential.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_secrets_file, SCOPES)
        credentials = tools.run_flow(flow, store)
    return build(API_SERVICE_NAME, API_VERSION, credentials = credentials)

def execute_api_request(client_library_function, **kwargs):
    response = client_library_function(**kwargs).execute()
    res = response['rows']
    print(res)
    return res

def write_excel(res):
    workbook = openpyxl.Workbook()
    sheet=workbook.active
    title = ['日期','觀看數','觀看時長(小時)','回放時間(秒)','回放比率','獲得訂閱者','失去訂閱者']
    sheet.append(title)
    for s in res:
        print(s)
        sheet.append(s)
    for xx in ['A','B','C','D','E','F','G']:
        sheet.column_dimensions[xx].width =15
    workbook.save('data.xlsx')

if __name__ == '__main__':
    # Disable OAuthlib's HTTPs verification when running locally.
    # *DO NOT* leave this option enabled when running in production.
    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

    youtubeAnalytics = get_service()
    res = execute_api_request(
        youtubeAnalytics.reports().query,
        ids='channel==MINE',
        startDate='2020-01-01',
        endDate='2020-08-02',
        metrics='views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage,subscribersGained,subscribersLost',
        dimensions='day',
        sort='day'
    )
    write_excel(res)
    