# -*- coding: utf-8 -*-

import os
import datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Series, Reference

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
    return res

def write_excel(res,today):
    workbook = Workbook()
    sheet=workbook.active
    title = ['日期','觀看數','觀看時長(小時)','回放時間(秒)','回放比率','獲得訂閱者','失去訂閱者']

    #將資料放入sheet中
    sheet.append(title)
    for s in res:  sheet.append(s)
    #調整每個欄位的寬度
    for col in ['A','B','C','D','E','F','G']:
        sheet.column_dimensions[col].width =15
    #畫柱狀圖
    chart = BarChart()
    chart.type = "col"
    chart.style = 2
    chart.title = "youtube數據分析"

    start = datetime.datetime.strptime('2020-03-01',"%Y-%m-%d")
    days_count = (datetime.datetime.today()-datetime.timedelta(days=3)-start).days
    data = Reference(sheet,min_row=1, max_row=int(days_count)+2, min_col=2,max_col=7)
    dates = Reference(sheet, min_row=2,max_row=int(days_count)+2, min_col=1,max_col=1) 
    chart.add_data(data, titles_from_data=True)
    chart.x_axis.number_format = "yyyy-mm-dd"
    chart.set_categories(dates)
    chart.shape = 10
    chart.width = 20
    chart.height = 15
    sheet.add_chart(chart, "I3")
    
    #存成當天日期.xlsx
    workbook.save(today+'.xlsx')

if __name__ == '__main__':
    # Disable OAuthlib's HTTPs verification when running locally.
    # *DO NOT* leave this option enabled when running in production.
    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

    today = str(datetime.datetime.today().strftime("%Y-%m-%d"))

    youtubeAnalytics = get_service()
    res = execute_api_request(
        youtubeAnalytics.reports().query,
        ids='channel==MINE',
        startDate='2020-03-01',
        endDate=today,
        metrics='views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage,subscribersGained,subscribersLost',
        dimensions='day',
        sort='day'
    )
    write_excel(res,today)
    