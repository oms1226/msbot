# -*- coding: utf-8 -*-
import re
import calendar

import datetime, time
from datetime import timedelta
import urllib.request
import requests, json
from http.cookiejar import CookieJar
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
from pandas import DataFrame
import pandas.io.sql as pdsql
from matplotlib import dates

import sqlite3


DATABASE = '..\\DATA\\mymoneybot.sqlite'

def sqliteconn():
    conn = sqlite3.connect(DATABASE)
    return conn


def get_webpage(url, encoding=""):
    cj = CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95  Safari/537.36')]

    respstr = ""
    try:
        op = opener.open(url)
        sourcecode = op.read()
    except Exception as e:
        time.sleep(1)
        op = opener.open(url)
        sourcecode = op.read()

    encodingmethod = op.info().get_param('charset')
    if encodingmethod == None:
        if encoding != "":
            encodingmethod = encoding

    if encoding != "":
        encodingmethod = encoding

    try:
        respstr = sourcecode.decode(encoding=encodingmethod, errors='ignore')
    except Exception as e:
        respstr = sourcecode.decode(encoding="cp949", errors='ignore')

    opener.close()

    return respstr


def get_company_fundamental_fnguide(code):

    def g(x):
        if type(x) == str:
            return datetime.datetime.strptime(x, '%Y-%m-%d')
        else:
            return x

    # url = "http://comp.fnguide.com/SVO2/ASP/SVD_main.asp?pGB=1&gicode=A%s&cID=&MenuYn=Y&ReportGB=&NewMenuID=11&stkGb=&strResearchYN=" % (code)
    url = "http://asp01.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A%s&NewMenuID=11&cID=50&MenuYn=N" % (code)
    respstr = get_webpage(url, encoding="utf8")
    # soup = BeautifulSoup(respstr)
    soup = BeautifulSoup(respstr, "lxml")

    # <!--IFRS 별도/연간 -->
    target_table = soup.find("div", class_="um_table", id="highlight_B_Y")
    # print(target_table)
    result = []

    try:
        target_table.find_all('tr')
    except Exception as e:
        return (DataFrame(), DataFrame())

    for tr in target_table.find_all('tr'):
    #     print("[%s]" % tr)
        for th in tr.find_all('th'):
            value = "%s" % th.text.replace('(P) : Provisional','').replace('(E) : Estimate','').replace('잠정실적','').replace('컨센서스, 추정치','').replace('(E)','').replace('(P)','').replace('/','-').strip()
            if ('-02' in value):
                value = value + '-28'
            elif ('-04' in value) or ('-06' in value) or ('-09' in value) or ('-11' in value):
                value = value + '-30'
            elif ('-01' in value) or ('-03' in value) or ('-05' in value) or ('-07' in value) or ('-08' in value) or ('-10' in value) or ('-12' in value):
                value = value + '-31'
            result.append(value)
    #         print("[%s]" % th.text.replace('(E) : Estimate','').replace('컨센서스, 추정치','').strip())
        for td in tr.find_all('td'):
            value = td.text.strip().replace(',','')
            try:
                value = float(value)
            except Exception as e:
                value = 0
            result.append(value)
    #         print(td.text.strip())

    # print(result[1:])
    result = result[1:]
    dfdata = []
    for x in range(0, len(result), 9):
        dfdata.append(result[x:x+9])
    df = DataFrame(data=dfdata, columns = [str(x) for x in range(1,10)]).T
    df.columns = ['날짜', '매출액', '영업이익', '당기순이익', '자산총계', '부채총계', '자본총계', '자본금', '부채비율', '유보율', '영업이익률', '순이익률', 'ROA', 'ROE', 'EPS', 'BPS', 'DPS', 'PER', 'PBR', '발행주식수', '배당수익률']
    df.drop(df.index[[0]], inplace=True)
    # df['날짜'] = df['date'].apply(g)
    # df.drop(['date'], axis=1, inplace=True)
    df = df.convert_objects(convert_numeric=True)
    # df.set_index('날짜', inplace=True)

    df_year = df

    # <!--IFRS 별도/분기 -->
    target_table = soup.find("div", class_="um_table", id="highlight_B_Q")
    # print(target_table)
    result = []
    for tr in target_table.find_all('tr'):
    #     print("[%s]" % tr)
        for th in tr.find_all('th'):
            value = "%s" % th.text.replace('(P) : Provisional','').replace('(E) : Estimate','').replace('잠정실적','').replace('컨센서스, 추정치','').replace('(E)','').replace('(P)','').replace('/','-').strip()
            if ('-02' in value):
                value = value + '-28'
            elif ('-04' in value) or ('-06' in value) or ('-09' in value) or ('-11' in value):
                value = value + '-30'
            elif ('-01' in value) or ('-03' in value) or ('-05' in value) or ('-07' in value) or ('-08' in value) or ('-10' in value) or ('-12' in value):
                value = value + '-31'
            result.append(value)
    #         print("[%s]" % th.text.replace('(E) : Estimate','').replace('컨센서스, 추정치','').strip())
        for td in tr.find_all('td'):
            value = td.text.strip().replace(',','')
            try:
                value = float(value)
            except Exception as e:
                value = 0
            result.append(value)
    #         print(td.text.strip())

    # print(result[1:])
    result = result[1:]
    dfdata = []
    for x in range(0, len(result), 9):
        dfdata.append(result[x:x+9])
    df = DataFrame(data=dfdata, columns = [str(x) for x in range(1,10)]).T
    df.columns = ['날짜', '매출액', '영업이익', '당기순이익', '자산총계', '부채총계', '자본총계', '자본금', '부채비율', '유보율', '영업이익률', '순이익률', 'ROA', 'ROE', 'EPS', 'BPS', 'DPS', 'PER', 'PBR', '발행주식수', '배당수익률']
    df.drop(df.index[[0]], inplace=True)
    # df['날짜'] = df['date'].apply(g)
    # df.drop(['date'], axis=1, inplace=True)
    df = df.convert_objects(convert_numeric=True)
    # df.set_index('날짜', inplace=True)

    df_qtr = df

    return (df_year, df_qtr)


def build_fundamental_data():
    with sqlite3.connect(DATABASE) as conn:
        cursor = conn.cursor()
        replace_sqlite = (
            "replace into 재무정보( 날짜,종목코드,기간구분,매출액,영업이익,당기순이익,자산총계,부채총계,자본총계,자본금,부채비율,유보율,영업이익률,순이익률,ROA,ROE,EPS,BPS,DPS,PER,PBR,발행주식수,배당수익률 ) "
            "values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
        )

        df = pdsql.read_sql_query('select 단축코드, 종목명 from 종목코드 ', con=conn)

        CODES = list(df.values)
        for code, name in CODES:
            print('FnGuide - %s %s' % (code, name))

            try:
                (df_year, df_qtr) = get_company_fundamental_fnguide(code)
            except Exception as e:
                continue

            if len(df_year.index) > 0 or len(df_qtr.index) > 0:
                if len(df_year.index) > 0:
                    기간구분 = '년간'
                    for idx, row in df_year.iterrows():
                        날짜, 매출액, 영업이익, 당기순이익, 자산총계, 부채총계, 자본총계, 자본금, 부채비율, 유보율, 영업이익률, 순이익률, ROA, ROE, EPS, BPS, DPS, PER, PBR, 발행주식수, 배당수익률 = row
                        종목코드 = code
                        d = (날짜,종목코드,기간구분,매출액,영업이익,당기순이익,자산총계,부채총계,자본총계,자본금,부채비율,유보율,영업이익률,순이익률,ROA,ROE,EPS,BPS,DPS,PER,PBR,발행주식수,배당수익률)
                        cursor.execute(replace_sqlite, d)
                        conn.commit()


                if len(df_qtr.index) > 0:
                    기간구분 = '분기'
                    for idx, row in df_qtr.iterrows():
                        날짜, 매출액, 영업이익, 당기순이익, 자산총계, 부채총계, 자본총계, 자본금, 부채비율, 유보율, 영업이익률, 순이익률, ROA, ROE, EPS, BPS, DPS, PER, PBR, 발행주식수, 배당수익률 = row
                        종목코드 = code
                        d = (날짜,종목코드,기간구분,매출액,영업이익,당기순이익,자산총계,부채총계,자본총계,자본금,부채비율,유보율,영업이익률,순이익률,ROA,ROE,EPS,BPS,DPS,PER,PBR,발행주식수,배당수익률)
                        cursor.execute(replace_sqlite, d)
                        conn.commit()

                # time.sleep(2)

            # except Exception as e:
            #     print(code, name, str(e))


if __name__ == "__main__":
    # 재무정보가져오기 - 분기에 한번 실행하면 됨
    build_fundamental_data()
