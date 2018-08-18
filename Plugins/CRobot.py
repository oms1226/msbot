import os, sys, datetime, time
sys.path.append('..')

import pandas as pd
import pandas.io.sql as pdsql
from pandas import DataFrame, Series
# from pandas.lib import Timestamp

import sqlite3

from Utils import *

##
## 포트폴리오에 사용되는 주식정보 클래스
class CPortStock(object):
    def __init__(self, 매수일, 종목코드, 종목명, 매수가, 수량, STATUS=''):
        self.매수일 = 매수일
        self.종목코드 = 종목코드
        self.종목명 = 종목명
        self.매수가 = 매수가
        self.수량 = 수량
        self.STATUS = STATUS

        self.매수후고가 = 매수가


class CRobot(object):
    def instance(self):
        pass

    def __init__(self, Name, UUID):
        self.매도 = 1
        self.매수 = 2
        self.지정가 = '00'
        self.시장가 = '03'
        self.조건없음 = '0'
        self.조건IOC = '1'
        self.조건FOK = '2'

        self.신용거래코드 = '000'

        self.Name = Name
        self.UUID = UUID
        self.DATABASE = None
        self.running = False
        self.portfolio = dict()

    def set_parent(self, parent):
        self.parent = parent

    def set_database(self, database):
        self.DATABASE = database

    def set_secret(self, 계좌번호='계좌번호', 비밀번호='비밀번호'):
        self.계좌번호 = 계좌번호
        self.비밀번호 = 비밀번호

    def modal(self, parent):
        pass

    def getstatus(self):
        result = []
        for p, v in self.portfolio.items():
            s = '%s(%s)[P%s/V%s/D%s]' % (v.종목명.strip(), v.종목코드, v.매수가, v.수량, v.매수일)
            result.append(s)

        return [self.__class__.__name__, self.Name, self.UUID, self.running, len(self.portfolio), ','.join(result)]

    def OnReceiveRealData(self, szTrCode, result):
        pass

    def 초기조건(self):
        pass

    def Run(self, flag=True, parent=None):
        pass

    def 포트폴리오읽기(self):
        pass

    def 포트폴리오쓰기(self):
        pass

    def 포트폴리오종목갱신(self, 포트폴리오키, P):
        pass

    def 체결기록(self, data):
        try:
            with sqlite3.connect(self.DATABASE) as conn:
                query = 'insert into 거래결과(로봇명, UUID, 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
                cursor = conn.cursor()
                cursor.execute(query, data)
                conn.commit()
        except Exception as e:
            클래스이름 = self.__class__.__name__
            함수이름 = inspect.currentframe().f_code.co_name
            print("%s-%s %s" % (클래스이름, 함수이름, get_linenumber()), e)

    def 주문기록(self, data):
        pass

def robot_loader():
    return None