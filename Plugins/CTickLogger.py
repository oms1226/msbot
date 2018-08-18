import os, sys
import uuid

import pandas as pd
import pandas.io.sql as pdsql
from pandas import DataFrame, Series
# from pandas.lib import Timestamp

import sqlite3

import PyQt5
from PyQt5 import QtCore, QtGui, uic
from PyQt5 import QAxContainer
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import (QApplication, QLabel, QLineEdit, QMainWindow, QDialog, QMessageBox, QProgressBar)
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *

from XASessions import *
from XAQuaries import *
from XAReals import *

from CRobot import *
from Utils import *

__PATHNAME__ = os.path.dirname(sys.argv[0])
__PLUGINDIR__ = os.path.abspath(__PATHNAME__)

ROBOT_NAME = "TickLogger"

Ui_CTickLogger, QtBaseClass_CTickLogger = uic.loadUiType("%s\\Plugins\\CTickLogger.ui" % __PLUGINDIR__)
class CUITickLogger(QDialog, Ui_CTickLogger):
    def __init__(self, parent=None):
        super(__class__, self).__init__(parent)
        # self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent


class CTickLogger(CRobot):
    def instance(self):
        UUID = uuid.uuid4().hex
        return CTickLogger(Name=ROBOT_NAME, UUID=UUID)

    def __init__(self, Name, UUID):
        super(__class__, self).__init__(Name, UUID)

        self.QueryInit()

    def QueryInit(self):
        self.kospi_real = None
        self.kosdaq_real = None

    def modal(self, parent):
        ui = CUITickLogger(parent=parent)
        ui.setModal(True)
        ui.lineEdit_name.setText(self.Name)

        r = ui.exec_()
        if r == 1:
            이름 = ui.lineEdit_name.text()
            종목유니버스 = ui.plainTextEdit_base_price.toPlainText()
            self.종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]
            self.Name = 이름

            print(이름, 종목유니버스, self.종목유니버스리스트)

        return r

    def OnReceiveRealData(self, szTrCode, result):
        if szTrCode in ['K3_','S3_']:
            lst = [
                result['체결시간'],
                result['전일대비구분'],
                result['전일대비'],
                result['등락율'],
                result['현재가'],
                result['시가시간'],
                result['시가'],
                result['고가시간'],
                result['고가'],
                result['저가시간'],
                result['저가'],
                result['체결구분'],
                result['체결량'],
                result['누적거래량'],
                result['누적거래대금'],
                result['매도누적체결량'],
                result['매도누적체결건수'],
                result['매수누적체결량'],
                result['매수누적체결건수'],
                result['체결강도'],
                result['가중평균가'],
                result['매도호가'],
                result['매수호가'],
                result['장정보'],
                result['전일동시간대거래량'],
                result['단축코드']
            ]

        str = '{},{},{},{}\r'.format(result['체결시간'], result['단축코드'], result['현재가'], result['체결량'])
        self.handle.write(str)
        self.handle.flush()
        print(str)

    def Run(self, flag=True, parent=None):
        self.parent = parent
        self.running = flag
        ret = 0
        if flag == True:
            ToTelegram("로직 [%s]을 시작합니다." % (__class__.__name__))

            if self.DATABASE != None:

                self.handle = open('DATA/TickLogger.csv', 'a+')

                self.kospi_real = S3_(parent=self)
                self.kosdaq_real = K3_(parent=self)

                with sqlite3.connect(self.DATABASE) as conn:
                    query = 'select 단축코드,종목명,ETF구분,구분 from 종목코드'
                    df = pdsql.read_sql_query(query, con=conn)

                # print(df)

                kospi_codes = df.query("구분=='1'")['단축코드'].values.tolist()
                kosdaq_codes = df.query("구분=='2'")['단축코드'].values.tolist()

                for code in self.종목유니버스리스트:
                    if code in kospi_codes:
                        self.kospi_real.AdviseRealData(종목코드=code)
                    if code in kosdaq_codes:
                        self.kosdaq_real.AdviseRealData(종목코드=code)

        else:
            if self.kospi_real != None:
                self.kospi_real.UnadviseRealData()
                self.kospi_real = None
            if self.kosdaq_real != None:
                self.kosdaq_real.UnadviseRealData()
                self.kosdaq_real = None

            try:
                self.handle.close()
                self.handle = None
            except Exception as e:
                self.handle = None
                pass

def robot_loader():
    UUID = uuid.uuid4().hex
    robot = CTickLogger(Name=ROBOT_NAME, UUID=UUID)
    return robot