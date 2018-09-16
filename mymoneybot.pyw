# -*- coding: utf-8 -*-

프로그램정보 = [
    ['프로그램명','mymoneybot-eBEST'],
    ['Version','1.4'],
    ['개발일','2018-02-28'],
    ['2018-06-04','포트폴리오 더블클릭으로 삭제 기능 추가'],
    ['2018-05-23','시장가매도, query->ActiveX 오류수정'],
    ['2018-07-19','국내선물옵션, 해외선물옵션에 필요한 모듈을 XAQuery, XAReals에 추가'],
    ['2018-07-19','검색식에서 종목이 빠지는 경우, 손절 및 익절이 나가지 않는 부분 추가'],
    ['2018-07-20','체결시간과 종목검색에서 종목이 빠지는 시간차가 있는 경우 주문이 나가지 않는 부분추가'],
    ['2018-07-25','종목검색 중지시 계속 검색된 종목이 들어오는 문제 수정'],
    ['2018-08-01','종목검색, Chartindex에서 식별자를 사용하는 방법 통일'],
    ['2018-08-01','한번에 수량이 다 체결된 경우 포트에 반영되지 않는 것을 수정'],
    ['2018-08-07','조건검색시 다른 조건검색과 섞이는 것을 수정'],
    ['2018-08-07','API메뉴중 백업에 OnReceiveMessage 추가']
]


import sys, os
import datetime, time
import win32com.client
import pythoncom
import inspect

import pickle
import uuid
import base64
import subprocess
from subprocess import Popen
import webbrowser

import PyQt5
from PyQt5 import QtCore, QtGui, uic
from PyQt5 import QAxContainer
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import (QApplication, QLabel, QLineEdit, QMainWindow, QDialog, QMessageBox, QProgressBar)
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *

import numpy as np
from numpy import NaN, Inf, arange, isscalar, asarray, array

import pandas as pd
import pandas.io.sql as pdsql
from pandas import DataFrame, Series

import sqlite3

import logging
import logging.handlers

from XASessions import *
from XAQuaries import *
from XAReals import *

from FileWatcher import *
from Utils import *


주문지연 = 3000

DATABASE = 'DATA\\mymoneybot.sqlite'
UI_DIR = "UI\\"


def sqliteconn():
    conn = sqlite3.connect(DATABASE)
    return conn

class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data=None, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data
        if data is None:
            self._data = DataFrame()

    def rowCount(self, parent=None):
        return len(self._data.index)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.values[index.row()][index.column()])
        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self._data.columns[column]
        return int(column + 1)

    def update(self, data):
        self._data = data
        self.reset()

    def reset(self):
        self.beginResetModel()
        self.endResetModel()

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled

class RealDataTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.realdata = {}
        self.headers = ['종목코드', '현재가' , '전일대비', '등락률' , '매도호가', '매수호가', '누적거래량', '시가' , '고가' , '저가' , '거래회전율', '시가총액']

    def rowCount(self, index=QModelIndex()):
        return len(self.realdata)

    def columnCount(self, index=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if (not index.isValid() or not (0 <= index.row() < len(self.realdata))):
            return None

        if role == Qt.DisplayRole:
            rows = []
            for k in self.realdata.keys():
                rows.append(k)
            one_row = rows[index.row()]
            selected_row = self.realdata[one_row]

            return selected_row[index.column()]

        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.headers[column]
        return int(column + 1)

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled

    def reset(self):
        self.beginResetModel()
        self.endResetModel()

class CPluginManager:
    plugins = None
    @classmethod
    def plugin_loader(cls):
        path = "plugins/"
        result = {}

        # Load plugins
        sys.path.insert(0, path)
        for f in os.listdir(path):
            fname, ext = os.path.splitext(f)
            if ext == '.py':
                mod = __import__(fname)
                robot = mod.robot_loader()
                if robot != None:
                    result[robot.Name] = robot
        sys.path.pop(0)

        CPluginManager.plugins = result

        return result


Ui_계좌정보조회, QtBaseClass_계좌정보조회 = uic.loadUiType(UI_DIR+"계좌정보조회.ui")
class 화면_계좌정보(QDialog, Ui_계좌정보조회):
    def __init__(self, parent=None):
        super(화면_계좌정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent
        self.model1 = PandasModel()
        self.tableView_1.setModel(self.model1)
        self.model2 = PandasModel()
        self.tableView_2.setModel(self.model2)

        self.result = []
        self.connection = self.parent.connection

        # 계좌정보 불러오기
        nCount = self.connection.ActiveX.GetAccountListCount()
        for i in range(nCount):
            self.comboBox.addItem(self.connection.ActiveX.GetAccountList(i))

        self.XQ_t0424 = t0424(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't0424':
            self.df1, self.df2 = result

            self.model1.update(self.df1)
            for i in range(len(self.df1.columns)):
                self.tableView_1.resizeColumnToContents(i)

            self.model2.update(self.df2)
            for i in range(len(self.df2.columns)):
                self.tableView_2.resizeColumnToContents(i)

            CTS_종목번호 = self.df1['CTS_종목번호'].values[0].strip()
            if CTS_종목번호 != '':
                self.XQ_t0424.Query(계좌번호=self.계좌번호, 비밀번호=self.비밀번호, 단가구분='1', 체결구분='0', 단일가구분='0', 제비용포함여부='1', CTS_종목번호=CTS_종목번호)

    def inquiry(self):
        self.계좌번호 = self.comboBox.currentText().strip()
        self.비밀번호 = self.lineEdit.text().strip()

        self.XQ_t0424.Query(계좌번호=self.계좌번호,비밀번호=self.비밀번호,단가구분='1',체결구분='0',단일가구분='0',제비용포함여부='1',CTS_종목번호='')

        QTimer().singleShot(3*1000, self.inquiry)


Ui_일별가격정보백업, QtBaseClass_일별가격정보백업 = uic.loadUiType(UI_DIR+"일별가격정보백업.ui")
class 화면_일별가격정보백업(QDialog, Ui_일별가격정보백업):
    def __init__(self, parent=None):
        super(화면_일별가격정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('가격 정보 백업')
        self.parent = parent
        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        XQ_t8436 = t8436(parent=self)
        XQ_t8436.Query(구분='0')

        self.조회건수 = 10
        self.XQ_t1305 = t1305(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1305':
            CNT, 날짜, IDX, df = result
            # print(self.단축코드, CNT, 날짜, IDX)
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 일별주가( 날짜, 시가, 고가, 저가, 종가, 전일대비구분, 전일대비, 등락율, 누적거래량, 거래증가율, 체결강도, 소진율, 회전율, 외인순매수, 기관순매수, 종목코드, 누적거래대금, 개인순매수, 시가대비구분, 시가대비, 시가기준등락율, 고가대비구분, 고가대비, 고가기준등락율, 저가대비구분, 저가대비, 저가기준등락율, 시가총액) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                cursor.executemany(query, df.values.tolist())
                conn.commit()

            try:
                if int(CNT) == int(self.조회건수) and self.radioButton_all.isChecked() == True:
                    QTimer.singleShot(주문지연, lambda: self.Request(result=result))
                else:
                    self.백업한종목수 += 1
                    if len(self.백업할종목코드) > 0:
                        self.단축코드 = self.백업할종목코드.pop(0)
                        self.result = []

                        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
                        self.label_codename.setText(S)

                        QTimer.singleShot(주문지연, lambda : self.Request([]))
                    else:
                        QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")
            except Exception as e:
                pass

    def Request(self, result=[]):
        if len(result) > 0:
            CNT, 날짜, IDX, df = result
            self.XQ_t1305.Query(단축코드=self.단축코드[0], 일주월구분='1', 날짜=날짜, IDX=IDX, 건수=self.조회건수, 연속조회=True)
        else:
            try:
                # print('%s %s' % (self.단축코드[0], self.단축코드[1]))
                self.XQ_t1305.Query(단축코드=self.단축코드[0], 일주월구분='1', 날짜='', IDX='', 건수=self.조회건수, 연속조회=False)
            except Exception as e:
                pass

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.단축코드 = self.종목코드테이블[idx:idx + 1][['단축코드','종목명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(result=[])

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['단축코드','종목명']].values)
        self.단축코드 = self.백업할종목코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(result=[])


Ui_일별업종정보백업, QtBaseClass_일별업종정보백업 = uic.loadUiType(UI_DIR+"일별업종정보백업.ui")
class 화면_일별업종정보백업(QDialog, Ui_일별업종정보백업):
    def __init__(self, parent=None):
        super(화면_일별업종정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('업종 정보 백업')
        self.parent = parent

        self.columns = ['현재가', '거래량', '일자', '시가', '고가', '저가','거래대금', '대업종구분', '소업종구분', '종목정보', '종목정보', '수정주가이벤트', '전일종가']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        XQ = t8424(parent=self)
        XQ.Query()

        self.조회건수 = 10
        self.XQ_t1514 = t1514(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8424':
            df = result[0]
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 업종코드(업종명, 업종코드) values(?, ?)"
                cursor.executemany(query, df.values.tolist())
                conn.commit()

            self.업종코드테이블 = result[0]
            self.업종코드테이블['컬럼'] = ">> " + self.업종코드테이블['업종코드'] + " : " + self.업종코드테이블['업종명']
            self.업종코드테이블 = self.업종코드테이블.sort_values(['업종코드', '업종명'], ascending=[True, True])
            self.comboBox.addItems(self.업종코드테이블['컬럼'].values)

        if szTrCode == 't1514':
            CTS일자, df = result
            # print(CTS일자)
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 업종정보(일자, 지수, 전일대비구분, 전일대비, 등락율, 거래량, 거래증가율, 거래대금1, 상승, 보합, 하락, 상승종목비율, 외인순매수, 시가, 고가, 저가, 거래대금2, 상한, 하한, 종목수, 기관순매수, 업종코드, 거래비중, 업종배당수익률) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                cursor.executemany(query, df.values.tolist())
                conn.commit()

            try:
                if len(df) == int(self.조회건수) and self.radioButton_all.isChecked() == True:
                    QTimer.singleShot(주문지연, lambda: self.Request(result=result))
                else:
                    self.백업한종목수 += 1
                    if len(self.백업할업종코드) > 0:
                        self.업종코드 = self.백업할업종코드.pop(0)
                        self.result = []

                        self.progressBar.setValue(int(self.백업한종목수 / (len(self.업종코드테이블.index) - self.comboBox.currentIndex()) * 100))
                        S = '%s %s' % (self.업종코드[0], self.업종코드[1])
                        self.label_codename.setText(S)

                        QTimer.singleShot(주문지연, lambda : self.Request([]))
                    else:
                        QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")
            except Exception as e:
                pass

    def Request(self, result=[]):
        if len(result) > 0:
            CTS일자, df = result
            self.XQ_t1514.Query(업종코드=self.업종코드[0],구분1='',구분2='1',CTS일자=CTS일자, 조회건수=self.조회건수,비중구분='', 연속조회=True)
        else:
            # print('%s %s' % (self.업종코드[0], self.업종코드[1]))
            self.XQ_t1514.Query(업종코드=self.업종코드[0], 구분1='', 구분2='1', CTS일자='', 조회건수=self.조회건수, 비중구분='', 연속조회=False)

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할업종코드 = []
        self.업종코드 = self.업종코드테이블[idx:idx + 1][['업종코드','업종명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(result=[])

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할업종코드 = list(self.업종코드테이블[idx:][['업종코드','업종명']].values)
        self.업종코드 = self.백업할업종코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.업종코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.업종코드[0], self.업종코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(result=[])


Ui_분별가격정보백업, QtBaseClass_분별가격정보백업 = uic.loadUiType(UI_DIR+"분별가격정보백업.ui")
class 화면_분별가격정보백업(QDialog, Ui_분별가격정보백업):
    def __init__(self, parent=None):
        super(화면_분별가격정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('가격 정보 백업')
        self.parent = parent

        self.columns = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

        self.result = []

        XQ_t8436 = t8436(parent=self)
        XQ_t8436.Query(구분='0')

        self.조회건수 = 10
        self.XQ_t1302 = t1302(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1302':
            시간CTS, df = result
            df['단축코드'] = self.단축코드[0]
            # print(시간CTS)
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 분별주가(시간, 종가, 전일대비구분, 전일대비, 등락율, 체결강도, 매도체결수량, 매수체결수량, 순매수체결량, 매도체결건수, 매수체결건수, 순체결건수, 거래량, 시가, 고가, 저가, 체결량, 매도체결건수시간, 매수체결건수시간, 매도잔량, 매수잔량, 시간별매도체결량, 시간별매수체결량,단축코드) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                cursor.executemany(query, df.values.tolist())
                conn.commit()

            try:
                if len(df) == int(self.조회건수) and self.radioButton_all.isChecked() == True:
                    QTimer.singleShot(주문지연, lambda: self.Request(result=result))
                else:
                    self.백업한종목수 += 1
                    if len(self.백업할종목코드) > 0:
                        self.단축코드 = self.백업할종목코드.pop(0)
                        self.result = []

                        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
                        self.label_codename.setText(S)

                        QTimer.singleShot(주문지연, lambda : self.Request([]))
                    else:
                        QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")
            except Exception as e:
                pass

    def Request(self, result=[]):
        if len(result) > 0:
            시간CTS, df = result
            self.XQ_t1302.Query(단축코드=self.단축코드[0], 작업구분=self.틱범위, 시간=시간CTS, 건수=self.조회건수, 연속조회=True)
        else:
            # print('%s %s' % (self.단축코드[0], self.단축코드[1]))
            self.XQ_t1302.Query(단축코드=self.단축코드[0], 작업구분=self.틱범위, 시간='', 건수=self.조회건수, 연속조회=False)

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.단축코드 = self.종목코드테이블[idx:idx + 1][['단축코드','종목명']].values[0]
        self.틱범위 = self.comboBox_min.currentText()[0:1].strip()
        if self.틱범위[0] == '0':
            self.틱범위 = self.틱범위[1:]
        self.result = []
        self.Request(result=[])

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['단축코드','종목명']].values)
        self.단축코드 = self.백업할종목코드.pop(0)
        self.틱범위 = self.comboBox_min.currentText()[0:1].strip()
        if self.틱범위[0] == '0':
            self.틱범위 = self.틱범위[1:]

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(result=[])


Ui_종목별투자자정보백업, QtBaseClass_종목별투자자정보백업 = uic.loadUiType(UI_DIR+"종목별투자자정보백업.ui")
class 화면_종목별투자자정보백업(QDialog, Ui_종목별투자자정보백업):
    def __init__(self, parent=None):
        super(화면_종목별투자자정보백업, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('종목별 투자자 정보 백업')
        self.parent = parent

        self.columns = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자','기관계','금융투자','보험','투신','기타금융','은행','연기금등','국가','내외국인','사모펀드','기타법인']

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        XQ_t8436 = t8436(parent=self)
        XQ_t8436.Query(구분='0')

        self.조회건수 = 10
        self.XQ_t1702 = t1702(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1702':
            CTSIDX, CTSDATE, df = result
            df['단축코드'] = self.단축코드[0]
            # print(CTSIDX, CTSDATE)
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 종목별투자자(일자, 종가, 전일대비구분, 전일대비, 등락율, 누적거래량, 사모펀드, 증권, 보험, 투신, 은행, 종금, 기금, 기타법인, 개인, 등록외국인, 미등록외국인, 국가외, 기관, 외인계, 기타계, 단축코드) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                cursor.executemany(query, df.values.tolist())
                conn.commit()

            try:
                if len(df) == int(self.조회건수) and self.radioButton_all.isChecked() == True:
                    QTimer.singleShot(주문지연, lambda: self.Request(result=result))
                else:
                    self.백업한종목수 += 1
                    if len(self.백업할종목코드) > 0:
                        self.단축코드 = self.백업할종목코드.pop(0)
                        self.result = []

                        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
                        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
                        self.label_codename.setText(S)

                        QTimer.singleShot(주문지연, lambda : self.Request([]))
                    else:
                        QMessageBox.about(self, "백업완료","백업을 완료하였습니다..")
            except Exception as e:
                pass

    def Request(self, result=[]):
        if len(result) > 0:
            CTSIDX, CTSDATE, df = result
            self.XQ_t1702.Query(종목코드=self.단축코드[0], 종료일자='', 금액수량구분='0', 매수매도구분='0', 누적구분='0', CTSDATE=CTSDATE, CTSIDX=CTSIDX)
        else:
            # print('%s %s' % (self.단축코드[0], self.단축코드[1]))
            self.XQ_t1702.Query(종목코드=self.단축코드[0], 종료일자='', 금액수량구분='0', 매수매도구분='0', 누적구분='0', CTSDATE='', CTSIDX='')

    def Backup_One(self):
        idx = self.comboBox.currentIndex()

        self.백업한종목수 = 1
        self.백업할종목코드 = []
        self.단축코드 = self.종목코드테이블[idx:idx + 1][['단축코드','종목명']].values[0]
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')
        self.result = []
        self.Request(result=[])

    def Backup_All(self):
        idx = self.comboBox.currentIndex()
        self.백업한종목수 = 1
        self.백업할종목코드 = list(self.종목코드테이블[idx:][['단축코드','종목명']].values)
        self.단축코드 = self.백업할종목코드.pop(0)
        self.기준일자 = self.lineEdit_date.text().strip().replace('-','')

        self.progressBar.setValue(int(self.백업한종목수 / (len(self.종목코드테이블.index) - self.comboBox.currentIndex()) * 100))
        S = '%s %s' % (self.단축코드[0], self.단축코드[1])
        self.label_codename.setText(S)

        self.result = []
        self.Request(result=[])

## ---------------------------------------------------------------------------------------------------------------------
Ui_종목코드, QtBaseClass_종목코드 = uic.loadUiType(UI_DIR+"종목코드조회.ui")
class 화면_종목코드(QDialog, Ui_종목코드):
    def __init__(self, parent=None):
        super(화면_종목코드, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.df = DataFrame()
        self.XQ_t8436 = t8436(parent=self)
        self.XQ_t8436.Query(구분='0')

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.df = result[0]
            self.model.update(self.df)
            for i in range(len(self.df.columns)):
                self.tableView.resizeColumnToContents(i)

    def SaveCode(self):
        with sqlite3.connect(DATABASE) as conn:
            cursor = conn.cursor()
            query = "insert or replace into 종목코드(종목명,단축코드,확장코드,ETF구분,상한가,하한가,전일가,주문수량단위,기준가,구분,증권그룹,기업인수목적회사여부) values(?,?,?,?,?,?,?,?,?,?,?,?)"
            cursor.executemany(query, self.df.values.tolist())
            conn.commit()

        QMessageBox.about(self, "종목코드 생성", " %s 항목의 종목코드를 생성하였습니다." % (len(self.df)))


Ui_업종정보, QtBaseClass_업종정보 = uic.loadUiType(UI_DIR+"업종정보조회.ui")
class 화면_업종정보(QDialog, Ui_업종정보):
    def __init__(self, parent=None):
        super(화면_업종정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종정보 조회')

        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']

        self.result = []

        d = datetime.date.today()

        XQ = t8424(parent=self)
        XQ.Query()

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8424':
            df = result[0]
            df['컬럼'] = df['업종코드'] + " : " + df['업종명']
            df = df.sort_values(['업종코드', '업종명'], ascending=[True, True])
            self.comboBox.addItems(df['컬럼'].values)

        if szTrCode == 't1514':
            CTS일자, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def inquiry(self):
        업종코드 = self.comboBox.currentText()[:3]
        조회건수 = self.lineEdit_date.text().strip().replace('-', '')

        XQ = t1514(parent=self)
        XQ.Query(업종코드=업종코드,구분1='',구분2='1',CTS일자='',조회건수=조회건수,비중구분='', 연속조회=False)


Ui_테마정보, QtBaseClass_테마정보 = uic.loadUiType(UI_DIR+"테마정보조회.ui")
class 화면_테마정보(QDialog, Ui_테마정보):
    def __init__(self, parent=None):
        super(화면_테마정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('테마정보 조회')

        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']

        self.result = []

        d = datetime.date.today()

        XQ = t8425(parent=self)
        XQ.Query()

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8425':
            df = result[0]
            df['컬럼'] = df['테마코드'] + " : " + df['테마명']
            df = df.sort_values(['테마코드', '테마명'], ascending=[True, True])
            self.comboBox.addItems(df['컬럼'].values)

        if szTrCode == 't1537':
            df0, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def inquiry(self):
        테마코드 = self.comboBox.currentText()[:4]

        XQ = t1537(parent=self)
        XQ.Query(테마코드=테마코드, 연속조회=False)


Ui_분별주가조회, QtBaseClass_분별주가조회 = uic.loadUiType(UI_DIR+"분별주가조회.ui")
class 화면_분별주가(QDialog, Ui_분별주가조회):
    def __init__(self, parent=None):
        super(화면_분별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('분별 주가 조회')
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = []

        self.result = []

        XQ = t8436(parent=self)
        XQ.Query(구분='0')

        self.XQ_t1302 = t1302(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1302':
            시간CTS, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def inquiry(self):
        단축코드 = self.comboBox.currentText().strip()[3:9]
        조회건수 = self.lineEdit_cnt.text().strip().replace('-', '')

        self.XQ_t1302.Query(단축코드=단축코드,작업구분='1',시간='',건수=조회건수, 연속조회=False)


Ui_일자별주가조회, QtBaseClass_일자별주가조회 = uic.loadUiType(UI_DIR+"일자별주가조회.ui")
class 화면_일별주가(QDialog, Ui_일자별주가조회):
    def __init__(self, parent=None):
        super(화면_일별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('일자별 주가 조회')

        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['날짜', '시가', '고가', '저가', '종가', '전일대비구분', '전일대비', '등락율', '누적거래량', '거래증가율', '체결강도', '소진율', '회전율',
                   '외인순매수', '기관순매수', '종목코드', '누적거래대금', '개인순매수', '시가대비구분', '시가대비', '시가기준등락율', '고가대비구분', '고가대비',
                   '고가기준등락율', '저가대비구분', '저가대비', '저가기준등락율', '시가총액']

        self.result = []

        d = datetime.date.today()

        XQ = t8436(parent=self)
        XQ.Query(구분='0')

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1305':
            CNT, 날짜, IDX, df = result
            # print(CNT, 날짜, IDX)

            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

            if int(CNT) == int(self.조회건수):
                QTimer.singleShot(주문지연, lambda: self.inquiry_repeatly(result=result))
            else:
                # print("===END===")
                pass

    def inquiry_repeatly(self, result):
        CNT, 날짜, IDX, df = result
        self.XQ.Query(단축코드=self.단축코드, 일주월구분='1', 날짜=날짜, IDX=IDX, 건수=self.조회건수, 연속조회=True)

    def inquiry(self):
        self.단축코드 = self.comboBox.currentText()[3:9]
        self.조회건수 = self.lineEdit_date.text().strip().replace('-', '')

        self.XQ = t1305(parent=self)
        self.XQ.Query(단축코드=self.단축코드,일주월구분='1',날짜='',IDX='',건수=self.조회건수, 연속조회=False)


Ui_종목별투자자조회, QtBaseClass_종목별투자자조회 = uic.loadUiType(UI_DIR+"종목별투자자조회.ui")
class 화면_종목별투자자(QDialog, Ui_종목별투자자조회):
    def __init__(self, parent=None):
        super(화면_종목별투자자, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('종목별 투자자 조회')
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '종가', '전일대비구분', '전일대비', '등락율', '누적거래량', '사모펀드', '증권', '보험', '투신', '은행', '종금', '기금', '기타법인',
                       '개인', '등록외국인', '미등록외국인', '국가외', '기관', '외인계', '기타계']

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        XQ = t8436(parent=self)
        XQ.Query(구분='0')

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1702':
            CTSIDX, CTSDATE, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

    def inquiry(self):
        단축코드 = self.comboBox.currentText()[3:9]
        조회건수 = self.lineEdit_date.text().strip().replace('-', '')

        XQ = t1702(parent=self)
        XQ.Query(종목코드=단축코드, 종료일자='', 금액수량구분='0', 매수매도구분='0', 누적구분='0', CTSDATE='', CTSIDX='')


class 화면_종목별투자자2(QDialog, Ui_종목별투자자조회):
    def __init__(self, parent=None):
        super(화면_종목별투자자2, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('종목별 투자자 조회')
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = []

        self.result = []

        d = datetime.date.today()
        self.lineEdit_date.setText(str(d))

        XQ = t8436(parent=self)
        XQ.Query(구분='0')

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 't1717':
            df = result[0]
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-','')

    def inquiry(self):
        단축코드 = self.comboBox.currentText()[3:9]
        조회건수 = self.lineEdit_date.text().strip().replace('-', '')

        XQ = t1717(parent=self)
        XQ.Query(종목코드=단축코드,구분='0',시작일자='20170101',종료일자='20172131')


Ui_차트인덱스, QtBaseClass_차트인덱스 = uic.loadUiType(UI_DIR+"차트인덱스.ui")
class 화면_차트인덱스(QDialog, Ui_차트인덱스):
    def __init__(self, parent=None):
        super(화면_차트인덱스, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.parent = parent

        self.columns = ['일자', '시간', '시가', '고가', '저가', '종가', '거래량', '지표값1', '지표값2', '지표값3', '지표값4', '지표값5', '위치']

        self.XQ_ChartIndex = ChartIndex(parent=self)
        XQ = t8436(parent=self)
        XQ.Query(구분='0')

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = ">> " + self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

        if szTrCode == 'CHARTINDEX':
            식별자, 지표ID, 레코드갯수, 유효데이터컬럼갯수, self.df = result

            self.model.update(self.df)
            for i in range(len(self.df.columns)):
                self.tableView.resizeColumnToContents(i)

    def OnReceiveChartRealData(self, szTrCode, lst):
        if szTrCode == 'CHARTINDEX':
            식별자, result = lst
            지표ID, 레코드갯수, 유효데이터컬럼갯수, d = result
            lst = [[d['일자'],d['시간'],d['시가'],d['고가'],d['저가'],d['종가'],d['거래량'],d['지표값1'],d['지표값2'],d['지표값3'],d['지표값4'],d['지표값5'],d['위치']]]
            self.df = self.df.append(pd.DataFrame(lst, columns=self.columns), ignore_index=True)

            try:
                self.model.update(self.df)
                for i in range(len(self.df.columns)):
                    self.tableView.resizeColumnToContents(i)
            except Exception as e:
                pass

    def inquiry(self):
        지표명 = self.lineEdit_name.text()
        단축코드 =  self.comboBox.currentText()[3:9]
        요청건수 = self.lineEdit_cnt.text()
        실시간 = '1' if self.checkBox.isChecked() == True else '0'

        self.XQ_ChartIndex.Query(지표ID='', 지표명=지표명, 지표조건설정='', 시장구분='1', 주기구분='0', 단축코드=단축코드, 요청건수=요청건수, 단위='3', 시작일자='',
                 종료일자='', 수정주가반영여부='1', 갭보정여부='1', 실시간데이터수신자동등록여부=실시간)


Ui_종목검색, QtBaseClass_종목검색 = uic.loadUiType(UI_DIR+"종목검색.ui")
class 화면_종목검색(QDialog, Ui_종목검색):
    def __init__(self, parent=None):
        super(화면_종목검색, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.parent = parent

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't1833':
            종목검색수, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def fileselect(self):
        pathname = os.path.dirname(sys.argv[0])
        RESDIR = "%s\\ADF\\" % os.path.abspath(pathname)

        fname = QFileDialog.getOpenFileName(self, 'Open file',RESDIR, "조검검색(*.adf)")
        self.lineEdit.setText(fname[0])

    def inquiry(self):
        filename = self.lineEdit.text()
        XQ = t1833(parent=self)
        XQ.Query(종목검색파일=filename)


Ui_e종목검색, QtBaseClass_e종목검색 = uic.loadUiType(UI_DIR+"e종목검색.ui")
class 화면_e종목검색(QDialog, Ui_e종목검색):
    def __init__(self, parent=None):
        super(화면_e종목검색, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.parent = parent

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't1857':
            검색종목수, 포착시간, 실시간키, df = result
            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)

    def OnReceiveSearchRealData(self, szTrCode, result):
        if szTrCode == 't1857':
            print(result)

    def fileselect(self):
        pathname = os.path.dirname(sys.argv[0])
        RESDIR = "%s\\ACF\\" % os.path.abspath(pathname)

        fname = QFileDialog.getOpenFileName(self, 'Open file',RESDIR, "조검검색(*.acf)")
        self.lineEdit.setText(fname[0])

    def inquiry(self):
        filename = self.lineEdit.text()
        XQ = t1857(parent=self)
        XQ.Query(실시간구분='0',종목검색구분='F',종목검색입력값=filename)


Ui_호가창정보, QtBaseClass_호가창정보 = uic.loadUiType(UI_DIR+"호가창정보.ui")
class 화면_호가창정보(QDialog, Ui_호가창정보):
    def __init__(self, parent=None):
        super(화면_호가창정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent

        self.매도호가컨트롤 = [
            self.label_offerho1, self.label_offerho2, self.label_offerho3, self.label_offerho4, self.label_offerho5,
            self.label_offerho6, self.label_offerho7, self.label_offerho8, self.label_offerho9, self.label_offerho10
        ]

        self.매수호가컨트롤 = [
            self.label_bidho1, self.label_bidho2, self.label_bidho3, self.label_bidho4, self.label_bidho5,
            self.label_bidho6, self.label_bidho7, self.label_bidho8, self.label_bidho9, self.label_bidho10
        ]

        self.매도호가잔량컨트롤 = [
            self.label_offerrem1, self.label_offerrem2, self.label_offerrem3, self.label_offerrem4,
            self.label_offerrem5,
            self.label_offerrem6, self.label_offerrem7, self.label_offerrem8, self.label_offerrem9,
            self.label_offerrem10
        ]

        self.매수호가잔량컨트롤 = [
            self.label_bidrem1, self.label_bidrem2, self.label_bidrem3, self.label_bidrem4, self.label_bidrem5,
            self.label_bidrem6, self.label_bidrem7, self.label_bidrem8, self.label_bidrem9, self.label_bidrem10
        ]

        with sqlite3.connect(DATABASE) as conn:
            query = 'select 단축코드,종목명,ETF구분,구분 from 종목코드'
            df = pdsql.read_sql_query(query, con=conn)

        self.kospi_codes = df.query("구분=='1'")['단축코드'].values.tolist()
        self.kosdaq_codes = df.query("구분=='2'")['단축코드'].values.tolist()

        XQ = t8436(parent=self)
        XQ.Query(구분='0')

        self.kospi_askbid = H1_(parent=self)
        self.kosdaq_askbid = HA_(parent=self)

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print(systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 't8436':
            self.종목코드테이블 = result[0]
            self.종목코드테이블['컬럼'] = self.종목코드테이블['단축코드'] + " : " + self.종목코드테이블['종목명']
            self.종목코드테이블 = self.종목코드테이블.sort_values(['단축코드', '종목명'], ascending=[True, True])
            self.comboBox.addItems(self.종목코드테이블['컬럼'].values)

    def OnReceiveRealData(self, szTrCode, result):
        try:
            s = "%s:%s:%s" % (result['호가시간'][0:2],result['호가시간'][2:4],result['호가시간'][4:6])
            self.label_hotime.setText(s)

            for i in range(0,10):
                self.매도호가컨트롤[i].setText(result['매도호가'][i])
                self.매수호가컨트롤[i].setText(result['매수호가'][i])
                self.매도호가잔량컨트롤[i].setText(result['매도호가잔량'][i])
                self.매수호가잔량컨트롤[i].setText(result['매수호가잔량'][i])

            self.label_offerremALL.setText(result['총매도호가잔량'])
            self.label_bidremALL.setText(result['총매수호가잔량'])
            self.label_donsigubun.setText(result['동시호가구분'])
            self.label_alloc_gubun.setText(result['배분적용구분'])
        except Exception as e:
            pass

    def AddCode(self):
        종목코드 = self.comboBox.currentText().strip()[0:6]

        self.kospi_askbid.UnadviseRealData()
        self.kosdaq_askbid.UnadviseRealData()

        if 종목코드 in self.kospi_codes:
            self.kospi_askbid.AdviseRealData(종목코드=종목코드)
        if 종목코드 in self.kosdaq_codes:
            self.kosdaq_askbid.AdviseRealData(종목코드=종목코드)


Ui_실시간정보, QtBaseClass_실시간정보 = uic.loadUiType(UI_DIR+"실시간정보.ui")
class 화면_실시간정보(QDialog, Ui_실시간정보):
    def __init__(self, parent=None):
        super(화면_실시간정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent

        self.kospi_real = S3_(parent=self)

    def OnReceiveRealData(self, szTrCode, result):
        try:
            str = '{}:{} - {}--{}\r'.format(result['체결시간'], result['단축코드'], result['현재가'], result['체결량'])
            self.textEdit.insertPlainText(str)
        except Exception as e:
            pass

    def AddCode(self):
        종목코드 = self.comboBox.currentText().strip()
        self.comboBox.addItems([종목코드])
        self.kospi_real.AdviseRealData(종목코드=종목코드)

    def RemoveCode(self):
        종목코드 = self.comboBox.currentText().strip()
        self.kospi_real.UnadviseRealDataWithKey(종목코드=종목코드)


Ui_뉴스, QtBaseClass_뉴스 = uic.loadUiType(UI_DIR+"뉴스.ui")
class 화면_뉴스(QDialog, Ui_뉴스):
    def __init__(self, parent=None):
        super(화면_뉴스, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.parent = parent

        self.news = NWS(parent=self)

    def OnReceiveRealData(self, szTrCode, result):
        str = '{}:{} - {}-{}-{}-{}\r'.format(result['날짜'], result['시간'], result['뉴스구분자'], result['키값'], result['단축종목코드'], result['제목'])
        try:
            self.textEdit.insertPlainText(str)
        except Exception as e:
            pass

    def AddCode(self):
        self.news.AdviseRealData()

    def RemoveCode(self):
        self.news.UnadviseRealData()


Ui_주문테스트, QtBaseClass_주문테스트 = uic.loadUiType(UI_DIR+"주문테스트.ui")
class 화면_주문테스트(QDialog, Ui_주문테스트):
    def __init__(self, parent=None):
        super(화면_주문테스트, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent

        self.connection = self.parent.connection

        # 계좌정보 불러오기
        nCount = self.connection.ActiveX.GetAccountListCount()
        for i in range(nCount):
            self.comboBox.addItem(self.connection.ActiveX.GetAccountList(i))

        self.QA_CSPAT00600 = CSPAT00600(parent=self)

        self.setup()

    def setup(self):
        self.XR_SC1 = SC1(parent=self)
        self.XR_SC1.AdviseRealData()
        self.주문번호리스트 = []

    def OnReceiveMessage(self, systemError, messageCode, message):
        self.textEdit.insertPlainText("systemError:[%s] messageCode:[%s] message:[%s]\r" % (systemError, messageCode, message))

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 'CSPAT00600':
            df, df1 = result
            주문번호 = df1['주문번호'].values[0]
            self.textEdit.insertPlainText("주문번호 : %s\r" % 주문번호)
            if 주문번호 != '0':
                # 주문번호처리
                self.주문번호리스트.append(str(주문번호))

    def OnReceiveRealData(self, szTrCode, result):
        try:
            self.textEdit.insertPlainText(szTrCode+'\r')
            self.textEdit.insertPlainText(str(result)+'\r')
        except Exception as e:
            pass

        if szTrCode == 'SC1':
            체결시각 = result['체결시각']
            단축종목번호 = result['단축종목번호'].strip().replace('A','')
            종목명 = result['종목명']
            매매구분 = result['매매구분']
            주문번호 = result['주문번호']
            체결번호 = result['체결번호']
            주문수량 = result['주문수량']
            주문가격 = result['주문가격']
            체결수량 = result['체결수량']
            체결가격 = result['체결가격']
            주문평균체결가격 = result['주문평균체결가격']
            주문계좌번호 = result['주문계좌번호']

            # 내가 주문한 것이 맞을 경우 처리
            if 주문번호 in self.주문번호리스트:
                s = "[%s] %s %s %s %s %s %s %s %s %s %s %s" % (szTrCode,체결시각,단축종목번호,매매구분,주문번호,체결번호,주문수량,주문가격,체결수량,체결가격,주문평균체결가격,주문계좌번호)
                try:
                    self.textEdit.insertPlainText(s + '\r')
                except Exception as e:
                    pass

                일자 = "{:%Y-%m-%d}".format(datetime.datetime.now())
                with sqlite3.connect(DATABASE) as conn:
                    query = 'insert into 거래결과(로봇명, UUID, 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
                    data = ['주문테스트', '주문테스트-UUID', 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격]
                    cursor = conn.cursor()
                    cursor.execute(query, data)
                    conn.commit()

    def Order(self):
        계좌번호 = self.comboBox.currentText().strip()
        비밀번호 = self.lineEdit_pwd.text().strip()
        종목코드 = self.lineEdit_code.text().strip()
        주문가 = self.lineEdit_price.text().strip()
        주문수량 = self.lineEdit_amt.text().strip()
        매매구분 = self.lineEdit_bs.text().strip()
        호가유형 = self.lineEdit_hoga.text().strip()
        신용거래 = self.lineEdit_sin.text().strip()
        주문조건 = self.lineEdit_jogun.text().strip()

        self.QA_CSPAT00600.Query(계좌번호=계좌번호, 입력비밀번호=비밀번호, 종목번호=종목코드, 주문수량=주문수량, 주문가=주문가, 매매구분=매매구분, 호가유형코드=호가유형, 신용거래코드=신용거래, 주문조건구분=주문조건)


Ui_외부신호2eBEST, QtBaseClass_외부신호2eBEST = uic.loadUiType(UI_DIR+"외부신호2eBEST.ui")
class 화면_외부신호2eBEST(QDialog, Ui_외부신호2eBEST):
    def __init__(self, parent=None):
        super(화면_외부신호2eBEST, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.parent = parent

        self.pathname = os.path.dirname(sys.argv[0])
        self.file = "%s\\" % os.path.abspath(self.pathname)

        self.매도 = 1
        self.매수 = 2
        self.매수방법 = '00'
        self.매도방법 = '00'
        self.조건없음 = 0
        self.조건IOC = 1
        self.조건FOK = 2

        self.신용거래코드 = '000'

        self.주문번호리스트 = []
        self.QA_CSPAT00600 = CSPAT00600(parent=self)
        self.XR_SC1 = SC1(parent=self)
        self.XR_SC1.AdviseRealData()

        self.connection = self.parent.connection

        # 계좌정보 불러오기
        nCount = self.connection.ActiveX.GetAccountListCount()
        for i in range(nCount):
            self.comboBox.addItem(self.connection.ActiveX.GetAccountList(i))

    def OnReceiveMessage(self, systemError, messageCode, message):
        s = "\r%s %s %s\r" % (systemError, messageCode, message)
        try:
            self.plainTextEdit.insertPlainText(s)
        except Exception as e:
            pass

    def OnReceiveData(self, szTrCode, result):
        if szTrCode == 'CSPAT00600':
            df, df1 = result
            주문번호 = df1['주문번호'].values[0]
            if 주문번호 != '0':
                self.주문번호리스트.append(str(주문번호))
                s = "주문번호 : %s\r" % 주문번호
                try:
                    self.plainTextEdit.insertPlainText(s)
                except Exception as e:
                    pass

    def OnReceiveRealData(self, szTrCode, result):
        if szTrCode == 'SC1':
            체결시각 = result['체결시각']
            단축종목번호 = result['단축종목번호'].strip().replace('A','')
            종목명 = result['종목명']
            매매구분 = result['매매구분']
            주문번호 = result['주문번호']
            체결번호 = result['체결번호']
            주문수량 = result['주문수량']
            주문가격 = result['주문가격']
            체결수량 = result['체결수량']
            체결가격 = result['체결가격']
            주문평균체결가격 = result['주문평균체결가격']
            주문계좌번호 = result['주문계좌번호']

            # 내가 주문한 것이 체결된 경우 처리
            if 주문번호 in self.주문번호리스트:
                s = "\r주문체결[%s] : %s %s %s %s %s %s %s %s %s %s %s\r" % (szTrCode,체결시각,단축종목번호,매매구분,주문번호,체결번호,주문수량,주문가격,체결수량,체결가격,주문평균체결가격,주문계좌번호)
                try:
                    self.plainTextEdit.insertPlainText(s)
                except Exception as e:
                    pass

                일자 = "{:%Y-%m-%d}".format(datetime.datetime.now())
                with sqlite3.connect(DATABASE) as conn:
                    query = 'insert into 거래결과(로봇명, UUID, 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
                    data = ['툴박스2EBEST', '툴박스2EBEST-UUID', 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격]
                    cursor = conn.cursor()
                    cursor.execute(query, data)
                    conn.commit()


    def OnReadFile(self, line):
        try:
            self.plainTextEdit.insertPlainText("\r>> " +line.strip() + '\r')
        except Exception as e:
            pass

        lst = line.strip().split(',')

        try:
            시각, 종류, 단축코드, 가격, 수량 = lst
            가격 = int(가격)
            수량 = int(수량)

            if 종류 == '매수':
                self.QA_CSPAT00600.Query(계좌번호=self.계좌번호, 입력비밀번호=self.비밀번호, 종목번호=단축코드, 주문수량=수량, 주문가=가격, 매매구분=self.매수, 호가유형코드=self.매수방법, 신용거래코드=self.신용거래코드, 주문조건구분=self.조건없음)
            if 종류 == '매도':
                self.QA_CSPAT00600.Query(계좌번호=self.계좌번호, 입력비밀번호=self.비밀번호, 종목번호=단축코드, 주문수량=수량, 주문가=가격, 매매구분=self.매도, 호가유형코드=self.매도방법, 신용거래코드=self.신용거래코드, 주문조건구분=self.조건없음)
        except Exception as e:
            pass

    def fileselect(self):
        ret = QFileDialog.getOpenFileName(self, 'Open file',self.file, "CSV,TXT(*.csv;*.txt)")
        self.file = ret[0]
        self.lineEdit.setText(self.file)

    def StartWatcher(self):
        self.계좌번호 = self.comboBox.currentText().strip()
        self.비밀번호 = self.lineEdit_pwd.text().strip()

        self.fw = FileWatcher(filename=self.file, callback=self.OnReadFile, encoding='utf-8')
        self.fw.start()


Ui_거래결과, QtBaseClass_거래결과 = uic.loadUiType(UI_DIR+"거래결과.ui")
class 화면_거래결과(QDialog, Ui_거래결과):
    def __init__(self, parent=None):
        super(화면_거래결과, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('거래결과 조회')
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = []

        with sqlite3.connect(DATABASE) as conn:
            query = "select distinct 로봇명 from 거래결과 order by 로봇명"
            df = pdsql.read_sql_query(query, con=conn)
            for name in df['로봇명'].values.tolist():
                self.comboBox.addItem(name)

    def inquiry(self):
        로봇명 = self.comboBox.currentText().strip()
        with sqlite3.connect(DATABASE) as conn:
            query = """
                select 로봇명, UUID, 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격 
                from 거래결과
                where  로봇명='%s'
                order by 일자, 체결시각
            """ % 로봇명
            df = pdsql.read_sql_query(query, con=conn)

            self.model.update(df)
            for i in range(len(df.columns)):
                self.tableView.resizeColumnToContents(i)


Ui_버전, QtBaseClass_버전 = uic.loadUiType(UI_DIR+"버전.ui")
class 화면_버전(QDialog, Ui_버전):
    def __init__(self, parent=None):
        super(화면_버전, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)
        self.setWindowTitle('버전')
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        df = DataFrame(data=프로그램정보,columns=['A','B'])

        self.model.update(df)
        for i in range(len(df.columns)):
            self.tableView.resizeColumnToContents(i)

##################################################################################
# 메인
##################################################################################

Ui_MainWindow, QtBaseClass_MainWindow = uic.loadUiType(UI_DIR+"mymoneybot.ui")

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, autoMode = False, *args, **kwargs):
        self.mAutoMode = autoMode
        super(MainWindow, self).__init__(*args, **kwargs)
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("mymoneybot for eBEST (www.thinkalgo.co.kr)")

        self.plugins = CPluginManager.plugin_loader()
        menuitems = self.plugins.keys()
        menu = self.menubar.addMenu('&플러그인로봇')
        for item in menuitems:
            icon = QIcon()
            icon.addPixmap(QtGui.QPixmap("PNG/approval.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            entry = menu.addAction(icon, item)
            entry.setObjectName(item)

        self.시작시각 = datetime.datetime.now()

        self.robots = []

        self.dialog = dict()

        self.portfolio_columns = ['종목코드', '종목명', 'TAG', '매수가', '수량', '매수일']
        self.robot_columns = ['Robot타입', 'Robot명', 'RobotID', '실행상태', '포트수', '포트폴리오']

        self.model = PandasModel()
        self.tableView_robot.setModel(self.model)
        self.tableView_robot.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_robot.setSelectionMode(QTableView.SingleSelection)

        self.tableView_robot.pressed.connect(self.RobotCurrentIndex)
        self.tableView_robot_current_index = None

        self.portfolio_model = PandasModel()
        self.tableView_portfolio.setModel(self.portfolio_model)
        self.tableView_portfolio.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_portfolio.setSelectionMode(QTableView.SingleSelection)
        self.tableView_portfolio.pressed.connect(self.PortfolioCurrentIndex)
        self.tableView_portfolio_current_index = None

        self.portfolio_model.update((DataFrame(columns=self.portfolio_columns)))

        self.주문제한 = 0
        self.조회제한 = 0
        self.금일백업작업중 = False
        self.종목선정작업중 = False

        self.계좌번호 = None
        self.거래비밀번호 = None

        # AxtiveX 설정
        # self.connection = XASession(parent=self)
        self.connection = None
        self.XQ_t0167 = t0167(parent=self)

    def OnQApplicationStarted(self):
        self.clock = QtCore.QTimer()
        self.clock.timeout.connect(self.OnClockTick)
        self.clock.start(1000)

        try:
            with open('mymoneybot.robot', 'rb') as handle:
                self.robots = pickle.load(handle)
        except Exception as e:
            pass

        self.RobotView()


        #TODO:자동로그인
        try:
            self.MyLogin()
        except Exception as e:
            logger.info("MyLogin's Error: %s", e)


    def OnClockTick(self):
        current = datetime.datetime.now()
        current_str = current.strftime('%H:%M:%S')

        if current.second == 0: # 매 0초
            try:
                if self.connection is not None:
                    msg = '오프라인'
                    if self.connection.IsConnected():
                        msg = "온라인"

                        # 현재시간 조회
                        self.XQ_t0167.Query()
                    else:
                        msg = "오프라인"
                    self.statusbar.showMessage(msg)
            except Exception as e:
                pass

            _temp = []
            for r in self.robots:
                if r.running == True:
                    _temp.append(r.Name)

            if current_str in ['09:01:00']:
                self.RobotRun()
                self.RobotView()

            if current_str in ['15:31:00']:
                self.SaveRobots()
                self.RobotView()

            if current_str[3:] in ['00:00', '30:00']:
                ToTelegram("%s : 로봇 %s개가 실행중입니다. ([%s])" % (current_str, len(_temp), ','.join(_temp)))

            if current.minute % 10 == 0: # 매 10 분
                pass

    def closeEvent(self,event):
        result = QMessageBox.question(self,"프로그램 종료","정말 종료하시겠습니까 ?", QMessageBox.Yes| QMessageBox.No)

        if result == QMessageBox.Yes:
            event.accept()
            self.clock.stop()
            self.SaveRobots()
        else:
            event.ignore()

    def SaveRobots(self):
        for r in self.robots:
            r.Run(flag=False, parent=None)

        try:
            with open('mymoneybot.robot', 'wb') as handle:
                pickle.dump(self.robots, handle, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as e:
            print(e)
        finally:
            for r in self.robots:
                r.Run(flag=False, parent=self)

    def LoadRobots(self):
        with open('mymoneybot.robot', 'rb') as handle:
            try:
                self.robots = pickle.load(handle)
            except Exception as e:
                print(e)
            finally:
                pass

    def robot_selected(self, QModelIndex):
        Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row()+1]['Robot타입'].values[0]

        uuid = self.model._data[QModelIndex.row():QModelIndex.row()+1]['RobotID'].values[0]
        portfolio = None
        for r in self.robots:
            if r.UUID == uuid:
                portfolio = r.portfolio
                model = PandasModel()
                result = []
                for p, v in portfolio.items():
                    result.append((v.종목코드, v.종목명.strip(), p, v.매수가, v.수량, v.매수일))
                self.portfolio_model.update((DataFrame(data=result, columns=['종목코드','종목명','TAG','매수가','수량','매수일'])))

                break

    def robot_double_clicked(self, QModelIndex):
        self.RobotEdit(QModelIndex)
        self.RobotView()

    def portfolio_selected(self, QModelIndex):
        pass

    def portfolio_double_clicked(self, QModelIndex):
        RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1]['RobotID'].values[0]
        Portfolio라벨 = self.portfolio_model._data[self.tableView_portfolio_current_index.row():self.tableView_portfolio_current_index.row() + 1]['TAG'].values[0]

        for r in self.robots:
            if r.UUID == RobotUUID:
                portfolio_keys = list(r.portfolio.keys())
                for k in portfolio_keys:
                    if k == Portfolio라벨:
                        v = r.portfolio[k]
                        result = QMessageBox.question(self, "포트폴리오 종목 삭제", "[%s-%s] 을/를 삭제 하시겠습니까 ?" %(v.종목코드, v.종목명), QMessageBox.Yes | QMessageBox.No)
                        if result == QMessageBox.Yes:
                            r.portfolio.pop(Portfolio라벨)

                        self.PortfolioView()

    def RobotCurrentIndex(self, index):
        self.tableView_robot_current_index = index

    def RobotRun(self):
        for r in self.robots:
            r.초기조건()
            # logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            r.Run(flag=True, parent=self)

    def RobotView(self):
        result = []
        for r in self.robots:
            result.append(r.getstatus())

        self.model.update(DataFrame(data=result, columns=self.robot_columns))

        # RobotID 숨김
        self.tableView_robot.setColumnHidden(2, True)

        for i in range(len(self.robot_columns)):
            self.tableView_robot.resizeColumnToContents(i)

    def RobotEdit(self, QModelIndex):
        Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row()+1]['Robot타입'].values[0]
        RobotUUID = self.model._data[QModelIndex.row():QModelIndex.row()+1]['RobotID'].values[0]

        for r in self.robots:
            if r.UUID == RobotUUID:
                r.modal(parent=self)

    def PortfolioView(self):
        RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1]['RobotID'].values[0]
        portfolio = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                portfolio = r.portfolio
                # model = PandasModel()
                result = []
                for p, v in portfolio.items():
                    매수일 = "%s" % v.매수일
                    result.append((v.종목코드, v.종목명.strip(), p, v.매수가, v.수량, 매수일[:19]))

                df = DataFrame(data=result, columns=self.portfolio_columns)
                df = df.sort_values(['종목명'], ascending=True)
                self.portfolio_model.update(df)

                for i in range(len(self.portfolio_columns)):
                    self.tableView_portfolio.resizeColumnToContents(i)

    def PortfolioCurrentIndex(self, index):
        self.tableView_portfolio_current_index = index

    # ------------------------------------------------------------------------------------------------------------------
    def MyLogin(self):
        계좌정보 = pd.read_csv("secret/passwords_oms1226.csv", converters={'계좌번호': str, '거래비밀번호': str, '비밀번호': str})
        주식계좌정보 = 계좌정보.query("구분 == '거래'")

        if len(주식계좌정보) > 0:
            if self.connection is None:
                self.connection = XASession(parent=self)

            self.계좌번호 = 주식계좌정보['계좌번호'].values[0].strip()
            self.id = 주식계좌정보['사용자ID'].values[0].strip()
            self.pwd = 주식계좌정보['비밀번호'].values[0].strip()
            self.cert = 주식계좌정보['공인인증비밀번호'].values[0].strip()
            self.거래비밀번호 = 주식계좌정보['거래비밀번호'].values[0].strip()
            self.url = 주식계좌정보['url'].values[0].strip()
            self.connection.login(url=self.url, id=self.id, pwd=self.pwd, cert=self.cert)
        else:
            print("secret디렉토리의 passwords.csv 파일에서 거래 계좌를 지정해 주세요")

    def OnLogin(self, code, msg):
        if code == '0000':
            if self.mAutoMode :
                self.RobotRun()
                self.RobotView()
            self.statusbar.showMessage("로그인 되었습니다.(%s:%s)" % ("self.mAutoMode", self.mAutoMode))
        else:
            self.statusbar.showMessage("%s %s" % (code, msg))

    def OnLogout(self):
        self.statusbar.showMessage("로그아웃 되었습니다.")

    def OnDisconnect(self):
        # 로봇 상태 저장
        self.SaveRobots()

        self.statusbar.showMessage("연결이 끊겼습니다.")

        self.connection.login(url='demo.ebestsec.co.kr', id=self.id, pwd=self.pwd, cert=self.cert)

    def OnReceiveMessage(self, systemError, messageCode, message):
        # 클래스이름 = self.__class__.__name__
        # 함수이름 = inspect.currentframe().f_code.co_name
        # print("%s-%s " % (클래스이름, 함수이름), systemError, messageCode, message)
        pass

    def OnReceiveData(self, szTrCode, result):
        # print(szTrCode, result)
        pass

    def OnReceiveRealData(self, szTrCode, result):
        # print(szTrCode, result)
        pass

    # ------------------------------------------------------------------------------------------------------------------
    def MENU_Action(self, qaction):
        logger.debug("Action Slot %s %s " % (qaction.objectName(), qaction.text()))
        _action = qaction.objectName()
        if _action == "actionExit":
            self.connection.disconnect()
            self.close()

        if _action == "actionLogin":
            self.MyLogin()

        if _action == "actionLogout":
            self.connection.logout()
            self.statusbar.showMessage("로그아웃 되었습니다.")

        # 일별가격정보 백업
        if _action == "actionPriceBackupDay":
            if self.dialog.get('일별가격정보백업') is not None:
                try:
                    self.dialog['일별가격정보백업'].show()
                except Exception as e:
                    self.dialog['일별가격정보백업'] = 화면_일별가격정보백업(parent=self)
                    self.dialog['일별가격정보백업'].show()
            else:
                self.dialog['일별가격정보백업'] = 화면_일별가격정보백업(parent=self)
                self.dialog['일별가격정보백업'].show()

        # 분별가격정보 백업
        if _action == "actionPriceBackupMin":
            if self.dialog.get('분별가격정보백업') is not None:
                try:
                    self.dialog['분별가격정보백업'].show()
                except Exception as e:
                    self.dialog['분별가격정보백업'] = 화면_분별가격정보백업(parent=self)
                    self.dialog['분별가격정보백업'].show()
            else:
                self.dialog['분별가격정보백업'] = 화면_분별가격정보백업(parent=self)
                self.dialog['분별가격정보백업'].show()

        # 일별업종정보 백업
        if _action == "actionSectorBackupDay":
            if self.dialog.get('일별업종정보백업') is not None:
                try:
                    self.dialog['일별업종정보백업'].show()
                except Exception as e:
                    self.dialog['일별업종정보백업'] = 화면_일별업종정보백업(parent=self)
                    self.dialog['일별업종정보백업'].show()
            else:
                self.dialog['일별업종정보백업'] = 화면_일별업종정보백업(parent=self)
                self.dialog['일별업종정보백업'].show()

        # 종목별 투자자정보 백업
        if _action == "actionInvestorBackup":
            if self.dialog.get('종목별투자자정보백업') is not None:
                try:
                    self.dialog['종목별투자자정보백업'].show()
                except Exception as e:
                    self.dialog['종목별투자자정보백업'] = 화면_종목별투자자정보백업(parent=self)
                    self.dialog['종목별투자자정보백업'].show()
            else:
                self.dialog['종목별투자자정보백업'] = 화면_종목별투자자정보백업(parent=self)
                self.dialog['종목별투자자정보백업'].show()

        # 종목코드 조회/저장
        if _action == "actionStockcode":
            if self.dialog.get('종목코드조회') is not None:
                try:
                    self.dialog['종목코드조회'].show()
                except Exception as e:
                    self.dialog['종목코드조회'] = 화면_종목코드(parent=self)
                    self.dialog['종목코드조회'].show()
            else:
                self.dialog['종목코드조회'] = 화면_종목코드(parent=self)
                self.dialog['종목코드조회'].show()

        # 거래결과
        if _action == "actionTool2ebest":
            if self.dialog.get('외부신호2eBEST') is not None:
                try:
                    self.dialog['외부신호2eBEST'].show()
                except Exception as e:
                    self.dialog['외부신호2eBEST'] = 화면_외부신호2eBEST(parent=self)
                    self.dialog['외부신호2eBEST'].show()
            else:
                self.dialog['외부신호2eBEST'] = 화면_외부신호2eBEST(parent=self)
                self.dialog['외부신호2eBEST'].show()

        if _action == "actionTradeResult":
            if self.dialog.get('거래결과') is not None:
                try:
                    self.dialog['거래결과'].show()
                except Exception as e:
                    self.dialog['거래결과'] = 화면_거래결과(parent=self)
                    self.dialog['거래결과'].show()
            else:
                self.dialog['거래결과'] = 화면_거래결과(parent=self)
                self.dialog['거래결과'].show()

        # 일자별 주가
        if _action == "actionDailyPrice":
            if self.dialog.get('일자별주가') is not None:
                try:
                    self.dialog['일자별주가'].show()
                except Exception as e:
                    self.dialog['일자별주가'] = 화면_일별주가(parent=self)
                    self.dialog['일자별주가'].show()
            else:
                self.dialog['일자별주가'] = 화면_일별주가(parent=self)
                self.dialog['일자별주가'].show()

        # 분별 주가
        if _action == "actionMinuitePrice":
            if self.dialog.get('분별주가') is not None:
                try:
                    self.dialog['분별주가'].show()
                except Exception as e:
                    self.dialog['분별주가'] = 화면_분별주가(parent=self)
                    self.dialog['분별주가'].show()
            else:
                self.dialog['분별주가'] = 화면_분별주가(parent=self)
                self.dialog['분별주가'].show()

        # 업종정보
        if _action == "actionSectorView":
            if self.dialog.get('업종정보조회') is not None:
                try:
                    self.dialog['업종정보조회'].show()
                except Exception as e:
                    self.dialog['업종정보조회'] = 화면_업종정보(parent=self)
                    self.dialog['업종정보조회'].show()
            else:
                self.dialog['업종정보조회'] = 화면_업종정보(parent=self)
                self.dialog['업종정보조회'].show()

        # 테마정보
        if _action == "actionTheme":
            if self.dialog.get('테마정보조회') is not None:
                try:
                    self.dialog['테마정보조회'].show()
                except Exception as e:
                    self.dialog['테마정보조회'] = 화면_테마정보(parent=self)
                    self.dialog['테마정보조회'].show()
            else:
                self.dialog['테마정보조회'] = 화면_테마정보(parent=self)
                self.dialog['테마정보조회'].show()

        # 종목별 투자자
        if _action == "actionInvestors":
            if self.dialog.get('종목별투자자') is not None:
                try:
                    self.dialog['종목별투자자'].show()
                except Exception as e:
                    self.dialog['종목별투자자'] = 화면_종목별투자자(parent=self)
                    self.dialog['종목별투자자'].show()
            else:
                self.dialog['종목별투자자'] = 화면_종목별투자자(parent=self)
                self.dialog['종목별투자자'].show()

        # 종목별 투자자2
        if _action == "actionInvestors2":
            if self.dialog.get('종목별투자자2') is not None:
                try:
                    self.dialog['종목별투자자2'].show()
                except Exception as e:
                    self.dialog['종목별투자자2'] = 화면_종목별투자자2(parent=self)
                    self.dialog['종목별투자자2'].show()
            else:
                self.dialog['종목별투자자2'] = 화면_종목별투자자2(parent=self)
                self.dialog['종목별투자자2'].show()

        # 호가창정보
        if _action == "actionAskBid":
            if self.dialog.get('호가창정보') is not None:
                try:
                    self.dialog['호가창정보'].show()
                except Exception as e:
                    self.dialog['호가창정보'] = 화면_호가창정보(parent=self)
                    self.dialog['호가창정보'].show()
            else:
                self.dialog['호가창정보'] = 화면_호가창정보(parent=self)
                self.dialog['호가창정보'].show()

        # 실시간정보
        if _action == "actionRealDataDialog":
            if self.dialog.get('실시간정보') is not None:
                try:
                    self.dialog['실시간정보'].show()
                except Exception as e:
                    self.dialog['실시간정보'] = 화면_실시간정보(parent=self)
                    self.dialog['실시간정보'].show()
            else:
                self.dialog['실시간정보'] = 화면_실시간정보(parent=self)
                self.dialog['실시간정보'].show()

        # 뉴스
        if _action == "actionNews":
            if self.dialog.get('뉴스') is not None:
                try:
                    self.dialog['뉴스'].show()
                except Exception as e:
                    self.dialog['뉴스'] = 화면_뉴스(parent=self)
                    self.dialog['뉴스'].show()
            else:
                self.dialog['뉴스'] = 화면_뉴스(parent=self)
                self.dialog['뉴스'].show()

        # 계좌정보 조회
        if _action == "actionAccountDialog":
            if self.dialog.get('계좌정보조회') is not None:
                try:
                    self.dialog['계좌정보조회'].show()
                except Exception as e:
                    self.dialog['계좌정보조회'] = 화면_계좌정보(parent=self)
                    self.dialog['계좌정보조회'].show()
            else:
                self.dialog['계좌정보조회'] = 화면_계좌정보(parent=self)
                self.dialog['계좌정보조회'].show()

        # 차트인덱스
        if _action == "actionChartIndex":
            if self.dialog.get('차트인덱스') is not None:
                try:
                    self.dialog['차트인덱스'].show()
                except Exception as e:
                    self.dialog['차트인덱스'] = 화면_차트인덱스(parent=self)
                    self.dialog['차트인덱스'].show()
            else:
                self.dialog['차트인덱스'] = 화면_차트인덱스(parent=self)
                self.dialog['차트인덱스'].show()

        # 종목검색
        if _action == "actionSearchItems":
            if self.dialog.get('종목검색') is not None:
                try:
                    self.dialog['종목검색'].show()
                except Exception as e:
                    self.dialog['종목검색'] = 화면_종목검색(parent=self)
                    self.dialog['종목검색'].show()
            else:
                self.dialog['종목검색'] = 화면_종목검색(parent=self)
                self.dialog['종목검색'].show()

        # e종목검색
        if _action == "actionESearchItems":
            if self.dialog.get('e종목검색') is not None:
                try:
                    self.dialog['e종목검색'].show()
                except Exception as e:
                    self.dialog['e종목검색'] = 화면_e종목검색(parent=self)
                    self.dialog['e종목검색'].show()
            else:
                self.dialog['e종목검색'] = 화면_e종목검색(parent=self)
                self.dialog['e종목검색'].show()

        if _action == "actionOpenScreen":
            XQ = t8430(parent=self)
            XQ.Query(구분='0')

            res = XQ.RequestLinkToHTS("&STOCK_CODE", "069500", "")

        # 주문테스트
        if _action == "actionOrder":
            if self.dialog.get('주문테스트') is not None:
                try:
                    self.dialog['주문테스트'].show()
                except Exception as e:
                    self.dialog['주문테스트'] = 화면_주문테스트(parent=self)
                    self.dialog['주문테스트'].show()
            else:
                self.dialog['주문테스트'] = 화면_주문테스트(parent=self)
                self.dialog['주문테스트'].show()

        # 사용법
        if _action == "actionMustRead":
            webbrowser.open('https://thinkpoolost.wixsite.com/moneybot')

        if _action == "actionUsage":
            webbrowser.open('https://docs.google.com/document/d/1BGENxWqJyZdihQFuWcmTNy3_4J0kHolCc-qcW3RULzs/edit')

        if _action == "actionVersion":
            if self.dialog.get('Version') is not None:
                try:
                    self.dialog['Version'].show()
                except Exception as e:
                    self.dialog['Version'] = 화면_버전(parent=self)
                    self.dialog['Version'].show()
            else:
                self.dialog['Version'] = 화면_버전(parent=self)
                self.dialog['Version'].show()

        if _action == "actionRobotLoad":
            reply = QMessageBox.question(self, "로봇 탑제", "저장된 로봇을 읽어올까요?", QMessageBox.Yes | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.Yes:
                self.LoadRobots()

            self.RobotView()

        elif _action == "actionRobotSave":
            reply = QMessageBox.question(self, "로봇 저장", "현재 로봇을 저장할까요?",
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                self.SaveRobots()

            self.RobotView()

        elif _action == "actionRobotOneRun":
            try:
                RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1]['RobotID'].values[0]
            except Exception as e:
                RobotUUID = ''

            robot_found = None
            for r in self.robots:
                if r.UUID == RobotUUID:
                    robot_found = r
                    break

            if robot_found == None:
                return

            robot_found.Run(flag=True, parent=self)

            self.RobotView()

        elif _action == "actionRobotOneStop":
            try:
                RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1]['RobotID'].values[0]
            except Exception as e:
                RobotUUID = ''

            robot_found = None
            for r in self.robots:
                if r.UUID == RobotUUID:
                    robot_found = r
                    break

            if robot_found == None:
                return

            reply = QMessageBox.question(self,"로봇 실행 중지", "로봇 실행을 중지할까요?\n%s" % robot_found.getstatus(),QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                robot_found.Run(flag=False, parent=None)

            self.RobotView()

        elif _action == "actionRobotRun":
            self.RobotRun()
            self.RobotView()

        elif _action == "actionRobotStop":
            reply = QMessageBox.question(self,"전체 로봇 실행 중지", "전체 로봇 실행을 중지할까요?",QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                for r in self.robots:
                    r.Run(flag=False, parent=None)

            self.RobotView()

        elif _action == "actionRobotRemove":
            try:
                RobotUUID = self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1]['RobotID'].values[0]

                robot_found = None
                for r in self.robots:
                    if r.UUID == RobotUUID:
                        robot_found = r
                        break

                if robot_found == None:
                    return

                reply = QMessageBox.question(self, "로봇 삭제", "로봇을 삭제할까요?\n%s" % robot_found.getstatus()[0:4], QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
                if reply == QMessageBox.Cancel:
                    pass
                elif reply == QMessageBox.No:
                    pass
                elif reply == QMessageBox.Yes:
                    self.robots.remove(robot_found)

                self.RobotView()
            except Exception as e:
                pass

        elif _action == "actionRobotClear":
            reply = QMessageBox.question(self, "로봇 전체 삭제", "로봇 전체를 삭제할까요?",
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                self.robots = []

            self.RobotView()

        elif _action == "actionRobotView":
            self.RobotView()
            for r in self.robots:
                logger.debug('%s %s %s %s' % (r.Name, r.UUID, len(r.portfolio), r.getstatus()))

        if _action in self.plugins.keys():
            robot = self.plugins[_action].instance()
            robot.set_database(database=DATABASE)
            robot.set_secret(계좌번호=self.계좌번호, 비밀번호=self.거래비밀번호)
            ret = robot.modal(parent=self)
            if ret == 1:
                self.robots.append(robot)
            self.RobotView()

    # ------------------------------------------------------------

if __name__ == "__main__":
    # Window 8, 10
    # Window 7은 한글을 못읽음
    # Speak("이베스트 API 프로그램을 시작합니다.")

    ToTelegram("mymoneybot for eBEST가 실행되었습니다.")

    # 1.로그 인스턴스를 만든다.
    logger = logging.getLogger('mymoneybot')
    # 2.formatter를 만든다.
    formatter = logging.Formatter('[%(levelname)s|%(filename)s:%(lineno)s]%(asctime)s>%(message)s')

    loggerLevel = logging.DEBUG
    filename = "LOG/mymoneybot.log"

    # 스트림과 파일로 로그를 출력하는 핸들러를 각각 만든다.
    filehandler = logging.FileHandler(filename)
    streamhandler = logging.StreamHandler()

    # 각 핸들러에 formatter를 지정한다.
    filehandler.setFormatter(formatter)
    streamhandler.setFormatter(formatter)

    # 로그 인스턴스에 스트림 핸들러와 파일 핸들러를 붙인다.
    logger.addHandler(filehandler)
    logger.addHandler(streamhandler)
    logger.setLevel(loggerLevel)
    logger.debug("=============================================================================")
    logger.info("LOG START")

    AUTOMODE = False
    while len(sys.argv) > 1:
        if len(sys.argv) > 1 and '-a' in sys.argv[1]:
            AUTOMODE = True

        sys.argv.pop(1)

    logger.debug("%s:%s" % ("AUTOMODE", AUTOMODE))

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    window = MainWindow(autoMode = AUTOMODE)
    window.show()

    QTimer().singleShot(3, window.OnQApplicationStarted)

    sys.exit(app.exec_())

