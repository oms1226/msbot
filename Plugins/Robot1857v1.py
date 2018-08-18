import os, sys
sys.path.append('..')
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

ROBOT_NAME = "Robot1857v1"

Ui_Robot1857v1, QtBaseClass_Robot1857v1 = uic.loadUiType("%s\\Plugins\\Robot1857v1.ui" % __PLUGINDIR__)
class CUIRobot1857v1(QDialog, Ui_Robot1857v1):
    def __init__(self, parent=None):
        super(__class__, self).__init__(parent)
        self.setupUi(self)
        self.parent = parent

    def SearchFile(self):
        pathname = os.path.dirname(sys.argv[0])
        RESDIR = "%s\\ACF\\" % os.path.abspath(pathname)

        fname = QFileDialog.getOpenFileName(self, 'Open file',RESDIR, "조검검색(*.acf)")
        self.lineEdit_filename.setText(fname[0])


class Robot1857v1(CRobot):
    def instance(self):
        UUID = uuid.uuid4().hex
        return Robot1857v1(Name=ROBOT_NAME, UUID=UUID)

    def __init__(self, Name, UUID):
        super(__class__, self).__init__(Name, UUID)
        self.parent = None

        self.단위투자금 = 100 * 10000
        self.매수방법 = '00'
        self.매도방법 = '00'
        self.시장가 = '03'
        self.포트폴리오수 = 10
        self.profitstop = 0.15
        self.losscut = 0.05
        self.ACF파일 = ''
        self.일괄매도시각 = '15:15:00'
        self.매수거래시간STR = '''09:00:00-11:00:00,
12:00:00-13:00:00,
14:00:00-15:20:00'''
        self.매도거래시간STR = '''09:00:00-11:00:00,
12:00:00-13:00:00,
14:00:00-15:20:00'''
        self.매수거래중 = False
        self.매도거래중 = False
        self.금일매도종목 = []

        self.주문번호리스트 = []
        self.매수Lock = dict()
        self.매도Lock = dict()

        self.QueryInit()

        self.clock = None
        self.전량매도 = False

    def QueryInit(self):
        self.XQ_t1857 = None
        self.XR_S3_ = None
        self.XR_K3_ = None
        self.QA_CSPAT00600 = None

        self.XR_SC1 = None # 체결

    def modal(self, parent):
        ui = CUIRobot1857v1(parent=parent)
        ui.setModal(True)

        ui.lineEdit_name.setText(self.Name)
        ui.lineEdit_unit.setText(str(self.단위투자금 // 10000))
        ui.lineEdit_profitstop.setText(str(self.profitstop))
        ui.lineEdit_losscut.setText(str(self.losscut))
        ui.lineEdit_portsize.setText(str(self.포트폴리오수))
        ui.comboBox_buy_sHogaGb.setCurrentIndex(ui.comboBox_buy_sHogaGb.findText(self.매수방법, flags=Qt.MatchContains))
        ui.comboBox_sell_sHogaGb.setCurrentIndex(ui.comboBox_sell_sHogaGb.findText(self.매도방법, flags=Qt.MatchContains))
        ui.lineEdit_filename.setText(self.ACF파일)
        ui.plainTextEdit_buytime.setPlainText(self.매수거래시간STR)
        ui.plainTextEdit_selltime.setPlainText(self.매도거래시간STR)
        ui.lineEdit_sellall.setText(self.일괄매도시각)

        r = ui.exec_()
        if r == 1:
            self.Name = ui.lineEdit_name.text().strip()
            self.단위투자금 = int(ui.lineEdit_unit.text().strip()) * 10000
            self.매수방법 = ui.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            self.매도방법 = ui.comboBox_sell_sHogaGb.currentText().strip()[0:2]
            self.포트폴리오수 = int(ui.lineEdit_portsize.text().strip())
            self.ACF파일 = ui.lineEdit_filename.text().strip()
            self.profitstop = float(ui.lineEdit_profitstop.text().strip())
            self.losscut = float(ui.lineEdit_losscut.text().strip())
            self.매수거래시간STR = ui.plainTextEdit_buytime.toPlainText().strip()
            self.매도거래시간STR = ui.plainTextEdit_selltime.toPlainText().strip()

            매수거래시간1 = self.매수거래시간STR
            매수거래시간2 = [x.strip() for x in 매수거래시간1.split(',')]

            result = []
            for temp in 매수거래시간2:
                result.append([x.strip() for x in temp.split('-')])

            self.매수거래시간 = result

            매도거래시간1 = self.매도거래시간STR
            매도거래시간2 = [x.strip() for x in 매도거래시간1.split(',')]

            result = []
            for temp in 매도거래시간2:
                result.append([x.strip() for x in temp.split('-')])

            self.매도거래시간 = result

            self.일괄매도시각 = ui.lineEdit_sellall.text().strip()

        return r

    def OnReceiveMessage(self, systemError, messageCode, message):
        일자 = "{:%Y-%m-%d %H:%M:%S.%f}".format(datetime.datetime.now())
        클래스이름 = self.__class__.__name__
        print(일자, 클래스이름, systemError, messageCode, message)

    def OnReceiveData(self, szTrCode, result):
        # 종목검색
        if szTrCode == 't1857':
            if self.running:
                식별자, 검색종목수, 포착시간, 실시간키, df = result
                if 식별자 == self.XQ_t1857.식별자:
                    for idx, row in df[['종목코드','종목상태']].iterrows():
                        code, flag = row
                        if type(code) == str:
                            if code in self.kospi_codes and flag in ['N','R']:
                                if type(self.XR_S3_) is not type(None):
                                    self.XR_S3_.AdviseRealData(종목코드=code)
                            if code in self.kospi_codes and flag in ['O']:
                                if type(self.XR_S3_) is not type(None):
                                    if code not in self.portfolio.keys() and code not in self.매수Lock.keys() and code not in self.매도Lock.keys():
                                        self.XR_S3_.UnadviseRealDataWithKey(종목코드=code)
                            if code in self.kosdaq_codes and flag in ['N','R']:
                                if type(self.XR_K3_) is not type(None):
                                    self.XR_K3_.AdviseRealData(종목코드=code)
                            if code in self.kosdaq_codes and flag in ['O']:
                                if type(self.XR_K3_) is not type(None):
                                    if code not in self.portfolio.keys() and code not in self.매수Lock.keys() and code not in self.매도Lock.keys():
                                        self.XR_K3_.UnadviseRealDataWithKey(종목코드=code)

                    # 현재 가지고 있는 포트폴리오의 실시간데이타를 받는다.
                    for code in self.portfolio.keys():
                        if code in self.kospi_codes:
                            if type(self.XR_S3_) is not type(None):
                                self.XR_S3_.AdviseRealData(종목코드=code)
                        if code in self.kosdaq_codes:
                            if type(self.XR_K3_) is not type(None):
                                self.XR_K3_.AdviseRealData(종목코드=code)
						
        # 체결
        if szTrCode == 'CSPAT00600':
            df, df1 = result
            if len(df1) > 0:
                주문번호 = df1['주문번호'].values[0]

                if 주문번호 != '0':
                    # 주문번호처리
                    self.주문번호리스트.append(str(주문번호))

    def OnReceiveSearchRealData(self, szTrCode, lst):
        식별자, result = lst
        if 식별자 == self.XQ_t1857.식별자:
            try:
                code = result['종목코드']
                flag = result['종목상태']
                if type(code) == str:
                    if code in self.kospi_codes and flag in ['N', 'R']:
                        if type(self.XR_S3_) is not type(None):
                            self.XR_S3_.AdviseRealData(종목코드=code)
                    if code in self.kospi_codes and flag in ['O']:
                        if type(self.XR_S3_) is not type(None):
                            if code not in self.portfolio.keys() and code not in self.매수Lock.keys() and code not in self.매도Lock.keys():
                                self.XR_S3_.UnadviseRealDataWithKey(종목코드=code)
                    if code in self.kosdaq_codes and flag in ['N', 'R']:
                        if type(self.XR_K3_) is not type(None):
                            self.XR_K3_.AdviseRealData(종목코드=code)
                    if code in self.kosdaq_codes and flag in ['O']:
                        if type(self.XR_K3_) is not type(None):
                            if code not in self.portfolio.keys() and code not in self.매수Lock.keys() and code not in self.매도Lock.keys():
                                self.XR_K3_.UnadviseRealDataWithKey(종목코드=code)
            except Exception as e:
                print(e)
            finally:
                pass

    def OnReceiveRealData(self, szTrCode, result):
        if szTrCode == 'SC1':
            체결시각 = result['체결시각']
            단축종목번호 = result['단축종목번호'].strip().replace('A','')
            종목명 = result['종목명']
            매매구분 = result['매매구분']
            주문번호 = result['주문번호']
            체결번호 = result['체결번호']
            주문수량 = int(result['주문수량'])
            주문가격 = int(result['주문가격'])
            체결수량 = int(result['체결수량'])
            체결가격 = int(result['체결가격'])
            주문평균체결가격 = int(result['주문평균체결가격'])
            주문계좌번호 = result['주문계좌번호']

            # 내가 주문한 것이 체결된 경우 처리
            if 주문번호 in self.주문번호리스트:
                if 매매구분 == '1' or 매매구분 == 1: # 매도
                    P = self.portfolio.get(단축종목번호, None)
                    if P != None:
                        P.수량 = P.수량 - 체결수량
                        if P.수량 == 0:
                            self.portfolio.pop(단축종목번호)
                            self.매도Lock.pop(단축종목번호)

                            #TODO: 빠른거래시 화면의 응답속도도 영향을 주므로 일단은 커멘트처리
                            # self.parent.RobotView()
                            # ToTelegram(__class__.__name__ + "매도 : %s 체결수량:%s 체결가격:%s" % (종목명, 주문수량, 주문평균체결가격))
                    else:
                        print("매도 주문이 없는데 매도가 들어옴")

                if 매매구분 == '2' or 매매구분 == 2: # 매수
                    P = self.portfolio.get(단축종목번호, None)
                    if P== None:
                        self.portfolio[단축종목번호] = CPortStock(종목코드=단축종목번호, 종목명=종목명, 매수가=주문평균체결가격, 수량=체결수량, 매수일=datetime.datetime.now())
                        if P.수량 == 주문수량:
                            self.매수Lock.pop(단축종목번호)
                    else:
                        P.수량 = P.수량 + 체결수량
                        if P.수량 == 주문수량:
                            self.매수Lock.pop(단축종목번호)

                    # 조건검색과 체결사이에 시간 간격차 때문에 등록이 안되어 있을수도 있음
                    # 체결된 종목은 실시간 가격을 받는다.
                    if 단축종목번호 in self.kospi_codes:
                        if type(self.XR_S3_) is not type(None):
                            self.XR_S3_.AdviseRealData(종목코드=단축종목번호)
                    if 단축종목번호 in self.kosdaq_codes:
                        if type(self.XR_K3_) is not type(None):
                            self.XR_K3_.AdviseRealData(종목코드=단축종목번호)

                if self.parent is not None:
                    self.parent.RobotView()

                일자 = "{:%Y-%m-%d}".format(datetime.datetime.now())
                data = [self.Name, self.UUID, 일자, 체결시각, 단축종목번호, 종목명, 매매구분, 주문번호, 체결번호, 주문수량, 주문가격, 체결수량, 체결가격, 주문평균체결가격]
                self.체결기록(data=data)

        if szTrCode in ['K3_', 'S3_']:
            if self.매수거래중 == True or self.매도거래중 == True:
                단축코드 = result['단축코드']
                try:
                    종목명 = self.종목코드테이블.query("단축코드=='%s'" % 단축코드)['종목명'].values[0]
                except Exception as e:
                    종목명 = ''
                현재가 = result['현재가']
                고가 = result['고가']
                수량 = self.단위투자금 // 현재가

                현재시각 = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if self.parent is not None:
                    self.parent.statusbar.showMessage("[%s]%s %s %s" % (현재시각, 단축코드, 종목명, 현재가))

                P = self.portfolio.get(단축코드, None)
                매수락 = self.매수Lock.get(단축코드, None)
                매도락 = self.매도Lock.get(단축코드, None)

                if P == None:
                    if 단축코드 not in self.금일매도종목 and 수량 > 0:
                        if self.매수거래중 == True:
                            lst = set(self.portfolio.keys()).union(self.매수Lock.keys())
                            if len(lst) < self.포트폴리오수:
                                if 매수락 == None:
                                    self.매수Lock[단축코드] = ''
                                    # 포트폴리오에 없으므로 현재가에 매수
                                    self.QA_CSPAT00600.Query(계좌번호=self.계좌번호,입력비밀번호=self.비밀번호,종목번호=단축코드,주문수량=수량,주문가=현재가,매매구분=self.매수,
                                        호가유형코드=self.매수방법, 신용거래코드=self.신용거래코드,주문조건구분=self.조건없음)
                                    # ToTelegram(__class__.__name__ + "매수주문 : %s %s 주문수량:%s 주문가격:%s" % (단축코드, 종목명, 수량, 현재가))

                else:
                    if self.매도거래중 == True:
                        if 현재가 > P.매수후고가:
                            P.매수후고가 = 현재가

                        if 매도락 == None:
                            수량 = P.수량
                            # 포트폴리오의 수익률이 지정한 이상이면 매도
                            if 현재가 > P.매수가 * (1+self.profitstop):
                                self.매도Lock[단축코드] = ''
                                self.금일매도종목.append(단축코드)
                                self.QA_CSPAT00600.Query(계좌번호=self.계좌번호,입력비밀번호=self.비밀번호,종목번호=단축코드,주문수량=수량,주문가=현재가,매매구분=self.매도,
                                    호가유형코드=self.매도방법, 신용거래코드=self.신용거래코드,주문조건구분=self.조건없음)
                                #TODO: 주문이 연속적으로 나가는 경우
                                # 텔레그렘의 메세지 전송속도가 약 1초이기 때문에
                                # 이베스트에서 오는 신호를 놓치는 경우가 있다.
                                # ToTelegram(__class__.__name__ + "매도주문 : %s %s 주문수량:%s 주문가격:%s" % (단축코드, 종목명, 수량, 현재가))

                            # 손절
                            if 현재가 < P.매수가 * (1-self.losscut):
                                self.매도Lock[단축코드] = ''
                                self.금일매도종목.append(단축코드)								
                                self.QA_CSPAT00600.Query(계좌번호=self.계좌번호,입력비밀번호=self.비밀번호,종목번호=단축코드,주문수량=수량,주문가=현재가,매매구분=self.매도,
                                    호가유형코드=self.시장가, 신용거래코드=self.신용거래코드,주문조건구분=self.조건없음)
                                #TODO: 주문이 연속적으로 나가는 경우
                                # 텔레그렘의 메세지 전송속도가 약 1초이기 때문에
                                # 이베스트에서 오는 신호를 놓치는 경우가 있다.
                                # ToTelegram(__class__.__name__ + "매도주문 : %s %s 주문수량:%s 주문가격:%s" % (단축코드, 종목명, 수량, 현재가))

    def OnClockTick(self):
        current = datetime.datetime.now()
        current_str = current.strftime('%H:%M:%S')

        거래중 = False
        for t in self.매수거래시간:
            if t[0] <= current_str and current_str <= t[1]:
                거래중 = True
        self.매수거래중 = 거래중

        거래중 = False
        for t in self.매도거래시간:
            if t[0] <= current_str and current_str <= t[1]:
                거래중 = True
        self.매도거래중 = 거래중

        #TODO: 특정시간의 강제매도
        #------------------------------------------------------------
        if self.일괄매도시각.strip() is not "":
            if self.일괄매도시각 < current_str and self.전량매도 == False:
                self.전량매도 = True
                #TODO:취소주문 ???
                for k,v in self.portfolio.items():
                    단축코드 = v.종목코드
                    수량 = v.수량
                    종목명 = v.종목명
                    주문가 = '0'
                    호가유형코드 = '03'
                    self.매도Lock[단축코드] = ''
                    self.금일매도종목.append(단축코드)
                    self.QA_CSPAT00600.Query(계좌번호=self.계좌번호, 입력비밀번호=self.비밀번호, 종목번호=단축코드, 주문수량=수량, 주문가=주문가, 매매구분=self.매도,
                                            호가유형코드=호가유형코드, 신용거래코드=self.신용거래코드, 주문조건구분=self.조건없음)
                    # ToTelegram(__class__.__name__ + "일괄매도 : %s %s 주문수량:%s 주문가격:%s" % (단축코드, 종목명, 수량, 주문가))

    def Run(self, flag=True, parent=None):
        if self.running == flag:
            return

        self.parent = parent
        self.running = flag
        ret = 0
        if flag == True:
            ToTelegram("로직 [%s]을 시작합니다." % (__class__.__name__))

            self.clock = QtCore.QTimer()
            self.clock.timeout.connect(self.OnClockTick)
            self.clock.start(1000)
            self.전량매도 = False

            self.금일매도종목 = []
            self.주문번호리스트 = []
            self.매수Lock = dict()
            self.매도Lock = dict()

            with sqlite3.connect(self.DATABASE) as conn:
                query = 'select 단축코드,종목명,ETF구분,구분 from 종목코드'
                self.종목코드테이블 = pdsql.read_sql_query(query, con=conn)
                self.kospi_codes = self.종목코드테이블.query("구분=='1'")['단축코드'].values.tolist()
                self.kosdaq_codes = self.종목코드테이블.query("구분=='2'")['단축코드'].values.tolist()

            self.XQ_t1857 = t1857(parent=self, 식별자=uuid.uuid4().hex)
            self.XQ_t1857.Query(실시간구분='1', 종목검색구분='F', 종목검색입력값=self.ACF파일)

            self.QA_CSPAT00600 = CSPAT00600(parent=self)

            self.XR_S3_ = S3_(parent=self)
            self.XR_K3_ = K3_(parent=self)

            self.XR_SC1 = SC1(parent=self)

            self.XR_SC1.AdviseRealData()

        else:
            if self.XQ_t1857 is not None:
                self.XQ_t1857.RemoveService()
                self.XQ_t1857 = None

            if self.clock is not None:
                try:
                    self.clock.stop()
                except Exception as e:
                    pass
                finally:
                    self.clock = None

            try:
                if self.XR_S3_ != None:
                    self.XR_S3_.UnadviseRealData()
            except Exception as e:
                pass
            finally:
                self.XR_S3_ = None

            try:
                if self.XR_K3_ != None:
                    self.XR_K3_.UnadviseRealData()
            except Exception as e:
                pass
            finally:
                self.XR_K3_ = None

            try:
                if self.XR_SC1 != None:
                    self.XR_SC1.UnadviseRealData()
            except Exception as e:
                pass
            finally:
                self.XR_SC1 = None

            self.QueryInit()


def robot_loader():
    UUID = uuid.uuid4().hex
    robot = Robot1857v1(Name=ROBOT_NAME, UUID=UUID)
    return robot