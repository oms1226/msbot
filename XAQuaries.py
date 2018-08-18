# -*- coding: utf-8 -*-

import sys, os
import datetime, time
import win32com.client
import pythoncom
import inspect

import pandas as pd
from pandas import Panel, DataFrame, Series

class XAQueryEvents(object):
    def __init__(self):
        self.parent = None

    def set_parent(self, parent):
        self.parent = parent

    def OnReceiveMessage(self, systemError, messageCode, message):
        if self.parent != None:
            self.parent.OnReceiveMessage(systemError, messageCode, message)

    def OnReceiveData(self, szTrCode):
        if self.parent != None:
            self.parent.OnReceiveData(szTrCode)

    def OnReceiveChartRealData(self, szTrCode):
        if self.parent != None:
            self.parent.OnReceiveChartRealData(szTrCode)

    def OnReceiveSearchRealData(self, szTrCode):
        if self.parent != None:
            self.parent.OnReceiveSearchRealData(szTrCode)

class XAQuery(object):
    def __init__(self, parent=None, 식별자='식별자'):
        self.parent = parent
        self.식별자 = 식별자

        self.ActiveX = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        self.ActiveX.set_parent(parent=self)

        pathname = os.path.dirname(sys.argv[0])
        self.RESDIR = os.path.abspath(pathname)

        self.MYNAME = self.__class__.__name__
        self.INBLOCK = "%sInBlock" % self.MYNAME
        self.INBLOCK1 = "%sInBlock1" % self.MYNAME
        self.OUTBLOCK = "%sOutBlock" % self.MYNAME
        self.OUTBLOCK1 = "%sOutBlock1" % self.MYNAME
        self.OUTBLOCK2 = "%sOutBlock2" % self.MYNAME
        self.OUTBLOCK3 = "%sOutBlock3" % self.MYNAME
        self.RESFILE = "%s\\Res\\%s.res" % (self.RESDIR, self.MYNAME)

    def toint(self, s):
        temp = s.strip()
        result = 0

        if temp not in ['-']:
            result = int(temp)
        else:
            result = 0

        return result

    def tofloat(self, s):
        temp = s.strip()
        result = 0

        if temp not in ['-']:
            result = float(temp)
        else:
            result = 0.0

        return result

    def OnReceiveMessage(self, systemError, messageCode, message):
        if self.parent != None:
            self.parent.OnReceiveMessage(systemError, messageCode, message)

    def OnReceiveData(self, szTrCode):
        pass

    def OnReceiveChartRealData(self, szTrCode):
        pass

    def RequestLinkToHTS(self, szLinkName, szData, szFiller):
        return self.ActiveX.RequestLinkToHTS(szLinkName, szData, szFiller)


##----------------------------------------------------------------------------------------------------------------------
# 주식

#주식현재가(시세)조회
class t1102(XAQuery):
    def Query(self, 종목코드):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 종목코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = dict()

        result['한글명'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "hname", 0)
        result['현재가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "price", 0)
        result['전일대비구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", 0)
        result['전일대비'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "change", 0)
        result['등락율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", 0)
        result['누적거래량'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", 0)
        result['기준가_평가가격'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "recprice", 0)
        result['가중평균'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "avg", 0)
        result['상한가_최고호가가격'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "uplmtprice", 0)
        result['하한가_최저호가가격'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dnlmtprice", 0)
        result['전일거래량'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilvolume", 0)
        result['거래량차'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "volumediff", 0)
        result['시가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "open", 0)
        result['시가시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "opentime", 0)
        result['고가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "high", 0)
        result['고가시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "hightime", 0)
        result['저가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "low", 0)
        result['저가시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "lowtime", 0)
        result['최고가_52'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "high52w", 0)
        result['최고가일_52'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "high52wdate", 0)
        result['최저가_52'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "low52w", 0)
        result['최저가일_52'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "low52wdate", 0)
        result['소진율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "exhratio", 0)
        result['PER'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "per", 0)
        result['PBRX'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "pbrx", 0)
        result['상장주식수_천'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "listing", 0)
        result['증거금율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "jkrate", 0)
        result['수량단위'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "memedan", 0)
        result['매도증권사코드1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offernocd1", 0)
        result['매수증권사코드1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidnocd1", 0)
        result['매도증권사명1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offerno1", 0)
        result['매수증권사명1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidno1", 0)
        result['총매도수량1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dvol1", 0)
        result['총매수수량1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svol1", 0)
        result['매도증감1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dcha1", 0)
        result['매수증감1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "scha1", 0)
        result['매도비율1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ddiff1", 0)
        result['매수비율1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sdiff1", 0)
        result['매도증권사코드2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offernocd2", 0)
        result['매수증권사코드2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidnocd2", 0)
        result['매도증권사명2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offerno2", 0)
        result['매수증권사명2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidno2", 0)
        result['총매도수량2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dvol2", 0)
        result['총매수수량2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svol2", 0)
        result['매도증감2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dcha2", 0)
        result['매수증감2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "scha2", 0)
        result['매도비율2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ddiff2", 0)
        result['매수비율2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sdiff2", 0)
        result['매도증권사코드3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offernocd3", 0)
        result['매수증권사코드3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidnocd3", 0)
        result['매도증권사명3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offerno3", 0)
        result['매수증권사명3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidno3", 0)
        result['총매도수량3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dvol3", 0)
        result['총매수수량3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svol3", 0)
        result['매도증감3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dcha3", 0)
        result['매수증감3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "scha3", 0)
        result['매도비율3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ddiff3", 0)
        result['매수비율3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sdiff3", 0)
        result['매도증권사코드4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offernocd4", 0)
        result['매수증권사코드4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidnocd4", 0)
        result['매도증권사명4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offerno4", 0)
        result['매수증권사명4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidno4", 0)
        result['총매도수량4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dvol4", 0)
        result['총매수수량4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svol4", 0)
        result['매도증감4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dcha4", 0)
        result['매수증감4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "scha4", 0)
        result['매도비율4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ddiff4", 0)
        result['매수비율4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sdiff4", 0)
        result['매도증권사코드5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offernocd5", 0)
        result['매수증권사코드5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidnocd5", 0)
        result['매도증권사명5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "offerno5", 0)
        result['매수증권사명5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bidno5", 0)
        result['총매도수량5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dvol5", 0)
        result['총매수수량5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svol5", 0)
        result['매도증감5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dcha5", 0)
        result['매수증감5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "scha5", 0)
        result['매도비율5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ddiff5", 0)
        result['매수비율5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sdiff5", 0)
        result['외국계매도합계수량'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "fwdvl", 0)
        result['외국계매도직전대비'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmdcha", 0)
        result['외국계매도비율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmddiff", 0)
        result['외국계매수합계수량'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "fwsvl", 0)
        result['외국계매수직전대비'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmscha", 0)
        result['외국계매수비율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmsdiff", 0)
        result['회전율'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "vol", 0)
        result['단축코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0)
        result['누적거래대금'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "value", 0)
        result['전일동시간거래량'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "jvolume", 0)
        result['연중최고가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "highyear", 0)
        result['연중최고일자'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "highyeardate", 0)
        result['연중최저가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "lowyear", 0)
        result['연중최저일자'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "lowyeardate", 0)
        result['목표가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "target", 0)
        result['자본금'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "capital", 0)
        result['유동주식수'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "abscnt", 0)
        result['액면가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "parprice", 0)
        result['결산월'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "gsmm", 0)
        result['대용가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "subprice", 0)
        result['시가총액'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "total", 0)
        result['상장일'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "listdate", 0)
        result['전분기명'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "name", 0)
        result['전분기매출액'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfsales", 0)
        result['전분기영업이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfoperatingincome", 0)
        result['전분기경상이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfordinaryincome", 0)
        result['전분기순이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfnetincome", 0)
        result['전분기EPS'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfeps", 0)
        result['전전분기명'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "name2", 0)
        result['전전분기매출액'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfsales2", 0)
        result['전전분기영업이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfoperatingincome2", 0)
        result['전전분기경상이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfordinaryincome2", 0)
        result['전전분기순이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfnetincome2", 0)
        result['전전분기EPS'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "bfeps2", 0)
        result['전년대비매출액'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "salert", 0)
        result['전년대비영업이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "opert", 0)
        result['전년대비경상이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ordrt", 0)
        result['전년대비순이익'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "netrt", 0)
        result['전년대비EPS'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "epsrt", 0)
        result['락구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "info1", 0)
        result['관리_급등구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "info2", 0)
        result['정지_연장구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "info3", 0)
        result['투자_불성실구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "info4", 0)
        result['장구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "janginfo", 0)
        result['TPER'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "t_per", 0)
        result['통화ISO코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "tonghwa", 0)
        result['총매도대금1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dval1", 0)
        result['총매수대금1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sval1", 0)
        result['총매도대금2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dval2", 0)
        result['총매수대금2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sval2", 0)
        result['총매도대금3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dval3", 0)
        result['총매수대금3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sval3", 0)
        result['총매도대금4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dval4", 0)
        result['총매수대금4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sval4", 0)
        result['총매도대금5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dval5", 0)
        result['총매수대금5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "sval5", 0)
        result['총매도평단가1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "davg1", 0)
        result['총매수평단가1'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "savg1", 0)
        result['총매도평단가2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "davg2", 0)
        result['총매수평단가2'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "savg2", 0)
        result['총매도평단가3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "davg3", 0)
        result['총매수평단가3'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "savg3", 0)
        result['총매도평단가4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "davg4", 0)
        result['총매수평단가4'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "savg4", 0)
        result['총매도평단가5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "davg5", 0)
        result['총매수평단가5'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "savg5", 0)
        result['외국계매도대금'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmdval", 0)
        result['외국계매수대금'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmsval", 0)
        result['외국계매도평단가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmdvag", 0)
        result['외국계매수평단가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "ftradmsvag", 0)
        result['투자주의환기'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "info5", 0)
        result['기업인수목적회사여부'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "spac_gubun", 0)
        result['발행가격'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "issueprice", 0)
        result['배분적용구분코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "alloc_gubun", 0)
        result['배분적용구분'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "alloc_text", 0)
        result['단기과열_VI발동'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "shterm_text", 0)
        result['정적VI상한가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svi_uplmtprice", 0)
        result['정적VI하한가'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "svi_dnlmtprice", 0)
        result['저유동성종목여부'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "low_lqdt_gu", 0)
        result['이상급등종목여부'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "abnormal_rise_gu", 0)
        result['대차불가표시'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "lend_text", 0)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [result])


# 현물정상주문
class CSPAT00600(XAQuery):
    def Query(self, 계좌번호='',입력비밀번호='',종목번호='',주문수량='',주문가='',매매구분='2',호가유형코드='00',신용거래코드='000',대출일='',주문조건구분='0'):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        if 호가유형코드 == '03':
            주문가=''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "InptPwd", 0, 입력비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuNo", 0, 종목번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdPrc", 0, 주문가)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdprcPtnCode", 0, 호가유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "MgntrnCode", 0, 신용거래코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "LoanDt", 0, 대출일)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdCndiTpCode", 0, 주문조건구분)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            입력비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "InptPwd", i).strip()
            종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuNo", i).strip()
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            주문가 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdPrc", i).strip()
            매매구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", i).strip()
            호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdprcPtnCode", i).strip()
            프로그램호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "PrgmOrdprcPtnCode", i).strip()
            공매도가능여부 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "StslAbleYn", i).strip()
            공매도호가구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "StslOrdprcTpCode", i).strip()
            통신매체코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CommdaCode", i).strip()
            신용거래코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "MgntrnCode", i).strip()
            대출일 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "LoanDt", i).strip()
            회원번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "MbrNo", i).strip()
            주문조건구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdCndiTpCode", i).strip()
            전략코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "StrtgCode", i).strip()
            그룹ID = self.ActiveX.GetFieldData(self.OUTBLOCK1, "GrpId", i).strip()
            주문회차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdSeqNo", i).strip())
            포트폴리오번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "PtflNo", i).strip())
            바스켓번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "BskNo", i).strip())
            트렌치번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrchNo", i).strip())
            아이템번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "ItemNo", i).strip())
            운용지시번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OpDrtnNo", i).strip()
            유동성공급자여부 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "LpYn", i).strip()
            반대매매구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CvrgTpCode", i).strip()

            lst = [레코드갯수, 계좌번호, 입력비밀번호, 종목번호, 주문수량, 주문가, 매매구분, 호가유형코드, 프로그램호가유형코드, 공매도가능여부, 공매도호가구분, 통신매체코드, 신용거래코드,
                   대출일, 회원번호, 주문조건구분, 전략코드, 그룹ID, 주문회차, 포트폴리오번호, 바스켓번호, 트렌치번호, 아이템번호, 운용지시번호, 유동성공급자여부, 반대매매구분]
            result.append(lst)

        columns = ['레코드갯수', '계좌번호', '입력비밀번호', '종목번호', '주문수량', '주문가', '매매구분', '호가유형코드', '프로그램호가유형코드',
                   '공매도가능여부', '공매도호가구분', '통신매체코드', '신용거래코드', '대출일', '회원번호', '주문조건구분', '전략코드', '그룹ID',
                   '주문회차', '포트폴리오번호', '바스켓번호', '트렌치번호', '아이템번호', '운용지시번호', '유동성공급자여부', '반대매매구분']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdNo", i).strip())
            주문시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdTime", i).strip()
            주문시장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdMktCode", i).strip()
            주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnCode", i).strip()
            단축종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "ShtnIsuNo", i).strip()
            관리사원번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgempNo", i).strip()
            주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAmt", i).strip())
            예비주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SpareOrdNo", i).strip())
            반대매매일련번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgSeqno", i).strip())
            예약주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RsvOrdNo", i).strip())
            실물주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SpotOrdQty", i).strip())
            재사용주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RuseOrdQty", i).strip())
            현금주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MnyOrdAmt", i).strip())
            대용주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SubstOrdAmt", i).strip())
            재사용주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RuseOrdAmt", i).strip())
            계좌명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNm", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()

            lst = [레코드갯수, 주문번호, 주문시각, 주문시장코드, 주문유형코드, 단축종목번호, 관리사원번호, 주문금액, 예비주문번호, 반대매매일련번호, 예약주문번호, 실물주문수량, 재사용주문수량,
                   현금주문금액, 대용주문금액, 재사용주문금액, 계좌명, 종목명]
            result.append(lst)

        columns = ['레코드갯수', '주문번호', '주문시각', '주문시장코드', '주문유형코드', '단축종목번호', '관리사원번호', '주문금액', '예비주문번호',
                   '반대매매일련번호', '예약주문번호', '실물주문수량', '재사용주문수량', '현금주문금액', '대용주문금액', '재사용주문금액', '계좌명',
                   '종목명']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 현물정정주문
class CSPAT00700(XAQuery):
    def Query(self, 원주문번호,계좌번호,입력비밀번호,종목번호,주문수량,호가유형코드,주문조건구분,주문가):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrgOrdNo", 0, 원주문번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "InptPwd", 0, 입력비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuNo", 0, 종목번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdprcPtnCode", 0, 호가유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdCndiTpCode", 0, 주문조건구분)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdPrc", 0, 주문가)
        self.ActiveX.Request(0)

    def OnReceiveMessage(self, systemError, messageCode, message):
        클래스이름 = self.__class__.__name__
        함수이름 = inspect.currentframe().f_code.co_name
        print("%s-%s " % (클래스이름, 함수이름), systemError, messageCode, message)
        self.주문결과코드 = messageCode
        self.주문결과메세지 = message

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            원주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrgOrdNo", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            입력비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "InptPwd", i).strip()
            종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuNo", i).strip()
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdprcPtnCode", i).strip()
            주문조건구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdCndiTpCode", i).strip()
            주문가 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdPrc", i).strip()
            통신매체코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CommdaCode", i).strip()
            전략코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "StrtgCode", i).strip()
            그룹ID = self.ActiveX.GetFieldData(self.OUTBLOCK1, "GrpId", i).strip()
            주문회차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdSeqNo", i).strip())
            포트폴리오번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "PtflNo", i).strip())
            바스켓번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "BskNo", i).strip())
            트렌치번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrchNo", i).strip())
            아이템번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "ItemNo", i).strip())

            lst = [레코드갯수, 원주문번호, 계좌번호, 입력비밀번호, 종목번호, 주문수량, 호가유형코드, 주문조건구분, 주문가, 통신매체코드, 전략코드, 그룹ID, 주문회차, 포트폴리오번호,
                   바스켓번호, 트렌치번호, 아이템번호]
            result.append(lst)

        columns = ['레코드갯수', '원주문번호', '계좌번호', '입력비밀번호', '종목번호', '주문수량', '호가유형코드', '주문조건구분', '주문가', '통신매체코드', '전략코드',
                   '그룹ID', '주문회차', '포트폴리오번호', '바스켓번호', '트렌치번호', '아이템번호']
        df = DataFrame(data=result, columns=columns)


        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdNo", i).strip())
            모주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrntOrdNo", i).strip())
            주문시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdTime", i).strip()
            주문시장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdMktCode", i).strip()
            주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnCode", i).strip()
            단축종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "ShtnIsuNo", i).strip()
            프로그램호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrgmOrdprcPtnCode", i).strip()
            공매도호가구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "StslOrdprcTpCode", i).strip()
            공매도가능여부 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "StslAbleYn", i).strip()
            신용거래코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgntrnCode", i).strip()
            대출일 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "LoanDt", i).strip()
            반대매매주문구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgOrdTp", i).strip()
            유동성공급자여부 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "LpYn", i).strip()
            관리사원번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgempNo", i).strip()
            주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAmt", i).strip())
            매매구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpCode", i).strip()
            예비주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SpareOrdNo", i).strip())
            반대매매일련번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgSeqno", i).strip())
            예약주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RsvOrdNo", i).strip())
            현금주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MnyOrdAmt", i).strip())
            대용주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SubstOrdAmt", i).strip())
            재사용주문금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RuseOrdAmt", i).strip())
            계좌명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNm", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()

            lst = [레코드갯수, 주문번호, 모주문번호, 주문시각, 주문시장코드, 주문유형코드, 단축종목번호, 프로그램호가유형코드, 공매도호가구분, 공매도가능여부, 신용거래코드, 대출일,
                   반대매매주문구분, 유동성공급자여부, 관리사원번호, 주문금액, 매매구분, 예비주문번호, 반대매매일련번호, 예약주문번호, 현금주문금액, 대용주문금액, 재사용주문금액, 계좌명, 종목명]
            result.append(lst)

        columns = ['레코드갯수', '주문번호', '모주문번호', '주문시각', '주문시장코드', '주문유형코드', '단축종목번호', '프로그램호가유형코드', '공매도호가구분', '공매도가능여부',
                   '신용거래코드', '대출일', '반대매매주문구분', '유동성공급자여부', '관리사원번호', '주문금액', '매매구분', '예비주문번호', '반대매매일련번호', '예약주문번호',
                   '현금주문금액', '대용주문금액', '재사용주문금액', '계좌명', '종목명']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 현물취소주문
class CSPAT00800(XAQuery):
    def Query(self, 원주문번호,계좌번호,입력비밀번호,종목번호,주문수량):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrgOrdNo", 0, 원주문번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "InptPwd", 0, 입력비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuNo", 0, 종목번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.Request(0)

    def OnReceiveMessage(self, systemError, messageCode, message):
        클래스이름 = self.__class__.__name__
        함수이름 = inspect.currentframe().f_code.co_name
        print("%s-%s " % (클래스이름, 함수이름), systemError, messageCode, message)
        self.주문결과코드 = messageCode
        self.주문결과메세지 = message

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            원주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrgOrdNo", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            입력비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "InptPwd", i).strip()
            종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuNo", i).strip()
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            통신매체코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CommdaCode", i).strip()
            그룹ID = self.ActiveX.GetFieldData(self.OUTBLOCK1, "GrpId", i).strip()
            전략코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "StrtgCode", i).strip()
            주문회차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdSeqNo", i).strip())
            포트폴리오번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "PtflNo", i).strip())
            바스켓번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "BskNo", i).strip())
            트렌치번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrchNo", i).strip())
            아이템번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "ItemNo", i).strip())

            lst = [레코드갯수, 원주문번호, 계좌번호, 입력비밀번호, 종목번호, 주문수량, 통신매체코드, 그룹ID, 전략코드, 주문회차, 포트폴리오번호, 바스켓번호, 트렌치번호, 아이템번호]
            result.append(lst)

        columns = ['레코드갯수', '원주문번호', '계좌번호', '입력비밀번호', '종목번호', '주문수량', '통신매체코드', '그룹ID', '전략코드', '주문회차', '포트폴리오번호',
                   '바스켓번호', '트렌치번호', '아이템번호']
        df = DataFrame(data=result, columns=columns)


        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdNo", i).strip())
            모주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrntOrdNo", i).strip())
            주문시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdTime", i).strip()
            주문시장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdMktCode", i).strip()
            주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnCode", i).strip()
            단축종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "ShtnIsuNo", i).strip()
            프로그램호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrgmOrdprcPtnCode", i).strip()
            공매도호가구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "StslOrdprcTpCode", i).strip()
            공매도가능여부 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "StslAbleYn", i).strip()
            신용거래코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgntrnCode", i).strip()
            대출일 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "LoanDt", i).strip()
            반대매매주문구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgOrdTp", i).strip()
            유동성공급자여부 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "LpYn", i).strip()
            관리사원번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgempNo", i).strip()
            매매구분 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpCode", i).strip()
            예비주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "SpareOrdNo", i).strip())
            반대매매일련번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgSeqno", i).strip())
            예약주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RsvOrdNo", i).strip())
            계좌명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNm", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()

            lst = [레코드갯수, 주문번호, 모주문번호, 주문시각, 주문시장코드, 주문유형코드, 단축종목번호, 프로그램호가유형코드, 공매도호가구분, 공매도가능여부, 신용거래코드, 대출일,
                   반대매매주문구분, 유동성공급자여부, 관리사원번호, 매매구분, 예비주문번호, 반대매매일련번호, 예약주문번호, 계좌명, 종목명]
            result.append(lst)

        columns = ['레코드갯수', '주문번호', '모주문번호', '주문시각', '주문시장코드', '주문유형코드', '단축종목번호', '프로그램호가유형코드', '공매도호가구분', '공매도가능여부',
                   '신용거래코드', '대출일', '반대매매주문구분', '유동성공급자여부', '관리사원번호', '매매구분', '예비주문번호', '반대매매일련번호', '예약주문번호', '계좌명',
                   '종목명']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 주식잔고2
class t0424(XAQuery):
    def Query(self, 계좌번호='', 비밀번호='', 단가구분='1', 체결구분='0', 단일가구분='0', 제비용포함여부='1', CTS_종목번호=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "accno", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK, "passwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK, "prcgb", 0, 단가구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "chegb", 0, 체결구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "dangb", 0, 단일가구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "charge", 0, 제비용포함여부)
        self.ActiveX.SetFieldData(self.INBLOCK, "cts_expcode", 0, CTS_종목번호)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            추정순자산 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "sunamt", i).strip())
            실현손익 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "dtsunik", i).strip())
            매입금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "mamt", i).strip())
            추정D2예수금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "sunamt1", i).strip())
            CTS_종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_expcode", i).strip()
            평가금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tappamt", i).strip())
            평가손익 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tdtsunik", i).strip())
    
            lst = [추정순자산, 실현손익, 매입금액, 추정D2예수금, CTS_종목번호, 평가금액, 평가손익]
            result.append(lst)
    
        columns = ['추정순자산', '실현손익', '매입금액', '추정D2예수금', 'CTS_종목번호', '평가금액', '평가손익']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "expcode", i).strip()
            잔고구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "jangb", i).strip()
            잔고수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "janqty", i).strip())
            매도가능수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdposqt", i).strip())
            평균단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "pamt", i).strip())
            매입금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mamt", i).strip())
            대출금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sinamt", i).strip())
            만기일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "lastdt", i).strip()
            당일매수금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "msat", i).strip())
            당일매수단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mpms", i).strip())
            당일매도금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdat", i).strip())
            당일매도단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mpmd", i).strip())
            전일매수금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jsat", i).strip())
            전일매수단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jpms", i).strip())
            전일매도금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jdat", i).strip())
            전일매도단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jpmd", i).strip())
            처리순번 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sysprocseq", i).strip())
            대출일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "loandt", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            시장구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "marketgb", i).strip()
            종목구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "jonggb", i).strip()
            보유비중 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "janrt", i).strip())
            현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            평가금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "appamt", i).strip())
            평가손익 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "dtsunik", i).strip())
            수익율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sunikrt", i).strip())
            수수료 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "fee", i).strip())
            제세금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "tax", i).strip())
            신용이자 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sininter", i).strip())
    
            lst = [종목번호, 잔고구분, 잔고수량, 매도가능수량, 평균단가, 매입금액, 대출금액, 만기일자, 당일매수금액,
                   당일매수단가, 당일매도금액, 당일매도단가, 전일매수금액, 전일매수단가, 전일매도금액, 전일매도단가,
                   처리순번, 대출일자, 종목명, 시장구분, 종목구분, 보유비중, 현재가, 평가금액, 평가손익, 수익율, 수수료, 제세금, 신용이자]
            result.append(lst)
    
        columns = ['종목번호', '잔고구분', '잔고수량', '매도가능수량', '평균단가', '매입금액', '대출금액', '만기일자', '당일매수금액', ' 당일매수단가', '당일매도금액',
                   '당일매도단가', '전일매수금액', '전일매수단가', '전일매도금액', '전일매도단가', ' 처리순번', '대출일자', '종목명', '시장구분', '종목구분', '보유비중', '현재가',
                   '평가금액', '평가손익', '수익율', '수수료', '제세금', '신용이자']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 주식분별주가조회
class t1302(XAQuery):
    def Query(self, 단축코드='',작업구분='1',시간='',건수='900', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 작업구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "time", 0, 시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cnt", 0, 건수)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 시간)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            시간CTS = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", i).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "chetime", i).strip()
            종가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            체결강도 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "chdegree", i).strip())
            매도체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdvolume", i).strip())
            매수체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "msvolume", i).strip())
            순매수체결량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "revolume", i).strip())
            매도체결건수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdchecnt", i).strip())
            매수체결건수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mschecnt", i).strip())
            순체결건수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "rechecnt", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            시가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            체결량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "cvolume", i).strip())
            매도체결건수시간 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdchecnttm", i).strip())
            매수체결건수시간 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mschecnttm", i).strip())
            매도잔량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "totofferrem", i).strip())
            매수잔량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "totbidrem", i).strip())
            시간별매도체결량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "mdvolumetm", i).strip())
            시간별매수체결량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "msvolumetm", i).strip())

            lst = [시간, 종가, 전일대비구분, 전일대비, 등락율, 체결강도, 매도체결수량, 매수체결수량, 순매수체결량, 매도체결건수, 매수체결건수, 순체결건수, 거래량, 시가, 고가, 저가, 체결량,
                   매도체결건수시간, 매수체결건수시간, 매도잔량, 매수잔량, 시간별매도체결량, 시간별매수체결량]
            result.append(lst)

        columns = ['시간', '종가', '전일대비구분', '전일대비', '등락율', '체결강도', '매도체결수량', '매수체결수량', '순매수체결량', '매도체결건수', '매수체결건수',
                   '순체결건수', '거래량', '시가', '고가', '저가', '체결량', '매도체결건수시간', '매수체결건수시간', '매도잔량', '매수잔량', '시간별매도체결량',
                   '시간별매수체결량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [시간CTS, df])


# 기간별주가
class t1305(XAQuery):
    def Query(self, 단축코드='',일주월구분='1',날짜='',IDX='',건수='900', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "dwmcode", 0, 일주월구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "date", 0, 날짜)
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)
            self.ActiveX.SetFieldData(self.INBLOCK, "cnt", 0, 건수)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "date", 0, 날짜)
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)
            self.ActiveX.SetFieldData(self.INBLOCK, "cnt", 0, 건수)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            CNT = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "cnt", i).strip())
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK, "date", i).strip()
            IDX = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            거래증가율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff_vol", i).strip())
            체결강도 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "chdegree", i).strip())
            소진율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sojinrate", i).strip())
            회전율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "changerate", i).strip())
            외인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "fpvolume", i).strip())
            기관순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "covolume", i).strip())
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            누적거래대금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value", i).strip())
            개인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "ppvolume", i).strip())
            시가대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "o_sign", i).strip()
            시가대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "o_change", i).strip())
            시가기준등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "o_diff", i).strip())
            고가대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "h_sign", i).strip()
            고가대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "h_change", i).strip())
            고가기준등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "h_diff", i).strip())
            저가대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "l_sign", i).strip()
            저가대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "l_change", i).strip())
            저가기준등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "l_diff", i).strip())
            시가총액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "marketcap", i).strip())

            lst = [날짜, 시가, 고가, 저가, 종가, 전일대비구분, 전일대비, 등락율, 누적거래량, 거래증가율, 체결강도, 소진율, 회전율, 외인순매수, 기관순매수, 종목코드, 누적거래대금,
                   개인순매수, 시가대비구분, 시가대비, 시가기준등락율, 고가대비구분, 고가대비, 고가기준등락율, 저가대비구분, 저가대비, 저가기준등락율, 시가총액]
            result.append(lst)

        columns = ['날짜', '시가', '고가', '저가', '종가', '전일대비구분', '전일대비', '등락율', '누적거래량', '거래증가율', '체결강도', '소진율', '회전율',
                   '외인순매수', '기관순매수', '종목코드', '누적거래대금', '개인순매수', '시가대비구분', '시가대비', '시가기준등락율', '고가대비구분', '고가대비',
                   '고가기준등락율', '저가대비구분', '저가대비', '저가기준등락율', '시가총액']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [CNT, 날짜, IDX, df])


# 거래량상위
class t1452(XAQuery):
    def Query(self, 구분='0',전일구분='',시작등락율='',종료등락율='',대상제외='',시작가격='',종료가격='',거래량='',IDX='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "jnilgubun", 0, 전일구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdiff", 0, 시작등락율)
            self.ActiveX.SetFieldData(self.INBLOCK, "ediff", 0, 종료등락율)
            self.ActiveX.SetFieldData(self.INBLOCK, "jc_num", 0, 대상제외)
            self.ActiveX.SetFieldData(self.INBLOCK, "sprice", 0, 시작가격)
            self.ActiveX.SetFieldData(self.INBLOCK, "eprice", 0, 종료가격)
            self.ActiveX.SetFieldData(self.INBLOCK, "volume", 0, 거래량)
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            IDX = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            회전율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "vol", i).strip())
            전일거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jnilvolume", i).strip())
            전일비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "bef_diff", i).strip())
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()

            lst = [종목명,현재가,전일대비구분,전일대비,등락율,누적거래량,회전율,전일거래량,전일비,종목코드]
            result.append(lst)

        columns = ['종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','회전율','전일거래량','전일비','종목코드']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [IDX, df])


# 거래대금상위
class t1463(XAQuery):
    def Query(self, 구분='0',전일구분='',대상제외='',시작가격='',종료가격='',거래량='',IDX='',대상제외2='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "jnilgubun", 0, 전일구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "jc_num", 0, 대상제외)
            self.ActiveX.SetFieldData(self.INBLOCK, "sprice", 0, 시작가격)
            self.ActiveX.SetFieldData(self.INBLOCK, "eprice", 0, 종료가격)
            self.ActiveX.SetFieldData(self.INBLOCK, "volume", 0, 거래량)
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)
            self.ActiveX.SetFieldData(self.INBLOCK, "jc_num2", 0, 대상제외2)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "idx", 0, IDX)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            IDX = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            한글명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            거래대금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value", i).strip())
            전일거래대금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jnilvalue", i).strip())
            전일비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "bef_diff", i).strip())
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            filler = self.ActiveX.GetFieldData(self.OUTBLOCK1, "filler", i).strip()
            전일거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jnilvolume", i).strip())

            lst = [한글명, 현재가, 전일대비구분, 전일대비, 등락율, 누적거래량, 거래대금, 전일거래대금, 전일비, 종목코드, filler, 전일거래량]
            result.append(lst)

        columns = ['한글명', '현재가', '전일대비구분', '전일대비', '등락율', '누적거래량', '거래대금', '전일거래대금', '전일비', '종목코드', 'filler', '전일거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [IDX, df])


# 업종기간별추이
class t1514(XAQuery):
    def Query(self, 업종코드='001',구분1='',구분2='1',CTS일자='',조회건수='100',비중구분='', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "upcode", 0, 업종코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun1", 0, 구분1)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun2", 0, 구분2)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, CTS일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cnt", 0, 조회건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "rate_gbn", 0, 비중구분)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, CTS일자)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            CTS일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", i).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            지수 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jisu", i))
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i))
            등락율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i))
            거래량 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i))
            거래증가율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff_vol", i))
            거래대금1 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value1", i))
            상승 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i))
            보합 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "unchg", i))
            하락 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i))
            상승종목비율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "uprate", i))
            외인순매수 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "frgsvolume", i))
            시가 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "openjisu", i))
            고가 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "highjisu", i))
            저가 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "lowjisu", i))
            거래대금2 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value2", i))
            상한 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "up", i))
            하한 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "down", i))
            종목수 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "totjo", i))
            기관순매수 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "orgsvolume", i))
            업종코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "upcode", i).strip()
            거래비중 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "rate", i))
            업종배당수익률 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "divrate", i))

            lst = [일자, 지수, 전일대비구분, 전일대비, 등락율, 거래량, 거래증가율, 거래대금1, 상승, 보합, 하락, 상승종목비율,
                   외인순매수, 시가, 고가, 저가, 거래대금2, 상한, 하한, 종목수, 기관순매수, 업종코드, 거래비중, 업종배당수익률]

            result.append(lst)

        columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [CTS일자, df])


# 업종별 종목시세
# 업종별종목 리스트
class t1516(XAQuery):
    def Query(self, 업종코드='001',구분='',종목코드='', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "upcode", 0, 업종코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 종목코드)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 종목코드)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            지수 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK, "pricejisu", i).strip())
            지수_전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            지수_전일대비 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            지수_등락율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK, "jdiff", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            현재가 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i))
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i))
            등락율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i))
            누적거래량 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i))
            시가 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i))
            고가 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i))
            저가 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i))
            소진율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sojinrate", i))
            베타계수 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "beta", i))
            PER = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "perx", i))
            외인순매수 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "frgsvolume", i))
            기관순매수 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "orgsvolume", i))
            거래증가율 = self.tofloat(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff_vol", i))
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            시가총액 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "total", i))
            거래대금 = self.toint(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value", i))

            lst = [종목명,현재가,전일대비구분,전일대비,등락율,누적거래량,시가,고가,저가,소진율,베타계수,PER,외인순매수,기관순매수,거래증가율,종목코드,시가총액,거래대금]

            result.append(lst)

        columns = ['종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','시가','고가','저가','소진율','베타계수','PER','외인순매수','기관순매수','거래증가율','종목코드','시가총액','거래대금']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [지수,지수_전일대비구분,지수_전일대비,지수_등락율, df])


# 테마종목별 시세조회
class t1537(XAQuery):
    def Query(self, 테마코드='0001', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "tmcode", 0, 테마코드)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, CTS일자)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            상승종목수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "upcnt", i).strip())
            테마종목수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tmcnt", i).strip())
            상승종목비율 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "uprate", i).strip())
            테마명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "tmname", i).strip()

            lst = [상승종목수, 테마종목수, 상승종목비율, 테마명]
            result.append(lst)

        columns = ['상승종목수', '테마종목수', '상승종목비율', '테마명']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            전일동시간 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jniltime", i).strip())
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            예상체결가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "yeprice", i).strip())
            시가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            누적거래대금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value", i).strip())
            시가총액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "marketcap", i).strip())

            lst = [종목명, 현재가, 전일대비구분, 전일대비, 등락율, 누적거래량, 전일동시간, 종목코드, 예상체결가, 시가, 고가, 저가, 누적거래대금, 시가총액]
            result.append(lst)

        columns = ['종목명', '현재가', '전일대비구분', '전일대비', '등락율', '누적거래량', '전일동시간', '종목코드', '예상체결가', '시가', '고가', '저가',
                       '누적거래대금', '시가총액']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 시간대별투자자매매추이(t1602)
class t1602(XAQuery):
    def Query(self, 시장구분='1', 업종코드='001', 수량구분='2', 전일분구분='0', CTSTIME='',CTSIDX='',조회건수=100,직전대비구분=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "market", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "upcode", 0, 업종코드)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun1", 0, 수량구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun2", 0, 전일분구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, CTSTIME)
        self.ActiveX.SetFieldData(self.INBLOCK, "cts_idx", 0, CTSIDX)
        self.ActiveX.SetFieldData(self.INBLOCK, "cnt", 0, 조회건수)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun3", 0, 직전대비구분)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            CTSTIME = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", i).strip()
            개인투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "tjjcode_08", i).strip()
            개인매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_08", i).strip())
            개인매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_08", i).strip())
            개인증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_08", i).strip())
            개인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_08", i).strip())
            외국인투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_17", i).strip()
            외국인매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_17", i).strip())
            외국인매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_17", i).strip())
            외국인증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_17", i).strip())
            외국인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_17", i).strip())
            기관계투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_18", i).strip()
            기관계매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_18", i).strip())
            기관계매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_18", i).strip())
            기관계증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_18", i).strip())
            기관계순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_18", i).strip())
            증권투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_01", i).strip()
            증권매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_01", i).strip())
            증권매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_01", i).strip())
            증권증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_01", i).strip())
            증권순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_01", i).strip())
            투신투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_03", i).strip()
            투신매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_03", i).strip())
            투신매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_03", i).strip())
            투신증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_03", i).strip())
            투신순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_03", i).strip())
            은행투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_04", i).strip()
            은행매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_04", i).strip())
            은행매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_04", i).strip())
            은행증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_04", i).strip())
            은행순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_04", i).strip())
            보험투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_02", i).strip()
            보험매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_02", i).strip())
            보험매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_02", i).strip())
            보험증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_02", i).strip())
            보험순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_02", i).strip())
            종금투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_05", i).strip()
            종금매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_05", i).strip())
            종금매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_05", i).strip())
            종금증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_05", i).strip())
            종금순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_05", i).strip())
            기금투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_06", i).strip()
            기금매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_06", i).strip())
            기금매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_06", i).strip())
            기금증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_06", i).strip())
            기금순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_06", i).strip())
            기타투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_07", i).strip()
            기타매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_07", i).strip())
            기타매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_07", i).strip())
            기타증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_07", i).strip())
            기타순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_07", i).strip())
            국가투자자코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_11", i).strip()
            국가매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_11", i).strip())
            국가매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_11", i).strip())
            국가증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_11", i).strip())
            국가순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_11", i).strip())
            사모펀드코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "jjcode_00", i).strip()
            사모펀드매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "ms_00", i).strip())
            사모펀드매도 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "md_00", i).strip())
            사모펀드증감 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rate_00", i).strip())
            사모펀드순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "svolume_00", i).strip())

        block = {'CTSTIME' : CTSTIME,
            '개인투자자코드' : 개인투자자코드,
            '개인매수' : 개인매수,
            '개인매도' : 개인매도,
            '개인증감' : 개인증감,
            '개인순매수' : 개인순매수,
            '외국인투자자코드' : 외국인투자자코드,
            '외국인매수' : 외국인매수,
            '외국인매도' : 외국인매도,
            '외국인증감' : 외국인증감,
            '외국인순매수' : 외국인순매수,
            '기관계투자자코드' : 기관계투자자코드,
            '기관계매수' : 기관계매수,
            '기관계매도' : 기관계매도,
            '기관계증감' : 기관계증감,
            '기관계순매수' : 기관계순매수,
            '증권투자자코드' : 증권투자자코드,
            '증권매수' : 증권매수,
            '증권매도' : 증권매도,
            '증권증감' : 증권증감,
            '증권순매수' : 증권순매수,
            '투신투자자코드' : 투신투자자코드,
            '투신매수' : 투신매수,
            '투신매도' : 투신매도,
            '투신증감' : 투신증감,
            '투신순매수' : 투신순매수,
            '은행투자자코드' : 은행투자자코드,
            '은행매수' : 은행매수,
            '은행매도' : 은행매도,
            '은행증감' : 은행증감,
            '은행순매수' : 은행순매수,
            '보험투자자코드' : 보험투자자코드,
            '보험매수' : 보험매수,
            '보험매도' : 보험매도,
            '보험증감' : 보험증감,
            '보험순매수' : 보험순매수,
            '종금투자자코드' : 종금투자자코드,
            '종금매수' : 종금매수,
            '종금매도' : 종금매도,
            '종금증감' : 종금증감,
            '종금순매수' : 종금순매수,
            '기금투자자코드' : 기금투자자코드,
            '기금매수' : 기금매수,
            '기금매도' : 기금매도,
            '기금증감' : 기금증감,
            '기금순매수' : 기금순매수,
            '기타투자자코드' : 기타투자자코드,
            '기타매수' : 기타매수,
            '기타매도' : 기타매도,
            '기타증감' : 기타증감,
            '기타순매수' : 기타순매수,
            '국가투자자코드' : 국가투자자코드,
            '국가매수' : 국가매수,
            '국가매도' : 국가매도,
            '국가증감' : 국가증감,
            '국가순매수' : 국가순매수,
            '사모펀드코드' : 사모펀드코드,
            '사모펀드매수' : 사모펀드매수,
            '사모펀드매도' : 사모펀드매도,
            '사모펀드증감' : 사모펀드증감,
            '사모펀드순매수' : 사모펀드순매수
            }

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            개인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_08", i).strip())
            외국인순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_17", i).strip())
            기관계순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_18", i).strip())
            증권순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_01", i).strip())
            투신순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_03", i).strip())
            은행순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_04", i).strip())
            보험순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_02", i).strip())
            종금순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_05", i).strip())
            기금순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_06", i).strip())
            기타순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_07", i).strip())
            국가순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_11", i).strip())
            사모펀드순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "sv_00", i).strip())

            lst = [시간,개인순매수,외국인순매수,기관계순매수,증권순매수,투신순매수,은행순매수,보험순매수,종금순매수,기금순매수,기타순매수,국가순매수,사모펀드순매수]

            result.append(lst)

        columns = ['시간','개인순매수','외국인순매수','기관계순매수','증권순매수','투신순매수','은행순매수','보험순매수','종금순매수','기금순매수','기타순매수','국가순매수','사모펀드순매수']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [block, df])


# 외인기관종목별동향
class t1702(XAQuery):
    def Query(self, 종목코드='069500',종료일자='',금액수량구분='0',매수매도구분='0',누적구분='0',CTSDATE='',CTSIDX=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 종목코드)
        self.ActiveX.SetFieldData(self.INBLOCK, "todt", 0, 종료일자)
        self.ActiveX.SetFieldData(self.INBLOCK, "volvalgb", 0, 금액수량구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "msmdgb", 0, 매수매도구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "cumulgb", 0, 누적구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, CTSDATE)
        self.ActiveX.SetFieldData(self.INBLOCK, "cts_idx", 0, CTSIDX)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            CTSIDX = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_idx", i).strip())
            CTSDATE = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", i).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            try:
                일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            except Exception as e:
                일자 = ''
            try:
                종가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            except Exception as e:
                종가 = 0
            try:
                전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            except Exception as e:
                전일대비구분 = 0
            try:
                전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            except Exception as e:
                전일대비 = 0
            try:
                등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            except Exception as e:
                등락율 = 0
            try:
                누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            except Exception as e:
                누적거래량 = 0
            try:
                사모펀드 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0000", i).strip())
            except Exception as e:
                사모펀드 = 0
            try:
                증권 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0001", i).strip())
            except Exception as e:
                증권 = 0
            try:
                보험 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0002", i).strip())
            except Exception as e:
                보험 = 0
            try:
                투신 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0003", i).strip())
            except Exception as e:
                투신 = 0
            try:
                은행 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0004", i).strip())
            except Exception as e:
                은행 = 0
            try:
                종금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0005", i).strip())
            except Exception as e:
                종금 = 0
            try:
                기금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0006", i).strip())
            except Exception as e:
                기금 = 0
            try:
                기타법인 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0007", i).strip())
            except Exception as e:
                기타법인 = 0
            try:
                개인 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0008", i).strip())
            except Exception as e:
                개인 = 0
            try:
                등록외국인 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0009", i).strip())
            except Exception as e:
                등록외국인 = 0
            try:
                미등록외국인 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0010", i).strip())
            except Exception as e:
                미등록외국인 = 0
            try:
                국가외 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0011", i).strip())
            except Exception as e:
                국가외 = 0
            try:
                기관 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0018", i).strip())
            except Exception as e:
                기관 = 0
            try:
                외인계 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0088", i).strip())
            except Exception as e:
                외인계 = 0
            try:
                기타계 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "amt0099", i).strip())
            except Exception as e:
                기타계 = 0

            lst = [일자, 종가, 전일대비구분, 전일대비, 등락율, 누적거래량, 사모펀드, 증권, 보험, 투신, 은행, 종금, 기금, 기타법인, 개인, 등록외국인, 미등록외국인, 국가외, 기관,
                   외인계, 기타계]

            result.append(lst)

        columns = ['일자', '종가', '전일대비구분', '전일대비', '등락율', '누적거래량', '사모펀드', '증권', '보험', '투신', '은행', '종금', '기금', '기타법인',
                       '개인', '등록외국인', '미등록외국인', '국가외', '기관', '외인계', '기타계']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [CTSIDX, CTSDATE, df])


# 외인기관종목별동향
class t1717(XAQuery):
    def Query(self, 종목코드='069500',구분='0',시작일자='20170101',종료일자='20172131'):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 종목코드)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "fromdt", 0, 시작일자)
        self.ActiveX.SetFieldData(self.INBLOCK, "todt", 0, 종료일자)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            try:
                일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "date", i).strip()
            except Exception as e:
                일자 = ''
            try:
                종가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "close", i).strip())
            except Exception as e:
                종가 = 0
            try:
                전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            except Exception as e:
                전일대비구분 = 0
            try:
                전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            except Exception as e:
                전일대비 = 0
            try:
                등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", i).strip())
            except Exception as e:
                등락율 = 0
            try:
                누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", i).strip())
            except Exception as e:
                누적거래량 = 0
            try:
                사모펀드_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0000_vol", i).strip())
            except Exception as e:
                사모펀드_순매수 = 0
            try:
                증권_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0001_vol", i).strip())
            except Exception as e:
                증권_순매수 = 0
            try:
                보험_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0002_vol", i).strip())
            except Exception as e:
                보험_순매수 = 0
            try:
                투신_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0003_vol", i).strip())
            except Exception as e:
                투신_순매수 = 0
            try:
                은행_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0004_vol", i).strip())
            except Exception as e:
                은행_순매수 = 0
            try:
                종금_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0005_vol", i).strip())
            except Exception as e:
                종금_순매수 = 0
            try:
                기금_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0006_vol", i).strip())
            except Exception as e:
                기금_순매수 = 0
            try:
                기타법인_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0007_vol", i).strip())
            except Exception as e:
                기타법인_순매수 = 0
            try:
                개인_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0008_vol", i).strip())
            except Exception as e:
                개인_순매수 = 0
            try:
                등록외국인_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0009_vol", i).strip())
            except Exception as e:
                등록외국인_순매수 = 0
            try:
                미등록외국인_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0010_vol", i).strip())
            except Exception as e:
                미등록외국인_순매수 = 0
            try:
                국가외_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0011_vol", i).strip())
            except Exception as e:
                국가외_순매수 = 0
            try:
                기관_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0018_vol", i).strip())
            except Exception as e:
                기관_순매수 = 0
            try:
                외인계_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0016_vol", i).strip())
            except Exception as e:
                외인계_순매수 = 0
            try:
                기타계_순매수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0017_vol", i).strip())
            except Exception as e:
                기타계_순매수 = 0
            try:
                사모펀드_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0000_dan", i).strip())
            except Exception as e:
                사모펀드_단가 = 0
            try:
                증권_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0001_dan", i).strip())
            except Exception as e:
                증권_단가 = 0
            try:
                보험_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0002_dan", i).strip())
            except Exception as e:
                보험_단가 = 0
            try:
                투신_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0003_dan", i).strip())
            except Exception as e:
                투신_단가 = 0
            try:
                은행_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0004_dan", i).strip())
            except Exception as e:
                은행_단가 = 0
            try:
                종금_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0005_dan", i).strip())
            except Exception as e:
                종금_단가 = 0
            try:
                기금_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0006_dan", i).strip())
            except Exception as e:
                기금_단가 = 0
            try:
                기타법인_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0007_dan", i).strip())
            except Exception as e:
                기타법인_단가 = 0
            try:
                개인_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0008_dan", i).strip())
            except Exception as e:
                개인_단가 = 0
            try:
                등록외국인_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0009_dan", i).strip())
            except Exception as e:
                등록외국인_단가 = 0
            try:
                미등록외국인_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0010_dan", i).strip())
            except Exception as e:
                미등록외국인_단가 = 0
            try:
                국가외_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0011_dan", i).strip())
            except Exception as e:
                국가외_단가 = 0
            try:
                기관_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0018_dan", i).strip())
            except Exception as e:
                기관_단가 = 0
            try:
                외인계_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0016_dan", i).strip())
            except Exception as e:
                외인계_단가 = 0
            try:
                기타계_단가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "tjj0017_dan", i).strip())
            except Exception as e:
                기타계_단가 = 0

            lst = [일자, 종가, 전일대비구분, 전일대비, 등락율, 누적거래량,
                   사모펀드_순매수, 증권_순매수, 보험_순매수, 투신_순매수, 은행_순매수, 종금_순매수, 기금_순매수, 기타법인_순매수, 개인_순매수, 등록외국인_순매수, 미등록외국인_순매수,
                   국가외_순매수, 기관_순매수, 외인계_순매수, 기타계_순매수,
                   사모펀드_단가, 증권_단가, 보험_단가, 투신_단가, 은행_단가, 종금_단가, 기금_단가, 기타법인_단가, 개인_단가, 등록외국인_단가, 미등록외국인_단가, 국가외_단가,
                   기관_단가, 외인계_단가, 기타계_단가]

            result.append(lst)

        columns = ['일자', '종가', '전일대비구분', '전일대비', '등락율', '누적거래량', '사모펀드_순매수', '증권_순매수', '보험_순매수', '투신_순매수', '은행_순매수',
                       '종금_순매수', '기금_순매수', '기타법인_순매수', '개인_순매수', '등록외국인_순매수', '미등록외국인_순매수', '국가외_순매수', '기관_순매수',
                       '외인계_순매수', '기타계_순매수', '사모펀드_단가', '증권_단가', '보험_단가', '투신_단가', '은행_단가', '종금_단가', '기금_단가', '기타법인_단가',
                       '개인_단가', '등록외국인_단가', '미등록외국인_단가', '국가외_단가', '기관_단가', '외인계_단가', '기타계_단가']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 주식챠트(틱/n틱)
class t8411(XAQuery):
    def Query(self, 단축코드,단위='1',요청건수='2000',조회영업일수='0',시작일자='',시작시간='',종료일자='',종료시간='',연속일자='',연속시간='',압축여부='Y', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "nday", 0, 조회영업일수)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "stime", 0, 시작시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "etime", 0, 종료시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "comp_yn", 0, 압축여부)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        block = dict()
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            block['단축코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            block['전일시가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jisiga", i).strip())
            block['전일고가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jihigh", i).strip())
            block['전일저가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jilow", i).strip())
            block['전일종가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jiclose", i).strip())
            block['전일거래량'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jivolume", i).strip())
            block['당일시가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "disiga", i).strip())
            block['당일고가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "dihigh", i).strip())
            block['당일저가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "dilow", i).strip())
            block['당일종가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "diclose", i).strip())
            block['상한가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "highend", i).strip())
            block['하한가'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "lowend", i).strip())
            block['연속일자'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", i).strip()
            block['연속시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", i).strip()
            block['장시작시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "s_time", i).strip()
            block['장종료시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "e_time", i).strip()
            block['동시호가처리시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dshmin", i).strip()
            block['레코드카운트'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jdiff_vol", i).strip())
            수정구분 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jongchk", i).strip())
            수정비율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "rate", i).strip())
            수정주가반영항목 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "pricechk", i).strip())

            lst = [날짜,시간,시가,고가,저가,종가,거래량,수정구분,수정비율,수정주가반영항목]

            result.append(lst)

        columns = ['날짜','시간','시가','고가','저가','종가','거래량','수정구분','수정비율','수정주가반영항목']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [block, df])


# 업종전체조회
class t8424(XAQuery):
    def Query(self, 구분=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun1", 0, 구분)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            업종명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "hname", i).strip()
            업종코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "upcode", i).strip()

            lst = [업종명, 업종코드]
            result.append(lst)

        columns = ['업종명', '업종코드']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 전체테마
class t8425(XAQuery):
    def Query(self):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "dummy", 0, "0")
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            테마명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "tmname", i).strip()
            테마코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "tmcode", i).strip()

            lst = [테마명, 테마코드]
            result.append(lst)

        columns = ['테마명', '테마코드']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 주식종목코드조회
class t8430(XAQuery):
    def Query(self, 구분='0'):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "hname", i).strip()
            단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            확장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "expcode", i).strip()
            ETF구분 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "etfgubun", i).strip())
            상한가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "uplmtprice", i).strip())
            하한가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "dnlmtprice", i).strip())
            전일가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            주문수량단위 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "memedan", i).strip())
            기준가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "recprice", i).strip())
            구분 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "gubun", i).strip())

            lst = [종목명, 단축코드, 확장코드, ETF구분, 상한가, 하한가, 전일가, 주문수량단위, 기준가, 구분]
            result.append(lst)

        columns = ['종목명', '단축코드', '확장코드', 'ETF구분', '상한가', '하한가', '전일가', '주문수량단위', '기준가', '구분']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 주식종목코드조회(API용)
class t8436(XAQuery):
    def Query(self, 구분='0'):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 구분)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "hname", i).strip()
            단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            확장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "expcode", i).strip()
            ETF구분 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "etfgubun", i).strip())
            상한가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "uplmtprice", i).strip())
            하한가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "dnlmtprice", i).strip())
            전일가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            주문수량단위 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "memedan", i).strip())
            기준가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "recprice", i).strip())
            구분 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "gubun", i).strip())
            증권그룹 = self.ActiveX.GetFieldData(self.OUTBLOCK, "bu12gubun", i).strip()
            기업인수목적회사여부 = self.ActiveX.GetFieldData(self.OUTBLOCK, "spac_gubun", i).strip()

            lst = [종목명, 단축코드, 확장코드, ETF구분, 상한가, 하한가, 전일가, 주문수량단위, 기준가, 구분, 증권그룹, 기업인수목적회사여부]
            result.append(lst)

        columns = ['종목명', '단축코드', '확장코드', 'ETF구분', '상한가', '하한가', '전일가', '주문수량단위', '기준가', '구분', '증권그룹', '기업인수목적회사여부']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 종목검색
class t1833(XAQuery):
    def Query(self, 종목검색파일=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "dummy", 0, "")
        self.ActiveX.RequestService(self.MYNAME, 종목검색파일)

    def OnReceiveData(self, szTrCode):
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            검색종목수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "JongCnt", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            연속봉수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "signcnt", i).strip())
            현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [종목코드, 종목명, 전일대비구분, 연속봉수, 현재가, 전일대비, 등락율, 거래량]
            result.append(lst)

        columns = ['종목코드', '종목명', '전일대비구분', '연속봉수', '현재가', '전일대비', '등락율', '거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [검색종목수,df])


# e종목검색
class t1857(XAQuery):
    def Query(self, 실시간구분,종목검색구분,종목검색입력값):
        self.실시간키 = ''
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.ClearBlockdata(self.OUTBLOCK)
        self.ActiveX.ClearBlockdata(self.OUTBLOCK1)
        self.ActiveX.SetFieldData(self.INBLOCK, "sRealFlag", 0, 실시간구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "sSearchFlag", 0, 종목검색구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "query_index", 0, 종목검색입력값)
        self.ActiveX.RequestService(self.MYNAME, "")

    def OnReceiveData(self, szTrCode):
        try:
            검색종목수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "result_count", 0).strip())
            포착시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "result_time", 0).strip()
            실시간키 = self.ActiveX.GetFieldData(self.OUTBLOCK, "AlertNum", 0).strip()
            self.실시간키 = 실시간키


            result = []
            for i in range(검색종목수):
                종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
                종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
                현재가 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
                전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
                전일대비 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
                등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
                거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
                종목상태 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "JobFlag", i).strip()

                lst = [종목코드, 종목명, 현재가, 전일대비구분, 전일대비, 등락율, 거래량, 종목상태]
                result.append(lst)

            columns = ['종목코드', '종목명', '현재가', '전일대비구분', '전일대비', '등락율', '거래량', '종목상태']
            df = DataFrame(data=result, columns=columns)

            if self.parent != None:
                self.parent.OnReceiveData(szTrCode, [self.식별자, 검색종목수,포착시간,실시간키,df])
        except Exception as e:
            pass

    def OnReceiveSearchRealData(self, szTrCode):
        result = dict()
        result['종목코드'] = self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "shcode").strip()
        result['종목명'] = self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "hname").strip()
        result['현재가'] = int(self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "price").strip())
        result['전일대비구분'] = self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "sign").strip()
        result['전일대비'] = int(self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "change").strip())
        result['등락율'] = float(self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "diff").strip())
        result['거래량'] = int(self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "volume").strip())
        result['종목상태'] = self.ActiveX.GetFieldSearchRealData(self.OUTBLOCK1, "JobFlag").strip()

        if self.parent != None:
            self.parent.OnReceiveSearchRealData(szTrCode, [self.식별자, result])

        # print(" EXIT : %s --> %s" % (클래스이름, 함수이름))

    def RemoveService(self):
        if self.실시간키 != '':
            result = self.ActiveX.RemoveService(self.MYNAME, self.실시간키)

# 주식종목코드조회(API용)
class t1866(XAQuery):
    def Query(self, 로그인ID, 조회구분, 그룹명, 연속여부, 연속키):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "user_id", 0, 로그인ID)
        self.ActiveX.SetFieldData(self.INBLOCK, "gb", 0, 조회구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "group_name", 0, 그룹명)
        self.ActiveX.SetFieldData(self.INBLOCK, "cont", 0, 연속여부)
        self.ActiveX.SetFieldData(self.INBLOCK, "cont_key", 0, 연속키)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        저장조건수 = self.ActiveX.GetFieldData(self.OUTBLOCK, "result_count", i).strip()
        연속여부 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cont", i).strip()
        연속키 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cont_key", i).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            서버저장인덱스 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "hname", i).strip()
            그룹명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "shcode", i).strip()
            조건저장명 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "expcode", i).strip()

            lst = [서버저장인덱스, 그룹명, 조건저장명]
            result.append(lst)

        columns = ['서버저장인덱스', '그룹명', '조건저장명']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [저장조건수, 연속여부, 연속키, df])

# 뉴스본문(t3102)
class t3102(XAQuery):
    def Query(self, 로그인ID, 조회구분, 그룹명, 연속여부, 연속키):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "sNewsno", 0, 뉴스번호)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        뉴스종목 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sJongcode", i).strip()

        뉴스본문 = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            body = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sBody", i).strip()
            뉴스본문.append(body)

        뉴스타이틀 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "sTitle", i).strip()

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [뉴스종목, 뉴스본문, 뉴스타이틀])

# 차트데이타조회
class ChartIndex(XAQuery):
    def Query(self, 지표ID='',지표명='',지표조건설정='',시장구분='',주기구분='',단축코드='',요청건수='500',단위='',시작일자='',종료일자='',수정주가반영여부='',갭보정여부='',실시간데이터수신자동등록여부='0'):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "indexid", 0, 지표ID)
        self.ActiveX.SetFieldData(self.INBLOCK, "indexname", 0, 지표명)
        self.ActiveX.SetFieldData(self.INBLOCK, "indexparam", 0, 지표조건설정)
        self.ActiveX.SetFieldData(self.INBLOCK, "market", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "period", 0, 주기구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
        self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
        self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
        self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
        self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
        self.ActiveX.SetFieldData(self.INBLOCK, "Isamend", 0, 수정주가반영여부)
        self.ActiveX.SetFieldData(self.INBLOCK, "Isgab", 0, 갭보정여부)
        self.ActiveX.SetFieldData(self.INBLOCK, "IsReal", 0, 실시간데이터수신자동등록여부)
        self.ActiveX.RequestService("ChartIndex", "")

    def RemoveService(self):
        try:
            지표ID = self.ActiveX.GetFieldData(self.OUTBLOCK, "indexid", 0).strip()
            self.ActiveX.RemoveService("ChartIndex", 지표ID)
        except Exception as e:
            pass

    def OnReceiveData(self, szTrCode):
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            지표ID = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "indexid", i).strip())
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_cnt", i).strip())
            유효데이터컬럼갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "validdata_cnt", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            지표값1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value1", i).strip())
            지표값2 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value2", i).strip())
            지표값3 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value3", i).strip())
            지표값4 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value4", i).strip())
            지표값5 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value5", i).strip())
            위치 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "pos", i).strip())

            lst = [일자, 시간, 시가, 고가, 저가, 종가, 거래량, 지표값1, 지표값2, 지표값3, 지표값4, 지표값5, 위치]
            result.append(lst)

        columns = ['일자', '시간', '시가', '고가', '저가', '종가', '거래량', '지표값1', '지표값2', '지표값3', '지표값4', '지표값5', '위치']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [self.식별자, 지표ID,레코드갯수,유효데이터컬럼갯수,df])

    def OnReceiveChartRealData(self, szTrCode):
        지표ID = self.ActiveX.GetFieldChartRealData(self.OUTBLOCK, "indexid").strip()
        레코드갯수 = self.ActiveX.GetFieldChartRealData(self.OUTBLOCK, "rec_cnt").strip()
        유효데이터컬럼갯수 = self.ActiveX.GetFieldChartRealData(self.OUTBLOCK, "validdata_cnt").strip()

        result = dict()
        result['일자'] = self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "date").strip()
        result['시간'] = self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "time").strip()
        result['시가'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "open").strip())
        result['고가'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "high").strip())
        result['저가'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "low").strip())
        result['종가'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "close").strip())
        result['거래량'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "volume").strip())
        result['지표값1'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "value1").strip())
        result['지표값2'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "value2").strip())
        result['지표값3'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "value3").strip())
        result['지표값4'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "value4").strip())
        result['지표값5'] = float(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "value5").strip())
        result['위치'] = int(self.ActiveX.GetFieldChartRealData(self.OUTBLOCK1, "pos").strip())

        if self.parent != None:
            self.parent.OnReceiveChartRealData(szTrCode, [self.식별자, 지표ID,레코드갯수,유효데이터컬럼갯수,result])


##----------------------------------------------------------------------------------------------------------------------
# 선물옵션 정상주문,CFOAT00100
class CFOAT00100(XAQuery):
    def Query(self, 계좌번호, 비밀번호, 선물옵션종목번호, 매매구분, 선물옵션호가유형코드, 주문가격, 주문수량):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        if 선물옵션호가유형코드 == '03':
            주문가격 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FnoIsuNo", 0, 선물옵션종목번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FnoOrdprcPtnCode", 0, 선물옵션호가유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdPrc", 0, 주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            주문시장코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdMktCode", i).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", i).strip()
            선물옵션종목번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FnoIsuNo", i).strip()
            매매구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", i).strip()
            선물옵션주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FnoOrdPtnCode", i).strip()
            선물옵션호가유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FnoOrdprcPtnCode", i).strip()
            선물옵션거래유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FnoTrdPtnCode", i).strip()
            주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdPrc", i).strip())
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            통신매체코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CommdaCode", i).strip()
            협의매매완료시각 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "DscusBnsCmpltTime", i).strip()
            그룹ID = self.ActiveX.GetFieldData(self.OUTBLOCK1, "GrpId", i).strip()
            주문일련번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdSeqno", i).strip())
            포트폴리오번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "PtflNo", i).strip())
            바스켓번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "BskNo", i).strip())
            트렌치번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrchNo", i).strip())
            항목번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "ItemNo", i).strip())
            운용지시번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OpDrtnNo", i).strip()
            관리사원번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "MgempNo", i).strip()
            펀드ID = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FundId", i).strip()
            펀드주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "FundOrdNo", i).strip())

            lst = [레코드갯수,주문시장코드,계좌번호,비밀번호,선물옵션종목번호,매매구분,선물옵션주문유형코드,선물옵션호가유형코드,선물옵션거래유형코드,주문가격,주문수량,통신매체코드,협의매매완료시각,그룹ID,주문일련번호,포트폴리오번호,바스켓번호,트렌치번호,항목번호,운용지시번호,관리사원번호,펀드ID,펀드주문번호]
            result.append(lst)

        columns = ['레코드갯수','주문시장코드','계좌번호','비밀번호','선물옵션종목번호','매매구분','선물옵션주문유형코드','선물옵션호가유형코드','선물옵션거래유형코드','주문가격','주문수량','통신매체코드','협의매매완료시각','그룹ID','주문일련번호','포트폴리오번호','바스켓번호','트렌치번호','항목번호','운용지시번호','관리사원번호','펀드ID','펀드주문번호']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            주문번호 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdNo", i).strip())
            지점명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BrnNm", i).strip()
            계좌명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNm", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()
            주문가능금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAbleAmt", i).strip())
            현금주문가능금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MnyOrdAbleAmt", i).strip())
            주문증거금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdMgn", i).strip())
            현금주문증거금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MnyOrdMgn", i).strip())
            주문가능수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAbleQty", i).strip())

            lst = [레코드갯수,주문번호,지점명,계좌명,종목명,주문가능금액,현금주문가능금액,주문증거금,현금주문증거금,주문가능수량]
            result.append(lst)

        columns = ['레코드갯수','주문번호','지점명','계좌명','종목명','주문가능금액','현금주문가능금액','주문증거금','현금주문증거금','주문가능수량']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 선물옵션차트(틱/n틱)(t8414)
class t8414(XAQuery):
    # def Query(self, 단축코드='201N7302', 단위='15', 요청건수='2000', 조회영업일수='0', 시작일자='20180629', 시작시간='', 종료일자='20180629', 종료시간='', 연속일자='', 연속시간='', 압축여부='N', 연속조회=False):
    def Query(self, 단축코드,단위='1',요청건수='2000',조회영업일수='0',시작일자='',시작시간='',종료일자='',종료시간='',연속일자='',연속시간='',압축여부='N', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "nday", 0, 조회영업일수)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "stime", 0, 시작시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "etime", 0, 종료시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "comp_yn", 0, 압축여부)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        block = dict()
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            block['단축코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            block['전일시가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jisiga", i).strip())
            block['전일고가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jihigh", i).strip())
            block['전일저가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jilow", i).strip())
            block['전일종가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jiclose", i).strip())
            block['전일거래량'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jivolume", i).strip())
            block['당일시가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "disiga", i).strip())
            block['당일고가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dihigh", i).strip())
            block['당일저가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dilow", i).strip())
            block['당일종가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diclose", i).strip())
            block['상한가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "highend", i).strip())
            block['하한가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "lowend", i).strip())
            block['연속일자'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", i).strip()
            block['연속시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", i).strip()
            block['장시작시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "s_time", i).strip()
            block['장종료시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "e_time", i).strip()
            block['동시호가처리시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dshmin", i).strip()
            block['레코드카운트'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jdiff_vol", i).strip())
            미결제약정 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "openyak", i).strip())

            lst = [날짜,시간,시가,고가,저가,종가,거래량,미결제약정]

            result.append(lst)

        columns = ['날짜','시간','시가','고가','저가','종가','거래량','미결제약정']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [block, df])

# 선물/옵션챠트(N분)(t8415)
class t8415(XAQuery):
    # def Query(self, 단축코드='201N7302', 단위='15', 요청건수='2000', 조회영업일수='0', 시작일자='20180629', 시작시간='', 종료일자='20180629', 종료시간='', 연속일자='', 연속시간='', 압축여부='N', 연속조회=False):
    def Query(self, 단축코드,단위='1',요청건수='2000',조회영업일수='0',시작일자='',시작시간='',종료일자='',종료시간='',연속일자='',연속시간='',압축여부='N', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "nday", 0, 조회영업일수)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "stime", 0, 시작시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "etime", 0, 종료시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "comp_yn", 0, 압축여부)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        block = dict()
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            block['단축코드'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", i).strip()
            block['전일시가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jisiga", i).strip())
            block['전일고가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jihigh", i).strip())
            block['전일저가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jilow", i).strip())
            block['전일종가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jiclose", i).strip())
            block['전일거래량'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jivolume", i).strip())
            block['당일시가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "disiga", i).strip())
            block['당일고가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dihigh", i).strip())
            block['당일저가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dilow", i).strip())
            block['당일종가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diclose", i).strip())
            block['상한가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "highend", i).strip())
            block['하한가'] = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "lowend", i).strip())
            block['연속일자'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", i).strip()
            block['연속시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", i).strip()
            block['장시작시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "s_time", i).strip()
            block['장종료시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "e_time", i).strip()
            block['동시호가처리시간'] = self.ActiveX.GetFieldData(self.OUTBLOCK, "dshmin", i).strip()
            block['레코드카운트'] = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", i).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "jdiff_vol", i).strip())
            거래대금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "value", i).strip())
            미결제약정 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "openyak", i).strip())

            lst = [날짜,시간,시가,고가,저가,종가,누적거래량,거래대금,미결제약정]

            result.append(lst)

        columns = ['날짜','시간','시가','고가','저가','종가','누적거래량','거래대금','미결제약정']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [block, df])

##----------------------------------------------------------------------------------------------------------------------
# 해외선물

# 해외선물 체결내역개별 조회
class CIDBQ01400(XAQuery):
    def Query(self, 레코드갯수='',조회구분코드='',계좌번호='',종목코드값='',매매구분코드='',해외파생주문가격='',해외선물주문유형코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QryTpCode", 0, 조회구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtOrdPrc", 0, 해외파생주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AbrdFutsOrdPtnCode", 0, 해외선물주문유형코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
            조회구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QryTpCode", 0).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", 0).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", 0).strip()
            해외파생주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtOrdPrc", 0).strip())
            해외선물주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AbrdFutsOrdPtnCode", 0).strip()

            lst = [레코드갯수,조회구분코드,계좌번호,종목코드값,매매구분코드,해외파생주문가격,해외선물주문유형코드]
            result.append(lst)

        columns = ['레코드갯수','조회구분코드','계좌번호','종목코드값','매매구분코드','해외파생주문가격','해외선물주문유형코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip()
            주문가능수량 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAbleQty", i).strip()

            lst = [레코드갯수,주문가능수량]
            result.append(lst)

        columns = ['레코드갯수','주문가능수량']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물 미결제 잔고내역
class CIDBQ01500(XAQuery):
    def Query(self, 레코드갯수='',계좌구분코드='',계좌번호='',FCM계좌번호='',비밀번호='',조회일자='',잔고구분코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntTpCode", 0, 계좌구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FcmAcntNo", 0, FCM계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QryDt", 0, 조회일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BalTpCode", 0, 잔고구분코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
            계좌구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntTpCode", 0).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
            FCM계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FcmAcntNo", 0).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", 0).strip()
            조회일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QryDt", 0).strip()
            잔고구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BalTpCode", 0).strip()

            lst = [레코드갯수,계좌구분코드,계좌번호,FCM계좌번호,비밀번호,조회일자,잔고구분코드]
            result.append(lst)

        columns = ['레코드갯수','계좌구분코드','계좌번호','FCM계좌번호','비밀번호','조회일자','잔고구분코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            기준일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BaseDt", i).strip()
            예수금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "Dps", i).strip())
            청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "LpnlAmt", i).strip())
            선물만기전청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsDueBfLpnlAmt", i).strip())
            선물만기전수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsDueBfCmsn", i).strip())
            위탁증거금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CsgnMgn", i).strip())
            유지증거금 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MaintMgn", i).strip())
            신용한도금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CtlmtAmt", i).strip())
            추가증거금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AddMgn", i).strip())
            마진콜율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MgnclRat", i).strip())
            주문가능금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdAbleAmt", i).strip())
            인출가능금액 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "WthdwAbleAmt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuCodeVal", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()
            통화코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CrcyCodeVal", i).strip()
            해외파생상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtPrdtCode", i).strip()
            해외파생옵션구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtOptTpCode", i).strip()
            만기일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "DueDt", i).strip()
            해외파생행사가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtXrcPrc", i).strip())
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpCode", i).strip()
            공통코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CmnCodeNm", i).strip()
            구분코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TpCodeNm", i).strip()
            잔고수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "BalQty", i).strip())
            매입가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PchsPrc", i).strip())
            해외파생현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtNowPrc", i).strip())
            해외선물평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsEvalPnlAmt", i).strip())
            위탁수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CsgnCmsn", i).strip())
            포지션번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "PosNo", i).strip()
            거래소비용1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "EufOneCmsnAmt", i).strip())
            거래소비용2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "EufTwoCmsnAmt", i).strip())

            lst = [
                기준일자,예수금,청산손익금액,선물만기전청산손익금액,선물만기전수수료,위탁증거금액,유지증거금,신용한도금액,추가증거금액,마진콜율,주문가능금액,
                인출가능금액,계좌번호,종목코드값,종목명,통화코드값,해외파생상품코드,해외파생옵션구분코드,만기일자,해외파생행사가격,매매구분코드,공통코드명,
                구분코드명,잔고수량,매입가격,해외파생현재가,해외선물평가손익금액,위탁수수료,포지션번호,거래소비용1수수료금액,거래소비용2수수료금액
            ]
            result.append(lst)

        columns = ['기준일자','예수금','청산손익금액','선물만기전청산손익금액','선물만기전수수료','위탁증거금액','유지증거금','신용한도금액','추가증거금액','마진콜율','주문가능금액','인출가능금액','계좌번호','종목코드값','종목명','통화코드값','해외파생상품코드','해외파생옵션구분코드','만기일자','해외파생행사가격','매매구분코드','공통코드명','구분코드명','잔고수량','매입가격','해외파생현재가','해외선물평가손익금액','위탁수수료','포지션번호','거래소비용1수수료금액','거래소비용2수수료금액']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물 주문체결내역 조회
class CIDBQ01800(XAQuery):
    def Query(self, 레코드갯수='',계좌번호='',비밀번호='',종목코드값='',주문일자='',당일구분코드='',주문상태코드='',매매구분코드='',조회구분코드='',주문유형코드='',해외파생선물옵션구분코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdDt", 0, 주문일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "ThdayTpCode", 0, 당일구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdStatCode", 0, 주문상태코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QryTpCode", 0, 조회구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdPtnCode", 0, 주문유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtFnoTpCode", 0, 해외파생선물옵션구분코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", 0).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", 0).strip()
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdDt", 0).strip()
            당일구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ThdayTpCode", 0).strip()
            주문상태코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdStatCode", 0).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", 0).strip()
            조회구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QryTpCode", 0).strip()
            주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdPtnCode", 0).strip()
            해외파생선물옵션구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtFnoTpCode", 0).strip()

            lst = [레코드갯수,계좌번호,비밀번호,종목코드값,주문일자,당일구분코드,주문상태코드,매매구분코드,조회구분코드,주문유형코드,해외파생선물옵션구분코드]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','비밀번호','종목코드값','주문일자','당일구분코드','주문상태코드','매매구분코드','조회구분코드','주문유형코드','해외파생선물옵션구분코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            해외선물주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrdNo", i).strip()
            해외선물원주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrgOrdNo", i).strip()
            FCM주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcmOrdNo", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuCodeVal", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()
            해외선물행사가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsXrcPrc", i).strip())
            FCM계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcmAcntNo", i).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpCode", i).strip()
            매매구분명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpNm", i).strip()
            선물주문상태코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsOrdStatCode", i).strip()
            구분코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TpCodeNm", i).strip()
            선물주문구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsOrdTpCode", i).strip()
            거래구분명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdTpNm", i).strip()
            해외선물주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsOrdPtnCode", i).strip()
            주문유형명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnNm", i).strip()
            주문유형기간구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnTermTpCode", i).strip()
            공통코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CmnCodeNm", i).strip()
            적용시작일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AppSrtDt", i).strip()
            적용종료일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AppEndDt", i).strip()
            해외파생주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtOrdPrc", i).strip())
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdQty", i).strip())
            해외선물체결가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsExecPrc", i).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "ExecQty", i).strip())
            주문조건가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdCndiPrc", i).strip())
            해외파생현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtNowPrc", i).strip())
            정정수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "MdfyQty", i).strip())
            취소수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CancQty", i).strip())
            거부수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RjtQty", i).strip())
            확인수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CnfQty", i).strip())
            반대매매여부 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CvrgYn", i).strip()
            등록단말번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RegTmnlNo", i).strip()
            등록지점번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RegBrnNo", i).strip()
            등록사용자ID = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RegUserId", i).strip()
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdDt", i).strip()
            주문시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdTime", i).strip()
            해외옵션행사예약구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsOptXrcRsvTpCode", i).strip()

            lst = [
                해외선물주문번호,해외선물원주문번호,FCM주문번호,종목코드값,종목명,해외선물행사가격,FCM계좌번호,매매구분코드,매매구분명,선물주문상태코드,
                구분코드명,선물주문구분코드,거래구분명,해외선물주문유형코드,주문유형명,주문유형기간구분코드,공통코드명,적용시작일자,적용종료일자,해외파생주문가격,
                주문수량,해외선물체결가격,체결수량,주문조건가격,해외파생현재가,정정수량,취소수량,거부수량,확인수량,반대매매여부,등록단말번호,등록지점번호,등록사용자ID,
                주문일자,주문시각,해외옵션행사예약구분코드
            ]
            result.append(lst)

        columns = ['해외선물주문번호','해외선물원주문번호','FCM주문번호','종목코드값','종목명','해외선물행사가격','FCM계좌번호','매매구분코드','매매구분명','선물주문상태코드','구분코드명','선물주문구분코드','거래구분명','해외선물주문유형코드','주문유형명','주문유형기간구분코드','공통코드명','적용시작일자','적용종료일자','해외파생주문가격','주문수량','해외선물체결가격','체결수량','주문조건가격','해외파생현재가','정정수량','취소수량','거부수량','확인수량','반대매매여부','등록단말번호','등록지점번호','등록사용자ID','주문일자','주문시각','해외옵션행사예약구분코드']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물 주문체결내역 상세 조회
class CIDBQ02400(XAQuery):
    def Query(self, 레코드갯수='',계좌번호='',비밀번호='',종목코드값='',조회시작일자='',조회종료일자='',당일구분코드='',주문상태코드='',매매구분코드='',조회구분코드='',주문유형코드='',해외파생선물옵션구분코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QrySrtDt", 0, 조회시작일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QryEndDt", 0, 조회종료일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "ThdayTpCode", 0, 당일구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdStatCode", 0, 주문상태코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "QryTpCode", 0, 조회구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdPtnCode", 0, 주문유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtFnoTpCode", 0, 해외파생선물옵션구분코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", 0).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", 0).strip()
            조회시작일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QrySrtDt", 0).strip()
            조회종료일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QryEndDt", 0).strip()
            당일구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ThdayTpCode", 0).strip()
            주문상태코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdStatCode", 0).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", 0).strip()
            조회구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "QryTpCode", 0).strip()
            주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdPtnCode", 0).strip()
            해외파생선물옵션구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtFnoTpCode", 0).strip()

            lst = [레코드갯수,계좌번호,비밀번호,종목코드값,조회시작일자,조회종료일자,당일구분코드,주문상태코드,매매구분코드,조회구분코드,주문유형코드,해외파생선물옵션구분코드]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','비밀번호','종목코드값','조회시작일자','조회종료일자','당일구분코드','주문상태코드','매매구분코드','조회구분코드','주문유형코드','해외파생선물옵션구분코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdDt", i).strip()
            해외선물주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrdNo", i).strip()
            해외선물원주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrgOrdNo", i).strip()
            FCM주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcmOrdNo", i).strip()
            해외선물체결번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsExecNo", i).strip()
            FCM계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcmAcntNo", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuCodeVal", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()
            해외선물행사가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsXrcPrc", i).strip())
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpCode", i).strip()
            매매구분명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpNm", i).strip()
            선물주문상태코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsOrdStatCode", i).strip()
            구분코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TpCodeNm", i).strip()
            선물주문구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsOrdTpCode", i).strip()
            거래구분명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdTpNm", i).strip()
            해외선물주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsOrdPtnCode", i).strip()
            주문유형명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnNm", i).strip()
            주문유형기간구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdPtnTermTpCode", i).strip()
            공통코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CmnCodeNm", i).strip()
            적용시작일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AppSrtDt", i).strip()
            적용종료일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AppEndDt", i).strip()
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdQty", i).strip())
            해외파생주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtOrdPrc", i).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "ExecQty", i).strip())
            해외선물체결가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsExecPrc", i).strip())
            주문조건가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdCndiPrc", i).strip())
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "NowPrc", i).strip())
            처리상태코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrxStatCode", i).strip()
            처리상태코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrxStatCodeNm", i).strip()
            위탁수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CsgnCmsn", i).strip())
            FCM수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcmCmsn", i).strip())
            당사수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "ThcoCmsn", i).strip())
            매체코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MdaCode", i).strip()
            매체코드명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "MdaCodeNm", i).strip()
            등록단말번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RegTmnlNo", i).strip()
            등록사용자ID = self.ActiveX.GetFieldData(self.OUTBLOCK2, "RegUserId", i).strip()
            주문일시 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdDttm", i).strip()
            주문시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OrdTime", i).strip()
            체결일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "ExecDt", i).strip()
            체결시각 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "ExecTime", i).strip()
            거래소비용1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "EufOneCmsnAmt", i).strip())
            거래소비용2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "EufTwoCmsnAmt", i).strip())
            런던청산소1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "LchOneCmsnAmt", i).strip())
            런던청산소2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "LchTwoCmsnAmt", i).strip())
            거래1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdOneCmsnAmt", i).strip())
            거래2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdTwoCmsnAmt", i).strip())
            거래3수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdThreeCmsnAmt", i).strip())
            단기1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "StrmOneCmsnAmt", i).strip())
            단기2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "StrmTwoCmsnAmt", i).strip())
            단기3수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "StrmThreeCmsnAmt", i).strip())
            전달1수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TransOneCmsnAmt", i).strip())
            전달2수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TransTwoCmsnAmt", i).strip())
            전달3수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TransThreeCmsnAmt", i).strip())
            전달4수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "TransFourCmsnAmt", i).strip())
            해외옵션행사예약구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsOptXrcRsvTpCode", i).strip()

            lst = [
                주문일자,해외선물주문번호,해외선물원주문번호,FCM주문번호,해외선물체결번호,FCM계좌번호,종목코드값,종목명,해외선물행사가격,매매구분코드,매매구분명,
                선물주문상태코드,구분코드명,선물주문구분코드,거래구분명,해외선물주문유형코드,주문유형명,주문유형기간구분코드,공통코드명,적용시작일자,적용종료일자,
                주문수량,해외파생주문가격,체결수량,해외선물체결가격,주문조건가격,현재가,처리상태코드,처리상태코드명,위탁수수료,FCM수수료,당사수수료,매체코드,
                매체코드명,등록단말번호,등록사용자ID,주문일시,주문시각,체결일자,체결시각,거래소비용1수수료금액,거래소비용2수수료금액,런던청산소1수수료금액,
                런던청산소2수수료금액,거래1수수료금액,거래2수수료금액,거래3수수료금액,단기1수수료금액,단기2수수료금액,단기3수수료금액,전달1수수료금액,전달2수수료금액,
                전달3수수료금액,전달4수수료금액,해외옵션행사예약구분코드
            ]
            result.append(lst)

        columns = ['주문일자','해외선물주문번호','해외선물원주문번호','FCM주문번호','해외선물체결번호','FCM계좌번호','종목코드값','종목명','해외선물행사가격','매매구분코드','매매구분명','선물주문상태코드','구분코드명','선물주문구분코드','거래구분명','해외선물주문유형코드','주문유형명','주문유형기간구분코드','공통코드명','적용시작일자','적용종료일자','주문수량','해외파생주문가격','체결수량','해외선물체결가격','주문조건가격','현재가','처리상태코드','처리상태코드명','위탁수수료','FCM수수료','당사수수료','매체코드','매체코드명','등록단말번호','등록사용자ID','주문일시','주문시각','체결일자','체결시각','거래소비용1수수료금액','거래소비용2수수료금액','런던청산소1수수료금액','런던청산소2수수료금액','거래1수수료금액','거래2수수료금액','거래3수수료금액','단기1수수료금액','단기2수수료금액','단기3수수료금액','전달1수수료금액','전달2수수료금액','전달3수수료금액','전달4수수료금액','해외옵션행사예약구분코드']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물 예수금/잔고현황
class CIDBQ03000(XAQuery):
    def Query(self, 레코드갯수='',계좌구분코드='',계좌번호='',계좌비밀번호='',거래일자=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntTpCode", 0, 계좌구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntPwd", 0, 계좌비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "TrdDt", 0, 거래일자)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
        계좌구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntTpCode", 0).strip()
        계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
        계좌비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntPwd", 0).strip()
        거래일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrdDt", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            거래일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdDt", i).strip()
            통화대상코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CrcyObjCode", i).strip()
            해외선물예수금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsDps", i).strip())
            고객입출금금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CustmMnyioAmt", i).strip())
            해외선물청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsLqdtPnlAmt", i).strip())
            해외선물수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsCmsnAmt", i).strip())
            가환전예수금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrexchDps", i).strip())
            평가자산금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "EvalAssetAmt", i).strip())
            해외선물위탁증거금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsCsgnMgn", i).strip())
            해외선물추가증거금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsAddMgn", i).strip())
            해외선물인출가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsWthdwAbleAmt", i).strip())
            해외선물주문가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsOrdAbleAmt", i).strip())
            해외선물평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsEvalPnlAmt", i).strip())
            최종결제손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "LastSettPnlAmt", i).strip())
            해외옵션결제금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsOptSettAmt", i).strip())
            해외옵션잔고평가금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsOptBalEvalAmt", i).strip())

            lst = [
                계좌번호,거래일자,통화대상코드,해외선물예수금,고객입출금금액,해외선물청산손익금액,해외선물수수료금액,가환전예수금,평가자산금액,해외선물위탁증거금액,
                해외선물추가증거금액,해외선물인출가능금액,해외선물주문가능금액,해외선물평가손익금액,최종결제손익금액,해외옵션결제금액,해외옵션잔고평가금액
            ]
            result.append(lst)

        columns = ['계좌번호','거래일자','통화대상코드','해외선물예수금','고객입출금금액','해외선물청산손익금액','해외선물수수료금액','가환전예수금','평가자산금액','해외선물위탁증거금액','해외선물추가증거금액','해외선물인출가능금액','해외선물주문가능금액','해외선물평가손익금액','최종결제손익금액','해외옵션결제금액','해외옵션잔고평가금액']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [레코드갯수,계좌구분코드,계좌번호,계좌비밀번호,거래일자, df1])


# 해외선물 계좌예탁자산조회
class CIDBQ05300(XAQuery):
    def Query(self, 레코드갯수='',해외계좌구분코드='',FCM계좌번호='',계좌번호='',계좌비밀번호='',통화코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsAcntTpCode", 0, 해외계좌구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FcmAcntNo", 0, FCM계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntPwd", 0, 계좌비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "CrcyCode", 0, 통화코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", 0).strip())
        해외계좌구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsAcntTpCode", 0).strip()
        FCM계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FcmAcntNo", 0).strip()
        계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", 0).strip()
        계좌비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntPwd", 0).strip()
        통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CrcyCode", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CrcyCode", i).strip()
            해외선물예수금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsDps", i).strip())
            해외선물위탁증거금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsCsgnMgn", i).strip())
            해외선물추가증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsSplmMgn", i).strip())
            고객청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CustmLpnlAmt", i).strip())
            해외선물평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsEvalPnlAmt", i).strip())
            해외선물수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsCmsnAmt", i).strip())
            해외선물평가예탁총금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsEvalDpstgTotAmt", i).strip())
            환율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "Xchrat", i).strip())
            외화실환전금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcurrRealMxchgAmt", i).strip())
            해외선물인출가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsWthdwAbleAmt", i).strip())
            해외선물주문가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsOrdAbleAmt", i).strip())
            선물만기미도래청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsDueNarrvLqdtPnlAmt", i).strip())
            선물만기미도래수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FutsDueNarrvCmsn", i).strip())
            해외선물청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsLqdtPnlAmt", i).strip())
            해외선물만기수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsDueCmsn", i).strip())
            해외선물옵션매수금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOptBuyAmt", i).strip())
            해외선물옵션매도금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOptSellAmt", i).strip())
            옵션매수시장가치금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OptBuyMktWrthAmt", i).strip())
            옵션매도시장가치금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OptSellMktWrthAmt", i).strip())

            lst = [
                계좌번호,통화코드,해외선물예수금,해외선물위탁증거금액,해외선물추가증거금,고객청산손익금액,해외선물평가손익금액,해외선물수수료금액,
                해외선물평가예탁총금액,환율,외화실환전금액,해외선물인출가능금액,해외선물주문가능금액,선물만기미도래청산손익금액,선물만기미도래수수료,
                해외선물청산손익금액,해외선물만기수수료,해외선물옵션매수금액,해외선물옵션매도금액,옵션매수시장가치금액,옵션매도시장가치금액
            ]
            result.append(lst)

        columns = ['계좌번호','통화코드','해외선물예수금','해외선물위탁증거금액','해외선물추가증거금','고객청산손익금액','해외선물평가손익금액','해외선물수수료금액','해외선물평가예탁총금액','환율','외화실환전금액','해외선물인출가능금액','해외선물주문가능금액','선물만기미도래청산손익금액','선물만기미도래수수료','해외선물청산손익금액','해외선물만기수수료','해외선물옵션매수금액','해외선물옵션매도금액','옵션매수시장가치금액','옵션매도시장가치금액']
        df1 = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK3)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK3, "RecCnt", i).strip())
            해외선물예수금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsDps", i).strip())
            해외선물청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsLqdtPnlAmt", i).strip())
            선물만기미도래청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "FutsDueNarrvLqdtPnlAmt", i).strip())
            해외선물평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsEvalPnlAmt", i).strip())
            해외선물평가예탁총금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsEvalDpstgTotAmt", i).strip())
            고객청산손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "CustmLpnlAmt", i).strip())
            해외선물만기수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsDueCmsn", i).strip())
            외화실환전금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "FcurrRealMxchgAmt", i).strip())
            해외선물수수료금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsCmsnAmt", i).strip())
            선물만기미도래수수료 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "FutsDueNarrvCmsn", i).strip())
            해외선물위탁증거금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsCsgnMgn", i).strip())
            해외선물유지증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsMaintMgn", i).strip())
            해외선물옵션매수금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsOptBuyAmt", i).strip())
            해외선물옵션매도금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsOptSellAmt", i).strip())
            신용한도금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "CtlmtAmt", i).strip())
            해외선물추가증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsFutsSplmMgn", i).strip())
            마진콜율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "MgnclRat", i).strip())
            해외선물주문가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsOrdAbleAmt", i).strip())
            해외선물인출가능금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "AbrdFutsWthdwAbleAmt", i).strip())
            옵션매수시장가치금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OptBuyMktWrthAmt", i).strip())
            옵션매도시장가치금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OptSellMktWrthAmt", i).strip())
            해외옵션결제금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsOptSettAmt", i).strip())
            해외옵션잔고평가금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK3, "OvrsOptBalEvalAmt", i).strip())

            lst = [
                레코드갯수,해외선물예수금,해외선물청산손익금액,선물만기미도래청산손익금액,해외선물평가손익금액,해외선물평가예탁총금액,고객청산손익금액,
                해외선물만기수수료,외화실환전금액,해외선물수수료금액,선물만기미도래수수료,해외선물위탁증거금액,해외선물유지증거금,해외선물옵션매수금액,
                해외선물옵션매도금액,신용한도금액,해외선물추가증거금,마진콜율,해외선물주문가능금액,해외선물인출가능금액,옵션매수시장가치금액,
                옵션매도시장가치금액,해외옵션결제금액,해외옵션잔고평가금액
            ]
            result.append(lst)

        columns = ['레코드갯수','해외선물예수금','해외선물청산손익금액','선물만기미도래청산손익금액','해외선물평가손익금액','해외선물평가예탁총금액','고객청산손익금액','해외선물만기수수료','외화실환전금액','해외선물수수료금액','선물만기미도래수수료','해외선물위탁증거금액','해외선물유지증거금','해외선물옵션매수금액','해외선물옵션매도금액','신용한도금액','해외선물추가증거금','마진콜율','해외선물주문가능금액','해외선물인출가능금액','옵션매수시장가치금액','옵션매도시장가치금액','해외옵션결제금액','해외옵션잔고평가금액']
        df2 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [레코드갯수,해외계좌구분코드,FCM계좌번호,계좌번호,계좌비밀번호,통화코드, df1, df2])


# 해외선물신규주문
class CIDBT00100(XAQuery):
    def Query(self, 레코드갯수='1',주문일자='',지점코드='',계좌번호='',비밀번호='',종목코드값='',선물주문구분코드='',매매구분코드='',해외선물주문유형코드='',통화코드='',해외파생주문가격='',조건주문가격='',주문수량='',상품코드='',만기년월='',거래소코드=''):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        if 해외선물주문유형코드 == '1':
            해외파생주문가격=''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdDt", 0, 주문일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BrnCode", 0, 지점코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FutsOrdTpCode", 0, 선물주문구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AbrdFutsOrdPtnCode", 0, 해외선물주문유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "CrcyCode", 0, 통화코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtOrdPrc", 0, 해외파생주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "CndiOrdPrc", 0, 조건주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.SetFieldData(self.INBLOCK1, "PrdtCode", 0, 상품코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "DueYymm", 0, 만기년월)
        self.ActiveX.SetFieldData(self.INBLOCK1, "ExchCode", 0, 거래소코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdDt", i).strip()
            지점코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BrnCode", i).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", i).strip()
            선물주문구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FutsOrdTpCode", i).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", i).strip()
            해외선물주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AbrdFutsOrdPtnCode", i).strip()
            통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CrcyCode", i).strip()
            해외파생주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtOrdPrc", i).strip())
            조건주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "CndiOrdPrc", i).strip())
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "PrdtCode", i).strip()
            만기년월 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "DueYymm", i).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ExchCode", i).strip()

            lst = [레코드갯수,주문일자,지점코드,계좌번호,비밀번호,종목코드값,선물주문구분코드,매매구분코드,해외선물주문유형코드,통화코드,해외파생주문가격,조건주문가격,주문수량,상품코드,만기년월,거래소코드]
            result.append(lst)

        columns = ['레코드갯수','주문일자','지점코드','계좌번호','비밀번호','종목코드값','선물주문구분코드','매매구분코드','해외선물주문유형코드','통화코드','해외파생주문가격','조건주문가격','주문수량','상품코드','만기년월','거래소코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            해외선물주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrdNo", i).strip()

            lst = [레코드갯수,계좌번호,해외선물주문번호]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','해외선물주문번호']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물정정주문
class CIDBT00900(XAQuery):
    def Query(self, 레코드갯수='',주문일자='',등록지점번호='',계좌번호='',비밀번호='',해외선물원주문번호='',종목코드값='',선물주문구분코드='',매매구분코드='',선물주문유형코드='',통화코드값='',해외파생주문가격='',조건주문가격='',주문수량='',해외파생상품코드='',만기년월='',거래소코드=''):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdDt", 0, 주문일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RegBrnNo", 0, 등록지점번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsFutsOrgOrdNo", 0, 해외선물원주문번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FutsOrdTpCode", 0, 선물주문구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BnsTpCode", 0, 매매구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FutsOrdPtnCode", 0, 선물주문유형코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "CrcyCodeVal", 0, 통화코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtOrdPrc", 0, 해외파생주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "CndiOrdPrc", 0, 조건주문가격)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdQty", 0, 주문수량)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsDrvtPrdtCode", 0, 해외파생상품코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "DueYymm", 0, 만기년월)
        self.ActiveX.SetFieldData(self.INBLOCK1, "ExchCode", 0, 거래소코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdDt", i).strip()
            등록지점번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "RegBrnNo", i).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", i).strip()
            해외선물원주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsFutsOrgOrdNo", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", i).strip()
            선물주문구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FutsOrdTpCode", i).strip()
            매매구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BnsTpCode", i).strip()
            선물주문유형코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FutsOrdPtnCode", i).strip()
            통화코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "CrcyCodeVal", i).strip()
            해외파생주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtOrdPrc", i).strip())
            조건주문가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "CndiOrdPrc", i).strip())
            주문수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdQty", i).strip())
            해외파생상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsDrvtPrdtCode", i).strip()
            만기년월 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "DueYymm", i).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ExchCode", i).strip()

            lst = [레코드갯수,주문일자,등록지점번호,계좌번호,비밀번호,해외선물원주문번호,종목코드값,선물주문구분코드,매매구분코드,선물주문유형코드,통화코드값,해외파생주문가격,조건주문가격,주문수량,해외파생상품코드,만기년월,거래소코드]
            result.append(lst)

        columns = ['레코드갯수','주문일자','등록지점번호','계좌번호','비밀번호','해외선물원주문번호','종목코드값','선물주문구분코드','매매구분코드','선물주문유형코드','통화코드값','해외파생주문가격','조건주문가격','주문수량','해외파생상품코드','만기년월','거래소코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            해외선물주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrdNo", i).strip()
            내부메시지내용 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "InnerMsgCnts", i).strip()

            lst = [레코드갯수,계좌번호,해외선물주문번호,내부메시지내용]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','해외선물주문번호','내부메시지내용']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물취소주문
class CIDBT01000(XAQuery):
    def Query(self, 레코드갯수='',주문일자='',지점번호='',계좌번호='',비밀번호='',종목코드값='',해외선물원주문번호='',선물주문구분코드='',상품구분코드='',거래소코드=''):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OrdDt", 0, 주문일자)
        self.ActiveX.SetFieldData(self.INBLOCK1, "BrnNo", 0, 지점번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "Pwd", 0, 비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "IsuCodeVal", 0, 종목코드값)
        self.ActiveX.SetFieldData(self.INBLOCK1, "OvrsFutsOrgOrdNo", 0, 해외선물원주문번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "FutsOrdTpCode", 0, 선물주문구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "PrdtTpCode", 0, 상품구분코드)
        self.ActiveX.SetFieldData(self.INBLOCK1, "ExchCode", 0, 거래소코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            주문일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OrdDt", i).strip()
            지점번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "BrnNo", i).strip()
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "Pwd", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "IsuCodeVal", i).strip()
            해외선물원주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "OvrsFutsOrgOrdNo", i).strip()
            선물주문구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "FutsOrdTpCode", i).strip()
            상품구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "PrdtTpCode", i).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ExchCode", i).strip()

            lst = [레코드갯수,주문일자,지점번호,계좌번호,비밀번호,종목코드값,해외선물원주문번호,선물주문구분코드,상품구분코드,거래소코드]
            result.append(lst)

        columns = ['레코드갯수','주문일자','지점번호','계좌번호','비밀번호','종목코드값','해외선물원주문번호','선물주문구분코드','상품구분코드','거래소코드']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "RecCnt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            해외선물주문번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsFutsOrdNo", i).strip()
            내부메시지내용 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "InnerMsgCnts", i).strip()

            lst = [레코드갯수,계좌번호,해외선물주문번호,내부메시지내용]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','해외선물주문번호','내부메시지내용']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 일자별 미결제 잔고내역
class CIDEQ00800(XAQuery):
    def Query(self, 레코드갯수='',계좌번호='',계좌비밀번호='',거래일자=''):
        self.주문결과코드 = ''
        self.주문결과메세지 = ''

        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK1, "RecCnt", 0, 레코드갯수)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntNo", 0, 계좌번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "AcntPwd", 0, 계좌비밀번호)
        self.ActiveX.SetFieldData(self.INBLOCK1, "TrdDt", 0, 거래일자)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "RecCnt", i).strip())
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntNo", i).strip()
            계좌비밀번호 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "AcntPwd", i).strip()
            거래일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "TrdDt", i).strip()

            lst = [레코드갯수,계좌번호,계좌비밀번호,거래일자]
            result.append(lst)

        columns = ['레코드갯수','계좌번호','계좌비밀번호','거래일자']
        df = DataFrame(data=result, columns=columns)

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK2)
        for i in range(nCount):
            계좌번호 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "AcntNo", i).strip()
            거래일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "TrdDt", i).strip()
            종목코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuCodeVal", i).strip()
            매매구분명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "BnsTpNm", i).strip()
            잔고수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "BalQty", i).strip())
            청산가능수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK2, "LqdtAbleQty", i).strip())
            매입가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PchsPrc", i).strip())
            해외파생현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtNowPrc", i).strip())
            해외선물평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "AbrdFutsEvalPnlAmt", i).strip())
            고객잔고금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "CustmBalAmt", i).strip())
            외화평가금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcurrEvalAmt", i).strip())
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "IsuNm", i).strip()
            통화코드값 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "CrcyCodeVal", i).strip()
            해외파생상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "OvrsDrvtPrdtCode", i).strip()
            만기일자 = self.ActiveX.GetFieldData(self.OUTBLOCK2, "DueDt", i).strip()
            계약당금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "PrcntrAmt", i).strip())
            외화평가손익금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK2, "FcurrEvalPnlAmt", i).strip())

            lst = [계좌번호,거래일자,종목코드값,매매구분명,잔고수량,청산가능수량,매입가격,해외파생현재가,해외선물평가손익금액,고객잔고금액,외화평가금액,종목명,통화코드값,해외파생상품코드,만기일자,계약당금액,외화평가손익금액]
            result.append(lst)

        columns = ['계좌번호','거래일자','종목코드값','매매구분명','잔고수량','청산가능수량','매입가격','해외파생현재가','해외선물평가손익금액','고객잔고금액','외화평가금액','종목명','통화코드값','해외파생상품코드','만기일자','계약당금액','외화평가손익금액']
        df1 = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df, df1])


# 해외선물마스터조회(o3101)-API용
class o3101(XAQuery):
    def Query(self):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, '')
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "SymbolNm", i).strip()
            종목배치수신일_한국일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ApplDate", i).strip()
            기초상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsCd", i).strip()
            기초상품명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsNm", i).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchCd", i).strip()
            거래소명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchNm", i).strip()
            기준통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "CrncyCd", i).strip()
            진법구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "NotaCd", i).strip()
            호가단위가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "UntPrc", i).strip())
            최소가격변동금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MnChgAmt", i).strip())
            가격조정계수 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "RgltFctr", i).strip())
            계약당금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CtrtPrAmt", i).strip())
            상품구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "GdsCd", i).strip()
            월물_년 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngYr", i).strip()
            월물_월 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngM", i).strip()
            정산가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "EcPrc", i).strip())
            거래시작시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlStrtTm", i).strip()
            거래종료시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlEndTm", i).strip()
            거래가능구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlPsblCd", i).strip()
            증거금징수구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MgnCltCd", i).strip()
            개시증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgn", i).strip())
            유지증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgn", i).strip())
            개시증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgnR", i).strip())
            유지증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgnR", i).strip())
            유효소수점자리수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "DotGb", i).strip())

            lst = [종목코드,종목명,종목배치수신일_한국일자,기초상품코드,기초상품명,거래소코드,거래소명,기준통화코드,진법구분코드,호가단위가격,최소가격변동금액,가격조정계수,계약당금액,상품구분코드,월물_년,월물_월,정산가격,거래시작시간,거래종료시간,거래가능구분코드,증거금징수구분코드,개시증거금,유지증거금,개시증거금율,유지증거금율,유효소수점자리수]
            result.append(lst)

        columns = ['종목코드','종목명','종목배치수신일_한국일자','기초상품코드','기초상품명','거래소코드','거래소명','기준통화코드','진법구분코드','호가단위가격','최소가격변동금액','가격조정계수','계약당금액','상품구분코드','월물_년','월물_월','정산가격','거래시작시간','거래종료시간','거래가능구분코드','증거금징수구분코드','개시증거금','유지증거금','개시증거금율','유지증거금율','유효소수점자리수']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물차트(분)(o3103)-API용
class o3103(XAQuery):
    def Query(self, 단축코드='',N분주기='',조회건수='',연속일자='',연속시간='', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, N분주기)
            self.ActiveX.SetFieldData(self.INBLOCK, "readcnt", 0, 조회건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)

            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cnt", i).strip()
            시차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "date", i).strip())
            조회건수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip())
            연속일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip()
            연속시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "idx", i).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            현지시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,현지시간,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','현지시간','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드, 시차, 조회건수, 연속일자, 연속시간, df])


# 해외선물일별체결조회(o3104)-API용
class o3104(XAQuery):
    def Query(self, 조회구분='',단축코드='',조회일자=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 조회구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
        self.ActiveX.SetFieldData(self.INBLOCK, "date", 0, 조회일자)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "chedate", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            체결구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "cgubun", i).strip()
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [일자,현재가,대비구분,대비,등락율,시가,고가,저가,체결구분,누적거래량]
            result.append(lst)

        columns = ['일자','현재가','대비구분','대비','등락율','시가','고가','저가','체결구분','누적거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물현재가(종목정보)조회(o3105)-API용
class o3105(XAQuery):
    def Query(self, 종목코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "symbol", 0, 종목코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "SymbolNm", i).strip()
            종목배치수신일 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ApplDate", i).strip()
            기초상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsCd", i).strip()
            기초상품명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsNm", i).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchCd", i).strip()
            거래소명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchNm", i).strip()
            정산구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "EcCd", i).strip()
            기준통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "CrncyCd", i).strip()
            진법구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "NotaCd", i).strip()
            호가단위가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "UntPrc", i).strip())
            최소가격변동금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MnChgAmt", i).strip())
            가격조정계수 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "RgltFctr", i).strip())
            계약당금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CtrtPrAmt", i).strip())
            상장개월수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngMCnt", i).strip())
            상품구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "GdsCd", i).strip()
            시장구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MrktCd", i).strip()
            Emini구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "EminiCd", i).strip()
            상장년 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngYr", i).strip()
            상장월 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngM", i).strip()
            월물순서 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "SeqNo", i).strip())
            상장일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngDt", i).strip()
            만기일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MtrtDt", i).strip()
            최종거래일 = self.ActiveX.GetFieldData(self.OUTBLOCK, "FnlDlDt", i).strip()
            최초인도통지일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "FstTrsfrDt", i).strip()
            정산가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "EcPrc", i).strip())
            거래시작일자_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlDt", i).strip()
            거래시작시간_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlStrtTm", i).strip()
            거래종료시간_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlEndTm", i).strip()
            거래시작일자_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsStrDay", i).strip()
            거래시작시간_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsStrTm", i).strip()
            거래종료일자_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsEndDay", i).strip()
            거래종료시간_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsEndTm", i).strip()
            거래가능구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlPsblCd", i).strip()
            증거금징수구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MgnCltCd", i).strip()
            개시증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgn", i).strip())
            유지증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgn", i).strip())
            개시증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgnR", i).strip())
            유지증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgnR", i).strip())
            유효소수점자리수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "DotGb", i).strip())
            시차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TimeDiff", i).strip())
            현지체결일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsDate", i).strip()
            한국체결일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "KorDate", i).strip()
            현지체결시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdTm", i).strip()
            한국체결시각 = self.ActiveX.GetFieldData(self.OUTBLOCK, "RcvTm", i).strip()
            체결가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdP", i).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdQ", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TotQ", i).strip())
            체결거래대금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdAmt", i).strip())
            누적거래대금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TotAmt", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpenP", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "HighP", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "LowP", i).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CloseP", i).strip())
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "YdiffP", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "YdiffSign", i).strip()
            체결구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Cgubun", i).strip()
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "Diff", i).strip())

            lst = [
                종목코드,종목명,종목배치수신일,기초상품코드,기초상품명,거래소코드,거래소명,정산구분코드,기준통화코드,진법구분코드,호가단위가격,
                최소가격변동금액,가격조정계수,계약당금액,상장개월수,상품구분코드,시장구분코드,Emini구분코드,상장년,상장월,월물순서,상장일자,만기일자,최종거래일,최초인도통지일자,정산가격,
                거래시작일자_한국,거래시작시간_한국,거래종료시간_한국,거래시작일자_현지,거래시작시간_현지,거래종료일자_현지,거래종료시간_현지,거래가능구분코드,증거금징수구분코드,개시증거금,유지증거금,개시증거금율,유지증거금율,유효소수점자리수,
                시차,현지체결일자,한국체결일자,현지체결시간,한국체결시각,체결가격,체결수량,누적거래량,체결거래대금,누적거래대금,시가,고가,저가,전일종가,전일대비,전일대비구분,체결구분,등락율
            ]
            result.append(lst)

        columns = ['종목코드','종목명','종목배치수신일','기초상품코드','기초상품명','거래소코드','거래소명','정산구분코드','기준통화코드','진법구분코드','호가단위가격','최소가격변동금액','가격조정계수','계약당금액','상장개월수','상품구분코드','시장구분코드','Emini구분코드','상장년','상장월','월물순서','상장일자','만기일자','최종거래일','최초인도통지일자','정산가격','거래시작일자_한국','거래시작시간_한국','거래종료시간_한국','거래시작일자_현지','거래시작시간_현지','거래종료일자_현지','거래종료시간_현지','거래가능구분코드','증거금징수구분코드','개시증거금','유지증거금','개시증거금율','유지증거금율','유효소수점자리수','시차','현지체결일자','한국체결일자','현지체결시간','한국체결시각','체결가격','체결수량','누적거래량','체결거래대금','누적거래대금','시가','고가','저가','전일종가','전일대비','전일대비구분','체결구분','등락율']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물현재가호가조회(o3106)-API용
class o3106(XAQuery):
    def Query(self, 단축코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "symbol", 0, 단축코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbolname", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", i).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "low", i).strip())
            호가수신시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "hotime", i).strip()
            매도호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho1", i).strip())
            매수호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho1", i).strip())
            매도호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt1", i).strip())
            매수호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt1", i).strip())
            매도호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem1", i).strip())
            매수호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem1", i).strip())
            매도호가2 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho2", i).strip())
            매수호가2 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho2", i).strip())
            매도호가건수2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt2", i).strip())
            매수호가건수2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt2", i).strip())
            매도호가수량2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem2", i).strip())
            매수호가수량2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem2", i).strip())
            매도호가3 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho3", i).strip())
            매수호가3 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho3", i).strip())
            매도호가건수3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt3", i).strip())
            매수호가건수3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt3", i).strip())
            매도호가수량3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem3", i).strip())
            매수호가수량3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem3", i).strip())
            매도호가4 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho4", i).strip())
            매수호가4 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho4", i).strip())
            매도호가건수4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt4", i).strip())
            매수호가건수4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt4", i).strip())
            매도호가수량4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem4", i).strip())
            매수호가수량4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem4", i).strip())
            매도호가5 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho5", i).strip())
            매수호가5 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho5", i).strip())
            매도호가건수5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt5", i).strip())
            매수호가건수5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt5", i).strip())
            매도호가수량5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem5", i).strip())
            매수호가수량5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem5", i).strip())
            매도호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt", i).strip())
            매수호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt", i).strip())
            매도호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offer", i).strip())
            매수호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bid", i).strip())

            lst = [
                종목코드, 종목명, 현재가, 전일대비구분, 전일대비, 등락율, 누적거래량, 전일종가, 시가, 고가, 저가, 호가수신시간,
                매도호가1, 매수호가1, 매도호가건수1, 매수호가건수1, 매도호가수량1, 매수호가수량1,
                매도호가2, 매수호가2, 매도호가건수2, 매수호가건수2, 매도호가수량2, 매수호가수량2,
                매도호가3, 매수호가3, 매도호가건수3, 매수호가건수3, 매도호가수량3, 매수호가수량3,
                매도호가4, 매수호가4, 매도호가건수4, 매수호가건수4, 매도호가수량4, 매수호가수량4,
                매도호가5, 매수호가5, 매도호가건수5, 매수호가건수5, 매도호가수량5, 매수호가수량5,
                매도호가건수합, 매수호가건수합, 매도호가수량합, 매수호가수량합
            ]
            result.append(lst)

        columns = ['종목코드','종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','전일종가','시가','고가','저가','호가수신시간','매도호가1','매수호가1','매도호가건수1','매수호가건수1','매도호가수량1','매수호가수량1','매도호가2','매수호가2','매도호가건수2','매수호가건수2','매도호가수량2','매수호가수량2','매도호가3','매수호가3','매도호가건수3','매수호가건수3','매도호가수량3','매수호가수량3','매도호가4','매수호가4','매도호가건수4','매수호가건수4','매도호가수량4','매수호가수량4','매도호가5','매수호가5','매도호가건수5','매수호가건수5','매도호가수량5','매수호가수량5','매도호가건수합','매수호가건수합','매도호가수량합','매수호가수량합']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물관심종목조회(o3107)-API용
class o3107(XAQuery):
    def Query(self, 종목심볼=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "symbol", 0, 단축코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbolname", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", i).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "low", i).strip())
            매도호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho1", i).strip())
            매수호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho1", i).strip())
            매도호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt1", i).strip())
            매수호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt1", i).strip())
            매도호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem1", i).strip())
            매수호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem1", i).strip())
            매도호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt", i).strip())
            매수호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt", i).strip())
            매도호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offer", i).strip())
            매수호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bid", i).strip())

            lst = [종목코드,종목명,현재가,전일대비구분,전일대비,등락율,누적거래량,전일종가,시가,고가,저가,매도호가1,매수호가1,매도호가건수1,매수호가건수1,매도호가수량1,매수호가수량1,매도호가건수합,매수호가건수합,매도호가수량합,매수호가수량합]
            result.append(lst)

        columns = ['종목코드','종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','전일종가','시가','고가','저가','매도호가1','매수호가1','매도호가건수1','매수호가건수1','매도호가수량1','매수호가수량1','매도호가건수합','매수호가건수합','매도호가수량합','매수호가수량합']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물차트(일주월)(o3108)-API용
class o3108(XAQuery):
    def Query(self, 단축코드='',주기구분='',요청건수='',시작일자='',종료일자='',연속시간='', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 주기구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0).strip()
        전일시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jisiga", 0).strip())
        전일고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jihigh", 0).strip())
        전일저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jilow", 0).strip())
        전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jiclose", 0).strip())
        전일거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jivolume", 0).strip())
        당일시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "disiga", 0).strip())
        당일고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dihigh", 0).strip())
        당일저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dilow", 0).strip())
        당일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diclose", 0).strip())
        장시작시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "mk_stime", 0).strip()
        장마감시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "mk_etime", 0).strip()
        연속일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", 0).strip()
        레코드카운트 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", 0).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드,전일시가,전일고가,전일저가,전일종가,전일거래량,당일시가,당일고가,당일저가,당일종가,장시작시간,장마감시간,연속일자,레코드카운트, df])


# 해외선물시간대별(Tick)체결(o3116)-API용
class o3116(XAQuery):
    def Query(self, 조회구분='',단축코드='',조회갯수='',순번CTS='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 조회구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "readcnt", 0, 조회갯수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 순번CTS)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 순번CTS)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        순번CTS = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_seq", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            현지일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ovsdate", i).strip()
            현지시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ovstime", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "cvolume", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [현지일자,현지시간,현재가,전일대비구분,전일대비,등락율,체결수량,누적거래량]
            result.append(lst)

        columns = ['현지일자','현지시간','현재가','전일대비구분','전일대비','등락율','체결수량','누적거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [순번CTS, df])


# 해외선물차트용NTick(o3117)-API용
class o3117(XAQuery):
    def Query(self, 단축코드='',단위='',건수='',연속시간='',연속당일구분='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK1)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_daygb", 0, 연속당일구분)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.ClearBlockdata(self.OUTBLOCK)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK1)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_daygb", 0, 연속당일구분)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0).strip()
        try:
            레코드카운트 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", 0).strip())
        except Exception as e:
            레코드카운트 = 0

        연속시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_seq", 0).strip()
        연속당일구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_daygb", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,시간,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','시간','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드,레코드카운트,연속시간,연속당일구분, df])


# 해외선물옵션마스터조회(o3121)-API용
class o3121(XAQuery):
    def Query(self, 시장구분='',옵션기초상품코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "MktGb", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "BscGdsCd", 0, 옵션기초상품코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Symbol", 0).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "SymbolNm", 0).strip()
            종목배치수신일_한국일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ApplDate", 0).strip()
            기초상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsCd", 0).strip()
            기초상품명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsNm", 0).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchCd", 0).strip()
            거래소명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchNm", 0).strip()
            기준통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "CrncyCd", 0).strip()
            진법구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "NotaCd", 0).strip()
            호가단위가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "UntPrc", 0).strip())
            최소가격변동금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MnChgAmt", 0).strip())
            가격조정계수 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "RgltFctr", 0).strip())
            계약당금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CtrtPrAmt", 0).strip())
            상품구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "GdsCd", 0).strip()
            월물_년 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngYr", 0).strip()
            월물_월 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngM", 0).strip()
            정산가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "EcPrc", 0).strip())
            거래시작시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlStrtTm", 0).strip()
            거래종료시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlEndTm", 0).strip()
            거래가능구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlPsblCd", 0).strip()
            증거금징수구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MgnCltCd", 0).strip()
            개시증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgn", 0).strip())
            유지증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgn", 0).strip())
            개시증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgnR", 0).strip())
            유지증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgnR", 0).strip())
            유효소수점자리수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "DotGb", 0).strip())
            옵션행사가 = self.ActiveX.GetFieldData(self.OUTBLOCK, "XrcPrc", 0).strip()
            기초자산기준가격 = self.ActiveX.GetFieldData(self.OUTBLOCK, "FdasBasePrc", 0).strip()
            옵션콜풋구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OptTpCode", 0).strip()
            권리행사구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "RgtXrcPtnCode", 0).strip()
            ATM구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Moneyness", 0).strip()
            해외파생기초자산종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LastSettPtnCode", 0).strip()
            해외옵션최소호가 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OptMinOrcPrc", 0).strip()
            해외옵션최소기준호가 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OptMinBaseOrcPrc", 0).strip()

            lst = [종목코드,종목명,종목배치수신일_한국일자,기초상품코드,기초상품명,거래소코드,거래소명,기준통화코드,진법구분코드,호가단위가격,최소가격변동금액,가격조정계수,계약당금액,상품구분코드,월물_년,월물_월,정산가격,거래시작시간,거래종료시간,거래가능구분코드,증거금징수구분코드,개시증거금,유지증거금,개시증거금율,유지증거금율,유효소수점자리수,옵션행사가,기초자산기준가격,옵션콜풋구분,권리행사구분코드,ATM구분,해외파생기초자산종목코드,해외옵션최소호가,해외옵션최소기준호가]
            result.append(lst)

        columns = ['종목코드','종목명','종목배치수신일_한국일자','기초상품코드','기초상품명','거래소코드','거래소명','기준통화코드','진법구분코드','호가단위가격','최소가격변동금액','가격조정계수','계약당금액','상품구분코드','월물_년','월물_월','정산가격','거래시작시간','거래종료시간','거래가능구분코드','증거금징수구분코드','개시증거금','유지증거금','개시증거금율','유지증거금율','유효소수점자리수','옵션행사가','기초자산기준가격','옵션콜풋구분','권리행사구분코드','ATM구분','해외파생기초자산종목코드','해외옵션최소호가','해외옵션최소기준호가']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물옵션차트(분)(o3123)-API용
class o3123(XAQuery):
    def Query(self, 시장구분='',단축코드='',N분주기='',조회건수='',연속일자='',연속시간='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, N분주기)
            self.ActiveX.SetFieldData(self.INBLOCK, "readcnt", 0, 조회건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0).strip()
        시차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "timediff", 0).strip())
        조회건수 = self.ActiveX.GetFieldData(self.OUTBLOCK, "readcnt", 0).strip()
        연속일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", 0).strip()
        연속시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_time", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            현지시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,현지시간,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','현지시간','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드,시차,조회건수,연속일자,연속시간, df])


# 해외선물옵션현재가(종목정보)조회(o3125)-API용
class o3125(XAQuery):
    def Query(self, 시장구분='',종목코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "symbol", 0, 종목코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Symbol", 0).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "SymbolNm", 0).strip()
            종목배치수신일 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ApplDate", 0).strip()
            기초상품코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsCd", 0).strip()
            기초상품명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "BscGdsNm", 0).strip()
            거래소코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchCd", 0).strip()
            거래소명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "ExchNm", 0).strip()
            정산구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "EcCd", 0).strip()
            기준통화코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "CrncyCd", 0).strip()
            진법구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "NotaCd", 0).strip()
            호가단위가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "UntPrc", 0).strip())
            최소가격변동금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MnChgAmt", 0).strip())
            가격조정계수 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "RgltFctr", 0).strip())
            계약당금액 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CtrtPrAmt", 0).strip())
            상장개월수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngMCnt", 0).strip())
            상품구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "GdsCd", 0).strip()
            시장구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MrktCd", 0).strip()
            Emini구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "EminiCd", 0).strip()
            상장년 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngYr", 0).strip()
            상장월 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngM", 0).strip()
            월물순서 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "SeqNo", 0).strip())
            상장일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "LstngDt", 0).strip()
            만기일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MtrtDt", 0).strip()
            최종거래일 = self.ActiveX.GetFieldData(self.OUTBLOCK, "FnlDlDt", 0).strip()
            최초인도통지일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "FstTrsfrDt", 0).strip()
            정산가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "EcPrc", 0).strip())
            거래시작일자_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlDt", 0).strip()
            거래시작시간_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlStrtTm", 0).strip()
            거래종료시간_한국 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlEndTm", 0).strip()
            거래시작일자_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsStrDay", 0).strip()
            거래시작시간_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsStrTm", 0).strip()
            거래종료일자_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsEndDay", 0).strip()
            거래종료시간_현지 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsEndTm", 0).strip()
            거래가능구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "DlPsblCd", 0).strip()
            증거금징수구분코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "MgnCltCd", 0).strip()
            개시증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgn", 0).strip())
            유지증거금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgn", 0).strip())
            개시증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpngMgnR", 0).strip())
            유지증거금율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MntncMgnR", 0).strip())
            유효소수점자리수 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "DotGb", 0).strip())
            시차 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TimeDiff", 0).strip())
            현지체결일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "OvsDate", 0).strip()
            한국체결일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "KorDate", 0).strip()
            현지체결시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdTm", 0).strip()
            한국체결시각 = self.ActiveX.GetFieldData(self.OUTBLOCK, "RcvTm", 0).strip()
            체결가격 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdP", 0).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdQ", 0).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "TotQ", 0).strip())
            체결거래대금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TrdAmt", 0).strip())
            누적거래대금 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "TotAmt", 0).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "OpenP", 0).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "HighP", 0).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "LowP", 0).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "CloseP", 0).strip())
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "YdiffP", 0).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "YdiffSign", 0).strip()
            체결구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "Cgubun", 0).strip()
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "Diff", 0).strip())
            최소호가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MinOrcPrc", 0).strip())
            최소기준호가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "MinBaseOrcPrc", 0).strip())

            lst = [
                종목코드,종목명,종목배치수신일,기초상품코드,기초상품명,거래소코드,거래소명,정산구분코드,기준통화코드,진법구분코드,호가단위가격,최소가격변동금액,가격조정계수,
                계약당금액,상장개월수,상품구분코드,시장구분코드,Emini구분코드,상장년,상장월,월물순서,상장일자,만기일자,최종거래일,최초인도통지일자,정산가격,
                거래시작일자_한국,거래시작시간_한국,거래종료시간_한국,거래시작일자_현지,거래시작시간_현지,거래종료일자_현지,거래종료시간_현지,거래가능구분코드,증거금징수구분코드,개시증거금,
                유지증거금,개시증거금율,유지증거금율,유효소수점자리수,시차,현지체결일자,한국체결일자,현지체결시간,한국체결시각,체결가격,체결수량,누적거래량,체결거래대금,누적거래대금,
                시가,고가,저가,전일종가,전일대비,전일대비구분,체결구분,등락율,최소호가,최소기준호가
            ]
            result.append(lst)

        columns = ['종목코드','종목명','종목배치수신일','기초상품코드','기초상품명','거래소코드','거래소명','정산구분코드','기준통화코드','진법구분코드','호가단위가격','최소가격변동금액','가격조정계수','계약당금액','상장개월수','상품구분코드','시장구분코드','Emini구분코드','상장년','상장월','월물순서','상장일자','만기일자','최종거래일','최초인도통지일자','정산가격','거래시작일자_한국','거래시작시간_한국','거래종료시간_한국','거래시작일자_현지','거래시작시간_현지','거래종료일자_현지','거래종료시간_현지','거래가능구분코드','증거금징수구분코드','개시증거금','유지증거금','개시증거금율','유지증거금율','유효소수점자리수','시차','현지체결일자','한국체결일자','현지체결시간','한국체결시각','체결가격','체결수량','누적거래량','체결거래대금','누적거래대금','시가','고가','저가','전일종가','전일대비','전일대비구분','체결구분','등락율','최소호가','최소기준호가']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물옵션현재가호가조회(o3126)-API용
class o3126(XAQuery):
    def Query(self, 단축코드=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK, "symbol", 0, 단축코드)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbolname", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", i).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "low", i).strip())
            호가수신시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "hotime", i).strip()
            매도호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho1", i).strip())
            매수호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho1", i).strip())
            매도호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt1", i).strip())
            매수호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt1", i).strip())
            매도호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem1", i).strip())
            매수호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem1", i).strip())
            매도호가2 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho2", i).strip())
            매수호가2 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho2", i).strip())
            매도호가건수2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt2", i).strip())
            매수호가건수2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt2", i).strip())
            매도호가수량2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem2", i).strip())
            매수호가수량2 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem2", i).strip())
            매도호가3 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho3", i).strip())
            매수호가3 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho3", i).strip())
            매도호가건수3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt3", i).strip())
            매수호가건수3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt3", i).strip())
            매도호가수량3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem3", i).strip())
            매수호가수량3 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem3", i).strip())
            매도호가4 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho4", i).strip())
            매수호가4 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho4", i).strip())
            매도호가건수4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt4", i).strip())
            매수호가건수4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt4", i).strip())
            매도호가수량4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem4", i).strip())
            매수호가수량4 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem4", i).strip())
            매도호가5 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho5", i).strip())
            매수호가5 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho5", i).strip())
            매도호가건수5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt5", i).strip())
            매수호가건수5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt5", i).strip())
            매도호가수량5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem5", i).strip())
            매수호가수량5 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem5", i).strip())
            매도호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt", i).strip())
            매수호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt", i).strip())
            매도호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offer", i).strip())
            매수호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bid", i).strip())

            lst = [
                종목코드, 종목명, 현재가, 전일대비구분, 전일대비, 등락율, 누적거래량, 전일종가, 시가, 고가, 저가, 호가수신시간,
                매도호가1, 매수호가1, 매도호가건수1, 매수호가건수1, 매도호가수량1, 매수호가수량1,
                매도호가2, 매수호가2, 매도호가건수2, 매수호가건수2, 매도호가수량2, 매수호가수량2,
                매도호가3, 매수호가3, 매도호가건수3, 매수호가건수3, 매도호가수량3, 매수호가수량3,
                매도호가4, 매수호가4, 매도호가건수4, 매수호가건수4, 매도호가수량4, 매수호가수량4,
                매도호가5, 매수호가5, 매도호가건수5, 매수호가건수5, 매도호가수량5, 매수호가수량5,
                매도호가건수합, 매수호가건수합, 매도호가수량합, 매수호가수량합
            ]
            result.append(lst)

        columns = ['종목코드','종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','전일종가','시가','고가','저가','호가수신시간','매도호가1','매수호가1','매도호가건수1','매수호가건수1','매도호가수량1','매수호가수량1','매도호가2','매수호가2','매도호가건수2','매수호가건수2','매도호가수량2','매수호가수량2','매도호가3','매수호가3','매도호가건수3','매수호가건수3','매도호가수량3','매수호가수량3','매도호가4','매수호가4','매도호가건수4','매수호가건수4','매도호가수량4','매수호가수량4','매도호가5','매수호가5','매도호가건수5','매수호가건수5','매도호가수량5','매수호가수량5','매도호가건수합','매수호가건수합','매도호가수량합','매수호가수량합']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물옵션관심종목조회(o3127)-API용
class o3127(XAQuery):
    def Query(self, 건수='',시장구분='',종목심볼=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.SetFieldData(self.INBLOCK, "nrec", 0, 건수)

        self.ActiveX.SetFieldData(self.INBLOCK1, "mktgb", 0, 시장구분)
        self.ActiveX.SetFieldData(self.INBLOCK1, "symbol", 0, 종목심볼)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            종목코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbol", i).strip()
            종목명 = self.ActiveX.GetFieldData(self.OUTBLOCK, "symbolname", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diff", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "volume", i).strip())
            전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jnilclose", i).strip())
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "low", i).strip())
            매도호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerho1", i).strip())
            매수호가1 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidho1", i).strip())
            매도호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt1", i).strip())
            매수호가건수1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt1", i).strip())
            매도호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offerrem1", i).strip())
            매수호가수량1 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidrem1", i).strip())
            매도호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offercnt", i).strip())
            매수호가건수합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bidcnt", i).strip())
            매도호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "offer", i).strip())
            매수호가수량합 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "bid", i).strip())

            lst = [종목코드,종목명,현재가,전일대비구분,전일대비,등락율,누적거래량,전일종가,시가,고가,저가,매도호가1,매수호가1,매도호가건수1,매수호가건수1,매도호가수량1,매수호가수량1,매도호가건수합,매수호가건수합,매도호가수량합,매수호가수량합]
            result.append(lst)

        columns = ['종목코드','종목명','현재가','전일대비구분','전일대비','등락율','누적거래량','전일종가','시가','고가','저가','매도호가1','매수호가1','매도호가건수1','매수호가건수1','매도호가수량1','매수호가수량1','매도호가건수합','매수호가건수합','매도호가수량합','매수호가수량합']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [df])


# 해외선물옵션차트일주월(o3128)-API용
class o3128(XAQuery):
    def Query(self, 시장구분='',단축코드='',주기구분='',요청건수='',시작일자='',종료일자='',연속일자='', 연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 주기구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "sdate", 0, 시작일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "edate", 0, 종료일자)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_date", 0, 연속일자)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0).strip()
        전일시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jisiga", 0).strip())
        전일고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jihigh", 0).strip())
        전일저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jilow", 0).strip())
        전일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "jiclose", 0).strip())
        전일거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "jivolume", 0).strip())
        당일시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "disiga", 0).strip())
        당일고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dihigh", 0).strip())
        당일저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "dilow", 0).strip())
        당일종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK, "diclose", 0).strip())
        장시작시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "mk_stime", 0).strip()
        장마감시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "mk_etime", 0).strip()
        연속일자 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_date", 0).strip()
        레코드카운트 = int(self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", 0).strip())

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드,전일시가,전일고가,전일저가,전일종가,전일거래량,당일시가,당일고가,당일저가,당일종가,장시작시간,장마감시간,연속일자,레코드카운트, df])


# 해외선물옵션시간대별(Tick)체결(o3136)-API용
class o3136(XAQuery):
    def Query(self, 조회구분='',시장구분='',단축코드='',조회갯수='',순번CTS='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.SetFieldData(self.INBLOCK, "gubun", 0, 조회구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "readcnt", 0, 조회갯수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 순번CTS)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 순번CTS)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        순번CTS = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_seq", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            현지일자 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ovsdate", i).strip()
            현지시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "ovstime", i).strip()
            현재가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "price", i).strip())
            전일대비구분 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            체결수량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "cvolume", i).strip())
            누적거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [현지일자,현지시간,현재가,전일대비구분,전일대비,등락율,체결수량,누적거래량]
            result.append(lst)

        columns = ['현지일자','현지시간','현재가','전일대비구분','전일대비','등락율','체결수량','누적거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [순번CTS, df])


# 해외선물옵션차트용NTick(o3137)-API용
class o3137(XAQuery):
    def Query(self, 시장구분='F',단축코드='',단위='',건수='',연속시간='',연속당일구분='',연속조회=False):
        if 연속조회 == False:
            self.ActiveX.LoadFromResFile(self.RESFILE)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK1)
            self.ActiveX.SetFieldData(self.INBLOCK, "mktgb", 0, 시장구분)
            self.ActiveX.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.ActiveX.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.ActiveX.SetFieldData(self.INBLOCK, "qrycnt", 0, 건수)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_daygb", 0, 연속당일구분)
            self.ActiveX.Request(0)
        else:
            self.ActiveX.ClearBlockdata(self.OUTBLOCK)
            self.ActiveX.ClearBlockdata(self.OUTBLOCK1)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_seq", 0, 연속시간)
            self.ActiveX.SetFieldData(self.INBLOCK, "cts_daygb", 0, 연속당일구분)
            err_code = self.ActiveX.Request(True)  # 연속조회인경우만 True
            if err_code < 0:
                클래스이름 = self.__class__.__name__
                함수이름 = inspect.currentframe().f_code.co_name
                print("%s-%s " % (클래스이름, 함수이름), "error... {0}".format(err_code))

    def OnReceiveData(self, szTrCode):
        단축코드 = self.ActiveX.GetFieldData(self.OUTBLOCK, "shcode", 0).strip()
        레코드카운트 = self.ActiveX.GetFieldData(self.OUTBLOCK, "rec_count", 0).strip()
        연속시간 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_seq", 0).strip()
        연속당일구분 = self.ActiveX.GetFieldData(self.OUTBLOCK, "cts_daygb", 0).strip()

        result = []
        nCount = self.ActiveX.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            날짜 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            시간 = self.ActiveX.GetFieldData(self.OUTBLOCK1, "time", i).strip()
            시가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "open", i).strip())
            고가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            저가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            종가 = float(self.ActiveX.GetFieldData(self.OUTBLOCK1, "close", i).strip())
            거래량 = int(self.ActiveX.GetFieldData(self.OUTBLOCK1, "volume", i).strip())

            lst = [날짜,시간,시가,고가,저가,종가,거래량]
            result.append(lst)

        columns = ['날짜','시간','시가','고가','저가','종가','거래량']
        df = DataFrame(data=result, columns=columns)

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [단축코드,레코드카운트,연속시간,연속당일구분, df])


# 시간조회
class t0167(XAQuery):
    def Query(self, id=''):
        self.ActiveX.LoadFromResFile(self.RESFILE)
        self.ActiveX.ClearBlockdata(self.OUTBLOCK)
        self.ActiveX.SetFieldData(self.INBLOCK, "id", 0, id)
        self.ActiveX.Request(0)

    def OnReceiveData(self, szTrCode):
        dt = self.ActiveX.GetFieldData(self.OUTBLOCK, "dt", 0).strip()
        time = self.ActiveX.GetFieldData(self.OUTBLOCK, "time", 0).strip()

        if self.parent != None:
            self.parent.OnReceiveData(szTrCode, [dt, time])
