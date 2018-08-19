# -*- coding: utf-8 -*-

import sys, os
import datetime, time
import win32com.client
import pythoncom
import inspect

class XASessionEvents(object):
    """
    https://www.youtube.com/watch?v=XAfq6Des0bI&list=PLKuN_z83AuF0I0dRIJmNAsFvi1DGiLTm9&index=13
    """
    def __init__(self):
        self.parent = None

    def set_parent(self, parent):
        self.parent = parent

    def OnLogin(self, code, msg):
        if self.parent != None:
            self.parent.OnLogin(code, msg)

    def OnLogout(self):
        if self.parent != None:
            self.parent.OnLogout()

    def OnDisconnect(self):
        if self.parent != None:
            self.parent.OnDisconnect()


class XASession:
    def __init__(self, parent=None):
        self.ActiveX = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
        self.ActiveX.SetMode("_XINGAPI7_","TRUE")
        if parent == None:
            self.ActiveX.set_parent(parent=self)
        else:
            self.ActiveX.set_parent(parent=parent)

    def login(self, url='demo.ebestsec.co.kr', port=200001, svrtype=0, id='userid', pwd='password', cert='공인인증 비밀번호'):
        """
        https://www.youtube.com/watch?v=XAfq6Des0bI&list=PLKuN_z83AuF0I0dRIJmNAsFvi1DGiLTm9&index=13 //[13강] 파이썬 기본 강의 -xingAPI를 이용한 로그인1
        https://www.youtube.com/watch?v=Qtd-zUnZZbk&list=PLKuN_z83AuF0I0dRIJmNAsFvi1DGiLTm9&index=12 //[13강] 파이썬 기본 강의 -xingAPI를 이용한 로그인2
        2 xingAPI COM 개발가이드.pdf
        """
        result = self.ActiveX.ConnectServer(url, port)
        if not result:
            nErrCode = self.ActiveX.GetLastError()
            strErrMsg = self.ActiveX.GetErrorMessage(nErrCode)
            return (False, nErrCode, strErrMsg)

        self.ActiveX.Login(id, pwd, cert, svrtype, 0)

        return (True, 0, "OK")

    def logout(self):
        self.ActiveX.Logout()

    def disconnect(self):
        self.ActiveX.DisconnectServer()

    def IsConnected(self):
        return self.ActiveX.IsConnected()
