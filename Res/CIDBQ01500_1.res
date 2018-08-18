BEGIN_FUNCTION_MAP
	.Func,해외선물 미결제 잔고내역,CIDBQ01500,SERVICE=CIDBQ01500,ENCRYPT,headtype=B,CREATOR=이호섭,CREDATE=2013/04/11 18:36:15;
	BEGIN_DATA_MAP
	CIDBQ01500InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌구분코드, AcntTpCode, AcntTpCode, char, 1;
		계좌번호, AcntNo, AcntNo, char, 20;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		조회일자, QryDt, QryDt, char, 8;
		잔고구분코드, BalTpCode, BalTpCode, char, 1;
	end
	CIDBQ01500OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌구분코드, AcntTpCode, AcntTpCode, char, 1;
		계좌번호, AcntNo, AcntNo, char, 20;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		조회일자, QryDt, QryDt, char, 8;
		잔고구분코드, BalTpCode, BalTpCode, char, 1;
	end
	CIDBQ01500OutBlock2,Out(*EMPTY*),output,occurs;
	begin
		기준일자, BaseDt, BaseDt, char, 8;
		예수금, Dps, Dps, long, 16;
		청산손익금액, LpnlAmt, LpnlAmt, double, 19.2;
		선물만기전청산손익금액, FutsDueBfLpnlAmt, FutsDueBfLpnlAmt, double, 23.2;
		선물만기전수수료, FutsDueBfCmsn, FutsDueBfCmsn, double, 23.2;
		위탁증거금액, CsgnMgn, CsgnMgn, long, 16;
		유지증거금, MaintMgn, MaintMgn, long, 16;
		신용한도금액, CtlmtAmt, CtlmtAmt, double, 23.2;
		추가증거금액, AddMgn, AddMgn, long, 16;
		마진콜율, MgnclRat, MgnclRat, double, 27.10;
		주문가능금액, OrdAbleAmt, OrdAbleAmt, long, 16;
		인출가능금액, WthdwAbleAmt, WthdwAbleAmt, long, 16;
		계좌번호, AcntNo, AcntNo, char, 20;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		종목명, IsuNm, IsuNm, char, 50;
		통화코드값, CrcyCodeVal, CrcyCodeVal, char, 3;
		해외파생상품코드, OvrsDrvtPrdtCode, OvrsDrvtPrdtCode, char, 10;
		해외파생옵션구분코드, OvrsDrvtOptTpCode, OvrsDrvtOptTpCode, char, 1;
		만기일자, DueDt, DueDt, char, 8;
		해외파생행사가격, OvrsDrvtXrcPrc, OvrsDrvtXrcPrc, double, 25.8;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		공통코드명, CmnCodeNm, CmnCodeNm, char, 100;
		구분코드명, TpCodeNm, TpCodeNm, char, 50;
		잔고수량, BalQty, BalQty, long, 16;
		매입가격, PchsPrc, PchsPrc, double, 25.8;
		해외파생현재가, OvrsDrvtNowPrc, OvrsDrvtNowPrc, double, 25.8;
		해외선물평가손익금액, AbrdFutsEvalPnlAmt, AbrdFutsEvalPnlAmt, double, 19.2;
		위탁수수료, CsgnCmsn, CsgnCmsn, double, 19.2;
		포지션번호, PosNo, PosNo, char, 13;
		거래소비용1수수료금액, EufOneCmsnAmt, EufOneCmsnAmt, double, 19.2;
		거래소비용2수수료금액, EufTwoCmsnAmt, EufTwoCmsnAmt, double, 19.2;
	end
	END_DATA_MAP
END_FUNCTION_MAP
