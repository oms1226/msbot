BEGIN_FUNCTION_MAP
	.Func,해외선물 계좌예탁자산조회,CIDBQ05300,SERVICE=CIDBQ05300,headtype=B,CREATOR=이호섭,CREDATE=2015/06/22 19:42:57;
	BEGIN_DATA_MAP
	CIDBQ05300InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		해외계좌구분코드, OvrsAcntTpCode, OvrsAcntTpCode, char, 1;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		계좌번호, AcntNo, AcntNo, char, 20;
		계좌비밀번호, AcntPwd, AcntPwd, char, 8;
		통화코드, CrcyCode, CrcyCode, char, 3;
	end
	CIDBQ05300OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		해외계좌구분코드, OvrsAcntTpCode, OvrsAcntTpCode, char, 1;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		계좌번호, AcntNo, AcntNo, char, 20;
		계좌비밀번호, AcntPwd, AcntPwd, char, 8;
		통화코드, CrcyCode, CrcyCode, char, 3;
	end
	CIDBQ05300OutBlock2,Out(*EMPTY*),output,occurs;
	begin
		계좌번호, AcntNo, AcntNo, char, 20;
		통화코드, CrcyCode, CrcyCode, char, 3;
		해외선물예수금, OvrsFutsDps, OvrsFutsDps, double, 23.2;
		해외선물위탁증거금액, AbrdFutsCsgnMgn, AbrdFutsCsgnMgn, double, 19.2;
		해외선물추가증거금, OvrsFutsSplmMgn, OvrsFutsSplmMgn, double, 23.2;
		고객청산손익금액, CustmLpnlAmt, CustmLpnlAmt, double, 19.2;
		해외선물평가손익금액, AbrdFutsEvalPnlAmt, AbrdFutsEvalPnlAmt, double, 19.2;
		해외선물수수료금액, AbrdFutsCmsnAmt, AbrdFutsCmsnAmt, double, 19.2;
		해외선물평가예탁총금액, AbrdFutsEvalDpstgTotAmt, AbrdFutsEvalDpstgTotAmt, double, 19.2;
		환율, Xchrat, Xchrat, double, 15.4;
		외화실환전금액, FcurrRealMxchgAmt, FcurrRealMxchgAmt, double, 19.2;
		해외선물인출가능금액, AbrdFutsWthdwAbleAmt, AbrdFutsWthdwAbleAmt, double, 19.2;
		해외선물주문가능금액, AbrdFutsOrdAbleAmt, AbrdFutsOrdAbleAmt, double, 19.2;
		선물만기미도래청산손익금액, FutsDueNarrvLqdtPnlAmt, FutsDueNarrvLqdtPnlAmt, double, 19.2;
		선물만기미도래수수료, FutsDueNarrvCmsn, FutsDueNarrvCmsn, double, 19.2;
		해외선물청산손익금액, AbrdFutsLqdtPnlAmt, AbrdFutsLqdtPnlAmt, double, 19.2;
		해외선물만기수수료, OvrsFutsDueCmsn, OvrsFutsDueCmsn, double, 19.2;
		해외선물옵션매수금액, OvrsFutsOptBuyAmt, OvrsFutsOptBuyAmt, double, 23.2;
		해외선물옵션매도금액, OvrsFutsOptSellAmt, OvrsFutsOptSellAmt, double, 23.2;
		옵션매수시장가치금액, OptBuyMktWrthAmt, OptBuyMktWrthAmt, double, 19.2;
		옵션매도시장가치금액, OptSellMktWrthAmt, OptSellMktWrthAmt, double, 19.2;
	end
	CIDBQ05300OutBlock3,SelOut(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		해외선물예수금, OvrsFutsDps, OvrsFutsDps, double, 23.2;
		해외선물청산손익금액, AbrdFutsLqdtPnlAmt, AbrdFutsLqdtPnlAmt, double, 19.2;
		선물만기미도래청산손익금액, FutsDueNarrvLqdtPnlAmt, FutsDueNarrvLqdtPnlAmt, double, 19.2;
		해외선물평가손익금액, AbrdFutsEvalPnlAmt, AbrdFutsEvalPnlAmt, double, 19.2;
		해외선물평가예탁총금액, AbrdFutsEvalDpstgTotAmt, AbrdFutsEvalDpstgTotAmt, double, 19.2;
		고객청산손익금액, CustmLpnlAmt, CustmLpnlAmt, double, 19.2;
		해외선물만기수수료, OvrsFutsDueCmsn, OvrsFutsDueCmsn, double, 19.2;
		외화실환전금액, FcurrRealMxchgAmt, FcurrRealMxchgAmt, double, 19.2;
		해외선물수수료금액, AbrdFutsCmsnAmt, AbrdFutsCmsnAmt, double, 19.2;
		선물만기미도래수수료, FutsDueNarrvCmsn, FutsDueNarrvCmsn, double, 19.2;
		해외선물위탁증거금액, AbrdFutsCsgnMgn, AbrdFutsCsgnMgn, double, 19.2;
		해외선물유지증거금, OvrsFutsMaintMgn, OvrsFutsMaintMgn, double, 19.2;
		해외선물옵션매수금액, OvrsFutsOptBuyAmt, OvrsFutsOptBuyAmt, double, 23.2;
		해외선물옵션매도금액, OvrsFutsOptSellAmt, OvrsFutsOptSellAmt, double, 23.2;
		신용한도금액, CtlmtAmt, CtlmtAmt, double, 23.2;
		해외선물추가증거금, OvrsFutsSplmMgn, OvrsFutsSplmMgn, double, 23.2;
		마진콜율, MgnclRat, MgnclRat, double, 27.10;
		해외선물주문가능금액, AbrdFutsOrdAbleAmt, AbrdFutsOrdAbleAmt, double, 19.2;
		해외선물인출가능금액, AbrdFutsWthdwAbleAmt, AbrdFutsWthdwAbleAmt, double, 19.2;
		옵션매수시장가치금액, OptBuyMktWrthAmt, OptBuyMktWrthAmt, double, 19.2;
		옵션매도시장가치금액, OptSellMktWrthAmt, OptSellMktWrthAmt, double, 19.2;
		해외옵션결제금액, OvrsOptSettAmt, OvrsOptSettAmt, double, 19.2;
		해외옵션잔고평가금액, OvrsOptBalEvalAmt, OvrsOptBalEvalAmt, double, 19.2;
	end
	END_DATA_MAP
END_FUNCTION_MAP
