BEGIN_FUNCTION_MAP
	.Func,해외선물 예수금/잔고현황,CIDBQ03000,SERVICE=CIDBQ03000,headtype=B,CREATOR=이호섭,CREDATE=2015/06/25 09:12:31;
	BEGIN_DATA_MAP
	CIDBQ03000InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌구분코드, AcntTpCode, AcntTpCode, char, 1;
		계좌번호, AcntNo, AcntNo, char, 20;
		계좌비밀번호, AcntPwd, AcntPwd, char, 8;
		거래일자, TrdDt, TrdDt, char, 8;
	end
	CIDBQ03000OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌구분코드, AcntTpCode, AcntTpCode, char, 1;
		계좌번호, AcntNo, AcntNo, char, 20;
		계좌비밀번호, AcntPwd, AcntPwd, char, 8;
		거래일자, TrdDt, TrdDt, char, 8;
	end
	CIDBQ03000OutBlock2,Out(*EMPTY*),output,occurs;
	begin
		계좌번호, AcntNo, AcntNo, char, 20;
		거래일자, TrdDt, TrdDt, char, 8;
		통화대상코드, CrcyObjCode, CrcyObjCode, char, 12;
		해외선물예수금, OvrsFutsDps, OvrsFutsDps, double, 23.2;
		고객입출금금액, CustmMnyioAmt, CustmMnyioAmt, double, 19.2;
		해외선물청산손익금액, AbrdFutsLqdtPnlAmt, AbrdFutsLqdtPnlAmt, double, 19.2;
		해외선물수수료금액, AbrdFutsCmsnAmt, AbrdFutsCmsnAmt, double, 19.2;
		가환전예수금, PrexchDps, PrexchDps, double, 19.2;
		평가자산금액, EvalAssetAmt, EvalAssetAmt, double, 19.2;
		해외선물위탁증거금액, AbrdFutsCsgnMgn, AbrdFutsCsgnMgn, double, 19.2;
		해외선물추가증거금액, AbrdFutsAddMgn, AbrdFutsAddMgn, double, 19.2;
		해외선물인출가능금액, AbrdFutsWthdwAbleAmt, AbrdFutsWthdwAbleAmt, double, 19.2;
		해외선물주문가능금액, AbrdFutsOrdAbleAmt, AbrdFutsOrdAbleAmt, double, 19.2;
		해외선물평가손익금액, AbrdFutsEvalPnlAmt, AbrdFutsEvalPnlAmt, double, 19.2;
		최종결제손익금액, LastSettPnlAmt, LastSettPnlAmt, double, 19.2;
		해외옵션결제금액, OvrsOptSettAmt, OvrsOptSettAmt, double, 19.2;
		해외옵션잔고평가금액, OvrsOptBalEvalAmt, OvrsOptBalEvalAmt, double, 19.2;
	end
	END_DATA_MAP
END_FUNCTION_MAP
