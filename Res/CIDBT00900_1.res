BEGIN_FUNCTION_MAP
	.Func,해외선물정정주문,CIDBT00900,SERVICE=CIDBT00900,ENCRYPT,SIGNATURE,headtype=B,CREATOR=김재홍,CREDATE=2012/02/23 15:00:10;
	BEGIN_DATA_MAP
	CIDBT00900InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		주문일자, OrdDt, OrdDt, char, 8;
		등록지점번호, RegBrnNo, RegBrnNo, char, 3;
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		선물주문구분코드, FutsOrdTpCode, FutsOrdTpCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		선물주문유형코드, FutsOrdPtnCode, FutsOrdPtnCode, char, 1;
		통화코드값, CrcyCodeVal, CrcyCodeVal, char, 3;
		해외파생주문가격, OvrsDrvtOrdPrc, OvrsDrvtOrdPrc, double, 25.8;
		조건주문가격, CndiOrdPrc, CndiOrdPrc, double, 25.8;
		주문수량, OrdQty, OrdQty, long, 16;
		해외파생상품코드, OvrsDrvtPrdtCode, OvrsDrvtPrdtCode, char, 10;
		만기년월, DueYymm, DueYymm, char, 6;
		거래소코드, ExchCode, ExchCode, char, 10;
	end
	CIDBT00900OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		주문일자, OrdDt, OrdDt, char, 8;
		등록지점번호, RegBrnNo, RegBrnNo, char, 3;
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		선물주문구분코드, FutsOrdTpCode, FutsOrdTpCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		선물주문유형코드, FutsOrdPtnCode, FutsOrdPtnCode, char, 1;
		통화코드값, CrcyCodeVal, CrcyCodeVal, char, 3;
		해외파생주문가격, OvrsDrvtOrdPrc, OvrsDrvtOrdPrc, double, 25.8;
		조건주문가격, CndiOrdPrc, CndiOrdPrc, double, 25.8;
		주문수량, OrdQty, OrdQty, long, 16;
		해외파생상품코드, OvrsDrvtPrdtCode, OvrsDrvtPrdtCode, char, 10;
		만기년월, DueYymm, DueYymm, char, 6;
		거래소코드, ExchCode, ExchCode, char, 10;
	end
	CIDBT00900OutBlock2,Out(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		해외선물주문번호, OvrsFutsOrdNo, OvrsFutsOrdNo, char, 10;
		내부메시지내용, InnerMsgCnts, InnerMsgCnts, char, 80;
	end
	END_DATA_MAP
END_FUNCTION_MAP
