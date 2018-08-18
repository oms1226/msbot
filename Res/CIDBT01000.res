BEGIN_FUNCTION_MAP
	.Func,해외선물취소주문,CIDBT01000,SERVICE=CIDBT01000,ENCRYPT,SIGNATURE,headtype=B,CREATOR=최영호,CREDATE=2012/04/26 14:52:30;
	BEGIN_DATA_MAP
	CIDBT01000InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		주문일자, OrdDt, OrdDt, char, 8;
		지점번호, BrnNo, BrnNo, char, 3;
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		선물주문구분코드, FutsOrdTpCode, FutsOrdTpCode, char, 1;
		상품구분코드, PrdtTpCode, PrdtTpCode, char, 2;
		거래소코드, ExchCode, ExchCode, char, 10;
	end
	CIDBT01000OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		주문일자, OrdDt, OrdDt, char, 8;
		지점번호, BrnNo, BrnNo, char, 3;
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		선물주문구분코드, FutsOrdTpCode, FutsOrdTpCode, char, 1;
		상품구분코드, PrdtTpCode, PrdtTpCode, char, 2;
		거래소코드, ExchCode, ExchCode, char, 10;
	end
	CIDBT01000OutBlock2,Out(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		해외선물주문번호, OvrsFutsOrdNo, OvrsFutsOrdNo, char, 10;
		내부메시지내용, InnerMsgCnts, InnerMsgCnts, char, 80;
	end
	END_DATA_MAP
END_FUNCTION_MAP
