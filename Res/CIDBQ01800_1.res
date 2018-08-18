BEGIN_FUNCTION_MAP
	.Func,해외선물 주문체결내역 조회,CIDBQ01800,SERVICE=CIDBQ01800,headtype=B,CREATOR=이호섭,CREDATE=2015/06/22 19:06:11;
	BEGIN_DATA_MAP
	CIDBQ01800InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		주문일자, OrdDt, OrdDt, char, 8;
		당일구분코드, ThdayTpCode, ThdayTpCode, char, 1;
		주문상태코드, OrdStatCode, OrdStatCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		조회구분코드, QryTpCode, QryTpCode, char, 1;
		주문유형코드, OrdPtnCode, OrdPtnCode, char, 2;
		해외파생선물옵션구분코드, OvrsDrvtFnoTpCode, OvrsDrvtFnoTpCode, char, 1;
	end
	CIDBQ01800OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		주문일자, OrdDt, OrdDt, char, 8;
		당일구분코드, ThdayTpCode, ThdayTpCode, char, 1;
		주문상태코드, OrdStatCode, OrdStatCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		조회구분코드, QryTpCode, QryTpCode, char, 1;
		주문유형코드, OrdPtnCode, OrdPtnCode, char, 2;
		해외파생선물옵션구분코드, OvrsDrvtFnoTpCode, OvrsDrvtFnoTpCode, char, 1;
	end
	CIDBQ01800OutBlock2,Out(*EMPTY*),output,occurs;
	begin
		해외선물주문번호, OvrsFutsOrdNo, OvrsFutsOrdNo, char, 10;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		FCM주문번호, FcmOrdNo, FcmOrdNo, char, 15;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		종목명, IsuNm, IsuNm, char, 50;
		해외선물행사가격, AbrdFutsXrcPrc, AbrdFutsXrcPrc, double, 29.10;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		매매구분명, BnsTpNm, BnsTpNm, char, 10;
		선물주문상태코드, FutsOrdStatCode, FutsOrdStatCode, char, 1;
		구분코드명, TpCodeNm, TpCodeNm, char, 50;
		선물주문구분코드, FutsOrdTpCode, FutsOrdTpCode, char, 1;
		거래구분명, TrdTpNm, TrdTpNm, char, 20;
		해외선물주문유형코드, AbrdFutsOrdPtnCode, AbrdFutsOrdPtnCode, char, 1;
		주문유형명, OrdPtnNm, OrdPtnNm, char, 40;
		주문유형기간구분코드, OrdPtnTermTpCode, OrdPtnTermTpCode, char, 2;
		공통코드명, CmnCodeNm, CmnCodeNm, char, 100;
		적용시작일자, AppSrtDt, AppSrtDt, char, 8;
		적용종료일자, AppEndDt, AppEndDt, char, 8;
		해외파생주문가격, OvrsDrvtOrdPrc, OvrsDrvtOrdPrc, double, 25.8;
		주문수량, OrdQty, OrdQty, long, 16;
		해외선물체결가격, AbrdFutsExecPrc, AbrdFutsExecPrc, double, 29.10;
		체결수량, ExecQty, ExecQty, long, 16;
		주문조건가격, OrdCndiPrc, OrdCndiPrc, double, 25.8;
		해외파생현재가, OvrsDrvtNowPrc, OvrsDrvtNowPrc, double, 25.8;
		정정수량, MdfyQty, MdfyQty, long, 16;
		취소수량, CancQty, CancQty, long, 16;
		거부수량, RjtQty, RjtQty, long, 13;
		확인수량, CnfQty, CnfQty, long, 16;
		반대매매여부, CvrgYn, CvrgYn, char, 1;
		등록단말번호, RegTmnlNo, RegTmnlNo, char, 3;
		등록지점번호, RegBrnNo, RegBrnNo, char, 3;
		등록사용자ID, RegUserId, RegUserId, char, 16;
		주문일자, OrdDt, OrdDt, char, 8;
		주문시각, OrdTime, OrdTime, char, 9;
		해외옵션행사예약구분코드, OvrsOptXrcRsvTpCode, OvrsOptXrcRsvTpCode, char, 1;
	end
	END_DATA_MAP
END_FUNCTION_MAP
