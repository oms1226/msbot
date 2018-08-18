BEGIN_FUNCTION_MAP
	.Func,해외선물 주문체결내역 상세 조회,CIDBQ02400,SERVICE=CIDBQ02400,ENCRYPT,headtype=B,CREATOR=이호섭,CREDATE=2015/06/18 18:34:41;
	BEGIN_DATA_MAP
	CIDBQ02400InBlock1,In(*EMPTY*),input;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		조회시작일자, QrySrtDt, QrySrtDt, char, 8;
		조회종료일자, QryEndDt, QryEndDt, char, 8;
		당일구분코드, ThdayTpCode, ThdayTpCode, char, 1;
		주문상태코드, OrdStatCode, OrdStatCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		조회구분코드, QryTpCode, QryTpCode, char, 1;
		주문유형코드, OrdPtnCode, OrdPtnCode, char, 2;
		해외파생선물옵션구분코드, OvrsDrvtFnoTpCode, OvrsDrvtFnoTpCode, char, 1;
	end
	CIDBQ02400OutBlock1,In(*EMPTY*),output;
	begin
		레코드갯수, RecCnt, RecCnt, long, 5
		계좌번호, AcntNo, AcntNo, char, 20;
		비밀번호, Pwd, Pwd, char, 8;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		조회시작일자, QrySrtDt, QrySrtDt, char, 8;
		조회종료일자, QryEndDt, QryEndDt, char, 8;
		당일구분코드, ThdayTpCode, ThdayTpCode, char, 1;
		주문상태코드, OrdStatCode, OrdStatCode, char, 1;
		매매구분코드, BnsTpCode, BnsTpCode, char, 1;
		조회구분코드, QryTpCode, QryTpCode, char, 1;
		주문유형코드, OrdPtnCode, OrdPtnCode, char, 2;
		해외파생선물옵션구분코드, OvrsDrvtFnoTpCode, OvrsDrvtFnoTpCode, char, 1;
	end
	CIDBQ02400OutBlock2,Out(*EMPTY*),output,occurs;
	begin
		주문일자, OrdDt, OrdDt, char, 8;
		해외선물주문번호, OvrsFutsOrdNo, OvrsFutsOrdNo, char, 10;
		해외선물원주문번호, OvrsFutsOrgOrdNo, OvrsFutsOrgOrdNo, char, 10;
		FCM주문번호, FcmOrdNo, FcmOrdNo, char, 15;
		해외선물체결번호, OvrsFutsExecNo, OvrsFutsExecNo, char, 10;
		FCM계좌번호, FcmAcntNo, FcmAcntNo, char, 20;
		종목코드값, IsuCodeVal, IsuCodeVal, char, 18;
		종목명, IsuNm, IsuNm, char, 50;
		해외선물행사가격, AbrdFutsXrcPrc, AbrdFutsXrcPrc, double, 29.10;
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
		주문수량, OrdQty, OrdQty, long, 16;
		해외파생주문가격, OvrsDrvtOrdPrc, OvrsDrvtOrdPrc, double, 25.8;
		체결수량, ExecQty, ExecQty, long, 16;
		해외선물체결가격, AbrdFutsExecPrc, AbrdFutsExecPrc, double, 29.10;
		주문조건가격, OrdCndiPrc, OrdCndiPrc, double, 25.8;
		현재가, NowPrc, NowPrc, double, 15.2;
		처리상태코드, TrxStatCode, TrxStatCode, char, 2;
		처리상태코드명, TrxStatCodeNm, TrxStatCodeNm, char, 40;
		위탁수수료, CsgnCmsn, CsgnCmsn, double, 19.2;
		FCM수수료, FcmCmsn, FcmCmsn, double, 21.4;
		당사수수료, ThcoCmsn, ThcoCmsn, double, 19.2;
		매체코드, MdaCode, MdaCode, char, 2;
		매체코드명, MdaCodeNm, MdaCodeNm, char, 40;
		등록단말번호, RegTmnlNo, RegTmnlNo, char, 3;
		등록사용자ID, RegUserId, RegUserId, char, 16;
		주문일시, OrdDttm, OrdDttm, char, 30;
		주문시각, OrdTime, OrdTime, char, 9;
		체결일자, ExecDt, ExecDt, char, 8;
		체결시각, ExecTime, ExecTime, char, 9;
		거래소비용1수수료금액, EufOneCmsnAmt, EufOneCmsnAmt, double, 19.2;
		거래소비용2수수료금액, EufTwoCmsnAmt, EufTwoCmsnAmt, double, 19.2;
		런던청산소1수수료금액, LchOneCmsnAmt, LchOneCmsnAmt, double, 19.2;
		런던청산소2수수료금액, LchTwoCmsnAmt, LchTwoCmsnAmt, double, 19.2;
		거래1수수료금액, TrdOneCmsnAmt, TrdOneCmsnAmt, double, 19.2;
		거래2수수료금액, TrdTwoCmsnAmt, TrdTwoCmsnAmt, double, 19.2;
		거래3수수료금액, TrdThreeCmsnAmt, TrdThreeCmsnAmt, double, 19.2;
		단기1수수료금액, StrmOneCmsnAmt, StrmOneCmsnAmt, double, 19.2;
		단기2수수료금액, StrmTwoCmsnAmt, StrmTwoCmsnAmt, double, 19.2;
		단기3수수료금액, StrmThreeCmsnAmt, StrmThreeCmsnAmt, double, 19.2;
		전달1수수료금액, TransOneCmsnAmt, TransOneCmsnAmt, double, 19.2;
		전달2수수료금액, TransTwoCmsnAmt, TransTwoCmsnAmt, double, 19.2;
		전달3수수료금액, TransThreeCmsnAmt, TransThreeCmsnAmt, double, 19.2;
		전달4수수료금액, TransFourCmsnAmt, TransFourCmsnAmt, double, 19.2;
		해외옵션행사예약구분코드, OvrsOptXrcRsvTpCode, OvrsOptXrcRsvTpCode, char, 1;
	end
	END_DATA_MAP
END_FUNCTION_MAP
