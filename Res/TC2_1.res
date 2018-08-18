BEGIN_FUNCTION_MAP
.Feed, 해외선물응답, TC2, block, key=7, group=1;
    BEGIN_DATA_MAP
    InBlock,입력,input;
    begin
    end
    OutBlock,출력,output;
    begin
		라인일련번호,  	lineseq,    lineseq,	long,   10;
		KEY,		    key,		key,		char,	11;
		조작자ID,		user,	    user,	   	char,	 8;

		서비스ID,           svc_id,             svc_id,	            char,    4;
		주문일자,           ordr_dt,            ordr_dt,	        char,    8;
		지점번호,           brn_cd,             brn_cd,             char,    3;
		주문번호,           ordr_no,            ordr_no,            long,    10;
		원주문번호,         orgn_ordr_no,       orgn_ordr_no,       long,    10;
		모주문번호,         mthr_ordr_no,       mthr_ordr_no,       long,    10;
		계좌번호,           ac_no,              ac_no,              char,    11;
		종목코드,           is_cd,              is_cd,              char,    30;
		매도매수유형,       s_b_ccd,            s_b_ccd,            char,    1;
		정정취소유형,       ordr_ccd,           ordr_ccd,           char,    1;
		주문유형코드,       ordr_typ_cd,        ordr_typ_cd,        char,    1;
		주문기간코드,       ordr_typ_prd_ccd,   ordr_typ_prd_ccd,   char,    2;
		주문적용시작일자,   ordr_aplc_strt_dt,  ordr_aplc_strt_dt,  char,    8;
		주문적용종료일자,   ordr_aplc_end_dt,   ordr_aplc_end_dt,   char,    8;
		주문가격,           ordr_prc,           ordr_prc,           double,  15.8;
		주문조건가격,       cndt_ordr_prc,      cndt_ordr_prc,      double,  15.8;
		주문수량,           ordr_q,             ordr_q,             long,    12;
		주문시간,           ordr_tm,            ordr_tm,            char,    9;
		호가확인수량,       cnfr_q,             cnfr_q,             long,   12;
		호가거부사유코드,   rfsl_cd,            rfsl_cd,            char,    4;
		호가거부사유코드명, text,               text,               char,   300;
    end
    END_DATA_MAP
END_FUNCTION_MAP
