BEGIN_FUNCTION_MAP
.Feed, 해외선물체결, TC3, block, key=7, group=1;
    BEGIN_DATA_MAP
    InBlock,입력,input;
    begin
    end
    OutBlock,출력,output;
    begin
		라인일련번호,  	lineseq,    lineseq,	long,   10;
		KEY,		    key,		key,		char,	11;
		조작자ID,		user,	    user,	   	char,	 8;

        서비스ID,           svc_id,             svc_id,             char,    4;
        주문일자,           ordr_dt,            ordr_dt,            char,    8;
        지점번호,           brn_cd,             brn_cd,             char,    3;
        주문번호,           ordr_no,            ordr_no,            long,    10;
        원주문번호,         orgn_ordr_no,       orgn_ordr_no,       long,    10;
        모주문번호,         mthr_ordr_no,       mthr_ordr_no,       long,    10;
        계좌번호,           ac_no,              ac_no,              char,    11;
        종목코드,           is_cd,              is_cd,              char,    30;
        매도매수유형,       s_b_ccd,            s_b_ccd,            char,    1;
        정정취소유형,       ordr_ccd,           ordr_ccd,           char,    1;
        체결수량,           ccls_q,             ccls_q,             long,    15;
        체결가격,           ccls_prc,           ccls_prc,           double,  15.8;
        체결번호,           ccls_no,            ccls_no,            char,    10;
        체결시간,           ccls_tm,            ccls_tm,            char,    9;
        매입평균단가,       avg_byng_uprc,      avg_byng_uprc,      double,  12.6;
        매입금액,           byug_amt,           byug_amt,           double,  25.8;
        청산손익,           clr_pl_amt,         clr_pl_amt,         double,  19.2;
        위탁수수료,         ent_fee,            ent_fee,            double,  19.2;
        FCM수수료,          fcm_fee,            fcm_fee,            double,  19.2;
		사용자ID,           userid,             userid,             char,    8;
        현재가격,           now_prc,            now_prc,            double,  15.8;
        통화코드,           crncy_cd,           crncy_cd,           char,    3;
        만기일자,           mtrt_dt,            mtrt_dt,            char,    8;
    end
    END_DATA_MAP
END_FUNCTION_MAP
