BEGIN_FUNCTION_MAP
.Feed, 해외선물 현재가체결(OVC), OVC, attr, svr=OVS, key=8, group=1;
    BEGIN_DATA_MAP
    InBlock,입력,input;
    begin
        종목코드,       symbol,    symbol,    char,   8;
    end
    OutBlock,출력,output;
    begin
		종목코드         , symbol     , symbol    , char   ,   8;
		체결일자(현지)   , ovsdate    , ovsdate   , char   ,   8;
		체결일자(한국)   , kordate    , kordate   , char   ,   8;
		체결시간(현지)   , trdtm      , trdtm     , char   ,   6;
		체결시간(한국)   , kortm      , kortm     , char   ,   6;
		체결가격         , curpr      , curpr     , double ,   15.8;
		전일대비         , ydiffpr    , ydiffpr   , double ,   15.8;
		전일대비기호     , ydiffSign  , ydiffSign , char   ,   1;
		시가			 , open		  , open	  , double ,   15.8;
		고가			 , high		  , high	  , double ,   15.8;
		저가			 , low 		  , low 	  , double ,   15.8;
		등락율			 , chgrate	  , chgrate   , float  , 6.2;
		건별체결수량     , trdq       , trdq      , long   ,  10;
		누적체결수량     , totq       , totq      , char   ,  15;
		체결구분		 , cgubun     , cgubun    , char   ,   1;
		매도누적체결수량 , mdvolume   , mdvolume  , char   ,  15;
		매수누적체결수량 , msvolume   , msvolume  , char   ,  15;
		장마감일 		 , ovsmkend   , ovsmkend  , char   ,   8;
    end
    END_DATA_MAP
END_FUNCTION_MAP
