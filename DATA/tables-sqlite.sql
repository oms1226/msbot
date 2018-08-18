CREATE TABLE [종목코드](
  [종목명] TEXT, 
  [단축코드] TEXT PRIMARY KEY, 
  [확장코드] TEXT, 
  [ETF구분] INTEGER, 
  [상한가] INTEGER, 
  [하한가] INTEGER, 
  [전일가] INTEGER, 
  [주문수량단위] INTEGER, 
  [기준가] INTEGER, 
  [구분] INTEGER, 
  [증권그룹] TEXT, 
  [기업인수목적회사여부] TEXT);


CREATE TABLE [일별주가](
  [날짜] TEXT, 
  [시가] INTEGER, 
  [고가] INTEGER, 
  [저가] INTEGER, 
  [종가] INTEGER, 
  [전일대비구분] TEXT, 
  [전일대비] INTEGER, 
  [등락율] REAL, 
  [누적거래량] INTEGER, 
  [거래증가율] REAL, 
  [체결강도] REAL, 
  [소진율] REAL, 
  [회전율] REAL, 
  [외인순매수] INTEGER, 
  [기관순매수] INTEGER, 
  [종목코드] TEXT, 
  [누적거래대금] INTEGER, 
  [개인순매수] INTEGER, 
  [시가대비구분] TEXT, 
  [시가대비] INTEGER, 
  [시가기준등락율] REAL, 
  [고가대비구분] TEXT, 
  [고가대비] INTEGER, 
  [고가기준등락율] REAL, 
  [저가대비구분] TEXT, 
  [저가대비] INTEGER, 
  [저가기준등락율] REAL, 
  [시가총액] INTEGER, 
  PRIMARY KEY([날짜], [종목코드]));

CREATE TABLE [분별주가](
  [시간] TEXT, 
  [종가] INTEGER, 
  [전일대비구분] TEXT, 
  [전일대비] INTEGER, 
  [등락율] REAL, 
  [체결강도] REAL, 
  [매도체결수량] INTEGER, 
  [매수체결수량] INTEGER, 
  [순매수체결량] INTEGER, 
  [매도체결건수] INTEGER, 
  [매수체결건수] INTEGER, 
  [순체결건수] INTEGER, 
  [거래량] INTEGER, 
  [시가] INTEGER, 
  [고가] INTEGER, 
  [저가] INTEGER, 
  [체결량] INTEGER, 
  [매도체결건수시간] INTEGER, 
  [매수체결건수시간] INTEGER, 
  [매도잔량] INTEGER, 
  [매수잔량] INTEGER, 
  [시간별매도체결량] INTEGER, 
  [시간별매수체결량] INTEGER, 
  [단축코드] TEXT, 
  PRIMARY KEY([시간], [단축코드]));

CREATE TABLE [업종코드](
  [업종명] TEXT, 
  [업종코드] TEXT PRIMARY KEY);

CREATE TABLE [업종정보](
  [일자] TEXT, 
  [지수] REAL, 
  [전일대비구분] TEXT, 
  [전일대비] REAL, 
  [등락율] REAL, 
  [거래량] INTEGER, 
  [거래증가율] REAL, 
  [거래대금1] INTEGER, 
  [상승] INTEGER, 
  [보합] INTEGER, 
  [하락] INTEGER, 
  [상승종목비율] REAL, 
  [외인순매수] INTEGER, 
  [시가] REAL, 
  [고가] REAL, 
  [저가] REAL, 
  [거래대금2] INTEGER, 
  [상한] INTEGER, 
  [하한] INTEGER, 
  [종목수] INTEGER, 
  [기관순매수] INTEGER, 
  [업종코드] TEXT, 
  [거래비중] REAL, 
  [업종배당수익률] REAL, 
  PRIMARY KEY([일자], [업종코드]));


CREATE TABLE [종목별투자자](
  [일자] TEXT, 
  [종가] INTEGER, 
  [전일대비구분] TEXT, 
  [전일대비] INTEGER, 
  [등락율] REAL, 
  [누적거래량] INTEGER, 
  [사모펀드] INTEGER, 
  [증권] INTEGER, 
  [보험] INTEGER, 
  [투신] INTEGER, 
  [은행] INTEGER, 
  [종금] INTEGER, 
  [기금] INTEGER, 
  [기타법인] INTEGER, 
  [개인] INTEGER, 
  [등록외국인] INTEGER, 
  [미등록외국인] INTEGER, 
  [국가외] INTEGER, 
  [기관] INTEGER, 
  [외인계] INTEGER, 
  [기타계] INTEGER, 
  [단축코드] TEXT, 
  PRIMARY KEY([일자], [단축코드]));

CREATE TABLE [재무정보](
  [날짜] DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', 
  [종목코드] VARCHAR(8) NOT NULL DEFAULT '', 
  [기간구분] VARCHAR(16) NOT NULL DEFAULT '', 
  [매출액] DOUBLE DEFAULT NULL, 
  [영업이익] DOUBLE DEFAULT NULL, 
  [당기순이익] DOUBLE DEFAULT NULL, 
  [자산총계] DOUBLE DEFAULT NULL, 
  [부채총계] DOUBLE DEFAULT NULL, 
  [자본총계] DOUBLE DEFAULT NULL, 
  [자본금] DOUBLE DEFAULT NULL, 
  [부채비율] DOUBLE DEFAULT NULL, 
  [유보율] DOUBLE DEFAULT NULL, 
  [영업이익률] DOUBLE DEFAULT NULL, 
  [순이익률] DOUBLE DEFAULT NULL, 
  [ROA] DOUBLE DEFAULT NULL, 
  [ROE] DOUBLE DEFAULT NULL, 
  [EPS] DOUBLE DEFAULT NULL, 
  [BPS] DOUBLE DEFAULT NULL, 
  [DPS] DOUBLE DEFAULT NULL, 
  [PER] DOUBLE DEFAULT NULL, 
  [PBR] DOUBLE DEFAULT NULL, 
  [발행주식수] DOUBLE DEFAULT NULL, 
  [배당수익률] DOUBLE DEFAULT NULL, 
  PRIMARY KEY([날짜], [종목코드], [기간구분]));

CREATE INDEX [idx_financial_date] ON [재무정보]([날짜]);

CREATE INDEX [idx_financial_code] ON [재무정보]([종목코드]);


CREATE TABLE [증권사추천주](
  [일자] DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', 
  [종목코드] VARCHAR(8) NOT NULL DEFAULT '', 
  [종목명] VARCHAR(64) DEFAULT NULL, 
  [의견] VARCHAR(16) DEFAULT NULL, 
  [목표가] BIGINT(20) DEFAULT NULL, 
  [추천일가격] BIGINT(20) DEFAULT NULL, 
  [추천증권사] VARCHAR(64) NOT NULL DEFAULT '', 
  [추천사유] VARCHAR(2048) DEFAULT NULL, 
  PRIMARY KEY([일자], [종목코드], [추천증권사]));

CREATE INDEX [ix_stockrecommandfnguide_date] ON [증권사추천주]([일자]);

CREATE INDEX [ix_stockrecommandfnguide_code] ON [증권사추천주]([종목코드]);

CREATE INDEX [ix_stockrecommandfnguide_position] ON [증권사추천주]([의견]);

CREATE INDEX [ix_stockrecommandfnguide_broker] ON [증권사추천주]([추천증권사]);

CREATE TABLE [거래결과](
  [로봇명] TEXT, 
  [UUID] TEXT, 
  [일자] TEXT, 
  [체결시각] TEXT, 
  [단축종목번호] TEXT, 
  [종목명] TEXT, 
  [매매구분] TEXT, 
  [주문번호] TEXT, 
  [체결번호] TEXT, 
  [주문수량] INT, 
  [주문가격] INT, 
  [체결수량] INT, 
  [체결가격] INT, 
  [주문평균체결가격] FLOAT);

CREATE INDEX [idx_trade_uuid] ON [거래결과]([UUID]);