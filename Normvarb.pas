unit normvarb;
interface

const
   done                  : boolean   = False;

type
   string27                    = string[27];
   string5                     = string[5];
   string8                     = string[8];
   string10                    = string[10];
   string14                    = string[14];
   string80                    = string[80];
   Num10Array10                = array[1..10,1..10] of double;
   RootStr                     = string[12];
   PrinterStr                  = string[5];
   GpCode                      = array[1..3] of byte;
   RealArray60                 = array[0..60] of double;


var
   GlobalChosenStyle : string;
   Lst : TextFile;
   AnyKey                : char;
   Toggle100             : byte;
   FilePrepared          : boolean;
   B                     : array [1..120] of string27;
   ElementPos            : array [1..25] of integer;
   OxFactor              : array [1..25] of double;
   InOption, OutRoute    : char;
   Nbeg, Nfin,
   I, IP, IP1, Item      : integer;
   FullFileName                : string;
   MineralTablePath, cdsPath  : string;
   MaxData,MinData             : RealArray60;
   OxData                      : RealArray60;
   Indx,TotalRecs,RecCount,RecNum
                               : Integer;
   FileName                    : String8;
   AnyName                     : String10;
   Device                      : text;
   ElemPos                     : array[1..11] of byte;
   OXTOT,OSI,OTI,OAL,OKONA,
   OF3,OF2,OMN,OMG,OCA,
   ONA,OKO,OPO,OHP,OH2OM,
   OCD,OCR,OSR,OBA,OCL,
   OFU,ONI,OZR,OSU,OSO,
   CSI,CTI,CAL,
   CF3,CF2,CMN,CMG,CCA,
   CNA,CKO,CPO,CHP,CH2OM,
   CCD,CCR,CSR,CBA,CCL,
   CFU,CNI,CZR,CSU,CSO,
   SI,TI,AL,F23,
   F3,F2,MN,MG,CA,
   NA,KO,PO,HP,H2OM,NOH2O,
   CD,NCR,SR,BA,CL,
   FU,NI,ZR,SU,SO,
   WAFMA,WAFMF,WAFMM,AFMA,AFMF,
   AFMM,WAGRAT,AGRAT,
   WFMRAT,FMRAT,WALRAT,
   ALRAT,WOXRAT,OXRAT,
   WALKIN,ALK,ALK1,PF2,PMG,FEMG,
   WTFAN,WTFOR,WTFAB,TFAN,TFOR,TFAB,
   WTQZ,WTTAB,WTOR,WTTQZ,WTKP,WTNE,
   TQZ,TTAB,TOR,TTQZ,TKP,TNE,
   WTAB,WPL,NPL,F,WF,WOR,WQZ,WAB,
   WKP,WLC,WNE,
   SALIC,DIX,CO,ZN,AN,HL,TH,RI,ED,ACT,QZ,ORT,AB,NE,LC,KP,
   WSALIC,WDIX,WCO,WZN,WAN,WHL,WTH,
   FEMIC,AC,NS,KS,WO,DI,HY,OL,BI,HO,SPIN,CS,MT,CM,
   IL,HM,SP,PF,RU,AP,FL,PY,CC,
   WFEMIC,WAC,WNS,WKS,WWO,WDI,WHY,WOL,WBI,WHO,WSPIN,WCS,WMT,WCM,
   WIL,WHM,WSP,WPF,WRU,WAP,WFL,WPY,WCC,WACT,WRI,WED,
   TOTAL,SUM,WWODI,WENDI,WFSDI,WODI,ENDI,FSDI,
   WEN,WFS,WFO,WFA,FO,FA,EN,FS,
   WANPL,WFAOL,WENHY,ANPL,FAOL,ENHY,
   WTOTAL               : double;
   CTL,CTTOT            : double;
   R1, R2               : double;
   DES                  : string[60];
   A14,A15,A16          : string[120];
   WatM                 : double;
   mcSI, mcAL, mcFE3, mcFE2, mcMN, mcMG, mcCA, mcNA, mcK, mcTI : double;
   CalciumInNonsilicates, CIA : double;
   RoserKorschD1, RoserKorschD2, RoserKorschD3,
   RoserKorschD4 : double;
   ACNK : double;
   DebonLefortA, DebonLefortB : double;
   ShowOnly50Rows : boolean;
   IniFileName,
   IniFilePath,
   CommonFilePath,
   LocalFilePath,
   ProgramFilePath,
   FlexTemplatePath,
   ExportPath,
   DataPath   : string;
   ImportSheetNumber,
   PositionCol, RequiredCol, EnteredCol,
   FactorCol,
   ColumnCol            : integer;
   PositionColStr, RequiredColStr, EnteredColStr,
   FactorColStr,
   ColumnColStr         : string;
   FromRowValueString, ToRowValueString : string;
   FromRow, ToRow : integer;
   RowCount             : array[1..10] of integer;
   DefaultMinimum : double;


implementation

end.
