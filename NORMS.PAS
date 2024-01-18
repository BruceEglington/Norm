program Norms;
{$B-} {boolean evaluation}
{$D-} {debug information}
{$E+} {8087 emulation}
{$L-} {local symbol information}
{$N+} {numeric processing}
{$R-} {range checking}
{$S-} {stack checking}

uses
  crt, dos, readsys2,
  protect, lptcheck,
  findfile, sys_def2,
  sys_def, inputset,
  normvarb, norm_min,
  normtern, normdesi,
  nor_asci, nor_prep,
  kem_varb, allsorts,
  xvideo, filselec;

const
   LocalProgName     = 'NORMS';
   LocalUserCode     = 'KN-SADTC/91';
   LocalYearStr      = '2000';
   LocalMonthStr     =   '12';
   LocalDayStr       =   '31';
   OvrName           = 'NORMS.OVR';
   BombChar          = 'N';

var
  L, LMin, LMax  : byte;


{------------------Procedures----------------------------------}

procedure FileBanner;
const
  ListLines = 21;
  StartLine = 2;
  List : array[1..ListLines] of string =
(
'     ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ   ',
'     ÛÛ ÉÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ» ÛÛ±± ',
'     ÛÛ º     **   N O R M S   *   N O R M S   *   N O R M S   **    º ÛÛ±± ',
'     ÛÛ º    *   N O R M S   ***   N O R M S   ***   N O R M S   *   º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º      A normative mineralogy calculation package which:     º ÛÛ±± ',
'     ÛÛ º                                                            º ÛÛ±± ',
'     ÛÛ º           û  Accesses data from files created by KEMFILE   º ÛÛ±± ',
'     ÛÛ º           û  Calculates CIPW norms                         º ÛÛ±± ',
'     ÛÛ º           û  Calculates either bi+px or bi+amph+px         º ÛÛ±± ',
'     ÛÛ º              meso-norms                                    º ÛÛ±± ',
'     ÛÛ º           û  Calculates various projections                º ÛÛ±± ',
'     ÛÛ º           û  Outputs results to printer or an ASCII        º ÛÛ±± ',
'     ÛÛ º              data interchange format on file               º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º   S.A. Development Trust Co. (Site License  KN-SADTC/91)   º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º    Version 2.1/180291            c) KEMSYS Software 1991   º ÛÛ±± ',
'     ÛÛ ÈÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼ ÛÛ±± ',
'     ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ±± ',
'        ±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±± '
);

var
 i : byte;
 ch : integer;

begin
    ClrScr;    OffScr;
    for i:=1 to ListLines do
      PrintS(i+StartLine,2,List[i],11);
    Centre(ListLines+StartLine+1,'Press [Space Bar] to Continue, or [Esc] to exit',10);
    Draw_Block(1,1,25,80,1,14);
(*       At({ListLines+StartLine+3}25,37);Write('(',MagicN,':',LocalMaxUse,')');*)
    ChangeAtt(ListLines+StartLine-3,15,53,9);
    ChangeAtt(StartLine+3,12,57,14);
    ChangeAtt(StartLine+4,12,57,14);
    OnScr;
    At(25,40);
    repeat
     ch:=Ord(ReadKey);
    until (ch in [32,esckey]);
    ClrScr;
     if ch=esckey then
     begin
       Terminate;
       Halt;
     end;
end;


procedure OpenBanner;
begin
  FileBanner;
end;{proc OpenBanner}


procedure SelectFile1;
begin
  ClrScr;
  {$I-}
  FileName:='        ';
  SExtension:='KEM';
  SelectFile(1,1,'Source file',138,'N',Drive2,FileName,SExtension,DataFile);
  if (FileName<>'ABORT') then begin
    TotalRecs:=FileSize(DataFile)-3;
    RecCount:=FileSize(DataFile);
    TFileName:=FileName;
  end;
  {$I+}
end;            {of SelectFile1 }


procedure Names;
begin
            B[1]:='SiO2 ';
            B[2]:=' Qz ';
            B[3]:='An/Pl  Wt pct';
            B[4]:='Mol pct';
            B[5]:='TiO2 ';
            B[6]:=' Co ';
            B[7]:='ZrO2 ';
            B[8]:=' Z  ';
            B[9]:='Fa/Ol  Wt pct';
            B[10]:='Mol pct';
            B[11]:='Al2O3';
            B[12]:=' Or ';
            B[13]:='Cr2O3';
            B[14]:=' Pl ';
            B[15]:='En/Hy  Wt pct';
            B[16]:='Mol pct';
            B[17]:='Fe2O3';
            B[18]:='(Ab)';
            B[19]:='FeO  ';
            B[20]:='(An)';
            B[21]:='Difndx Wt pct';
            B[22]:='Cat pct';
            B[23]:='MnO  ';
            B[24]:=' Lc ';
            B[25]:='NiO  ';
            B[26]:=' Ne ';
            B[27]:='QZ-Ab-Or diagram       ';
            B[28]:='MgO  ';
            B[29]:=' Ks ';
            B[30]:='CaO  ';
            B[31]:=' Hl ';
            B[32]:='Wt pct  Qz';
            B[33]:=' Ab ';
            B[34]:=' Or ';
            B[35]:='SrO  ';
            B[36]:=' Th ';
            B[37]:='BaO  ';
            B[38]:='Mol pct Qz';
            B[39]:=' Ab ';
            B[40]:=' Or ';
            B[41]:='Na2O ';

            B[43]:='K2O  ';
            B[44]:=' Ns ';
            B[45]:='Qz-Ne-Ks diagram';
            B[46]:='P2O5 ';
            B[47]:=' Ks ';
            B[48]:='H2O+ ';
            B[49]:=' Wo ';
            B[50]:='Wt pct  Qz';
            B[51]:=' Ne ';
            B[52]:=' Ks ';
            B[53]:='H2O- ';
            B[54]:=' Di ';
            B[55]:='SO3  ';
            B[56]:='(Wo)';
            B[57]:='Mol pct Qz';
            B[58]:=' Ne ';
            B[59]:=' Ks ';
            B[60]:='S    ';
            B[61]:='(En)';
            B[62]:='Cl   ';
            B[63]:='(Fs)';
            B[64]:='An-Ab-Or diagram';
            B[65]:='F    ';
            B[66]:=' Hy ';
            B[67]:='CO2  ';
            B[68]:='(En)';
            B[69]:='Wt pct  An';
            B[70]:=' Ab ';
            B[71]:=' Or ';
            B[72]:='(Fs)';
            B[73]:='Total';
            B[74]:=' Ol ';
            B[75]:='Mol pct An';
            B[76]:=' Ab ';
            B[77]:=' Or ';
            B[78]:='-H2O ';
            B[79]:='(Fo)';
            B[80]:='(Fa)';
            B[81]:=' A-F-M diagram';
            B[115]:='Watson & Harrison  M     ';
   case InOption of
     '0' : begin
            B[42]:=' Ac ';

            B[82]:=' Cs ';
            B[83]:=' Mt ';
            B[84]:='Wt pct Alk';
            B[85]:=' Fe ';
            B[86]:=' Mg ';
            B[87]:=' Cm ';
            B[88]:=' Il ';
            B[89]:='Mol pct Alk';
            B[90]:=' Fe ';
            B[91]:=' Mg ';
            B[92]:=' Hm ';
            B[93]:=' Sp ';
            B[94]:='(Na+K)/Al          Atm wt %';
            B[95]:='Grmatm %  ';
            B[96]:=' Pf ';
            B[97]:=' Ru ';
            B[98]:='(Fe+Mn)/(Fe+Mn+Mg) Atm wt %';
            B[99]:='Grmatm %  ';
            B[100]:=' Ap ';
            B[101]:=' Fl ';
            B[102]:='Na2O/(Na2O+K2O)        wt %';
            B[103]:='   Mol %  ';
            B[104]:=' Py ';
            B[105]:=' Cc ';
            B[106]:='FeO/(FeO+Fe2O3)        wt %';
            B[107]:='   Mol %  ';
            B[108]:='Salic';
            B[109]:='Wright`s alkalinity index';
            B[110]:='Femic';
            B[111]:='Total';
         end;
   '1','2' : begin
            B[42]:=' Bi ';

            B[82]:=' Ho ';
            B[83]:='(Ri)';
            B[84]:='Wt pct Alk';
            B[85]:=' Fe ';
            B[86]:=' Mg ';
            B[87]:='(Act)';
            B[88]:='(Ed)';
            B[89]:='Mol pct Alk';
            B[90]:=' Fe ';
            B[91]:=' Mg ';
            B[92]:='Spnl';
            B[93]:=' Cs ';
            B[94]:='(Na+K)/Al          Atm wt %';
            B[95]:='Grmatm %  ';
            B[96]:=' Mt ';
            B[97]:=' Cm ';
            B[98]:='(Fe+Mn)/(Fe+Mn+Mg) Atm wt %';
            B[99]:='Grmatm %  ';
            B[100]:=' Hm ';
            B[101]:=' Sp ';
            B[102]:='Na2O/(Na2O+K2O)        wt %';
            B[103]:='   Mol %  ';
            B[104]:=' Ru ';
            B[105]:=' Ap ';
            B[106]:='FeO/(FeO+Fe2O3)        wt %';
            B[107]:='   Mol %  ';
            B[108]:=' Fl ';
            B[109]:=' Py ';
            B[110]:='Wright`s alkalinity index';
            B[111]:=' Cc ';
            B[112]:='Salic';
            B[113]:='Femic';
            B[114]:='Total';
         end;
   end;{case}
end;{proc Names}


procedure GetElementOrder;
var
  NumMaj  : byte;
begin
  Seek(DataFile,0);
  Read(DataFile,DataField);
  NumMaj:=DataField.MajorNo;
  Seek(DataFile,1);
  Read(DataFile,DataField);
  if DataField.NormFlag=0 then
  begin
    ClrScr;
    Centre(12,'Data file not prepared for NORM element order',NormBld);
    Beep; Beep;
    Delay(4000);
    ClrScr;
    FilePrepared:=false;
  end
  else begin
    FilePrepared:=true;
  end;
  if FilePrepared then
  begin
   for I:=1 to 23 do begin
     ElementPos[I]:=DataField.Pos25[I];
   end;
   OxFactor[1]:=0.0002139;  {Si}
   OxFactor[2]:=0.0001668;  {Ti}
   OxFactor[3]:=0.000135;   {Zr}
   OxFactor[4]:=0.0001889;  {Al}
   OxFactor[5]:=0.0001461;  {Cr}
   OxFactor[6]:=0.0001429;  {Fe3}
   OxFactor[7]:=0.0001286;  {Fe2}
   OxFactor[8]:=0.0001291;  {Mn}
   OxFactor[9]:=0.0001272;  {Ni}
   OxFactor[10]:=0.0001658; {Mg}
   OxFactor[11]:=0.0001399; {Ca}
   OxFactor[12]:=0.0001182; {Sr}
   OxFactor[13]:=0.0001116; {Ba}
   OxFactor[14]:=0.0001347; {Na}
   OxFactor[15]:=0.0001204; {K }
   OxFactor[16]:=0.0002291; {P }
   OxFactor[17]:=0.0001000; {H+}
   OxFactor[18]:=0.0001;    {H-}
   OxFactor[19]:=0.0001;    {SO3}
   OxFactor[20]:=0.0001;    {S }
   OxFactor[21]:=0.0001;    {Cl}
   OxFactor[22]:=0.0001;    {F }
   OxFactor[23]:=0.0001;    {CO2}

   for I:=24 to 25 do begin
     if (ElementPos[I]<>0) then begin
       OxFactor[I]:=1.0;
     end;
   end;
   for I:=1 to 23 do begin
    if (ElementPos[I]<=NumMaj) then OxFactor[I]:=1.0;
    if (ElementPos[I]=0) then begin
      OxFactor[I]:=0.0;
      ElementPos[I]:=1;
    end;
   end;
  end;
end;{proc GetElementOrder}


procedure CIPWPrint;
begin
   Writeln(ListFile,CondPrn);
   Writeln(ListFile);
   Writeln(ListFile);
   Writeln(ListFile);
   Write(ListFile,' ':10,DataField.SampleNo:10,' ':4,DES:60,' ':8,'CIPW-norm');
   Writeln(ListFile,' version '+LocalUserCode);
   Writeln(ListFile,StringOf(92,32),'Printed on ',DateString);
   if Toggle100=1 then
   begin
     Write(ListFile,StringOf(92,32),'Mineral wt % recalculated to 100%');
   end;
   Writeln(ListFile);
   Writeln(ListFile,A14);
   Writeln(ListFile,A15);
   Writeln(ListFile,A16);
   Writeln(ListFile);
   Writeln(ListFile, B[1]:15,OSI:12:2,CSI:9:2,SI:9:2, B[2]:10,WQZ:9:2,QZ:9:2, B[3]:19,WANPL:9:2, B[4]:13,ANPL:9:2);
   Writeln(ListFile, B[5]:15,OTI:12:2,CTI:9:2,TI:9:2, B[6]:10,WCO:9:2,CO:9:2);
   Writeln(ListFile, B[7]:15,OZR:12:2,CZR:9:2,ZR:9:2, B[8]:10,WZN:9:2,ZN:9:2, B[9]:19,WFAOL:9:2, B[10]:13,FAOL:9:2);
   Writeln(ListFile,B[11]:15,OAL:12:2,CAL:9:2,AL:9:2,B[12]:10,WOR:9:2,ORT:9:2);
   Writeln(ListFile,B[13]:15,OCR:12:2,CCR:9:2,NCR:9:2,B[14]:10,WPL:9:2,NPL:9:2,B[15]:19,WENHY:9:2,B[16]:13,ENHY:9:2);
   Writeln(ListFile,B[17]:15,OF3:12:2,CF3:9:2,F3:9:2,B[18]:10,WAB:9:2,AB:9:2);
   Writeln(ListFile,B[19]:15,OF2:12:2,CF2:9:2,F2:9:2,B[20]:10,WAN:9:2,AN:9:2,B[21]:19,WDIX:9:2,B[22]:13,DIX:9:2);
   Writeln(ListFile,B[23]:15,OMN:12:2,CMN:9:2,B[24]:19,WLC:9:2,LC:9:2);
   Writeln(ListFile,B[25]:15,ONI:12:2,CNI:9:2,B[26]:19,WNE:9:2,NE:9:2,B[27]:44);
   Writeln(ListFile,B[28]:15,OMG:12:2,CMG:9:2,MG:9:2,B[29]:10,WKP:9:2,KP:9:2);
   Write(ListFile,B[30]:15,OCA:12:2,CCA:9:2,CA:9:2,B[31]:10,WHL:9:2,HL:9:2);
   Writeln(ListFile,B[32]:15,WTQZ:7:2,B[33]:7,WTAB:7:2,B[34]:7,WTOR:7:2);
   Writeln(ListFile,B[35]:15,OSR:12:2,CSR:9:2,B[36]:19,WTH:9:2,TH:9:2);
   Writeln(ListFile,B[37]:15,OBA:12:2,CBA:9:2,B[38]:52,TQZ:7:2,B[39]:7,TTAB:7:2,B[40]:7,TOR:7:2);
   Writeln(ListFile,B[41]:15,ONA:12:2,CNA:9:2,NA:9:2,B[42]:10,WAC:9:2,AC:9:2);
   Writeln(ListFile,B[43]:15,OKO:12:2,CKO:9:2,KO:9:2,B[44]:10,WNS:9:2,NS:9:2,B[45]:37);
   Writeln(ListFile,B[46]:15,OPO:12:2,CPO:9:2,PO:9:2,B[47]:10,WKS:9:2,KS:9:2);
   Writeln(ListFile,B[48]:15,OHP:12:2,B[49]:28,WWO:9:2,WO:9:2,B[50]:15,WTTQZ:7:2,B[51]:7,WTNE:7:2,B[52]:7,WTKP:7:2);
   Writeln(ListFile,B[53]:15,OH2OM:12:2,B[54]:28,WDI:9:2,DI:9:2);
   Write(ListFile,B[55]:15,OSO:12:2,CSO:9:2,SO:9:2,B[56]:10,WWODI:9:2,WODI:9:2);
   Writeln(ListFile,B[57]:15,TTQZ:7:2,B[58]:7,TNE:7:2,B[59]:7,TKP:7:2);
   Writeln(ListFile,B[60]:15,OSU:12:2,CSU:9:2,SU:9:2,B[61]:10,WENDI:9:2,ENDI:9:2);
   Writeln(ListFile,B[62]:15,OCL:12:2,CCL:9:2,CL:9:2,B[63]:10,WFSDI:9:2,FSDI:9:2,B[64]:37);
   Writeln(ListFile,B[65]:15,OFU:12:2,CFU:9:2,FU:9:2,B[66]:10,WHY:9:2,HY:9:2);
   Write(ListFile,B[67]:15,OCD:12:2,CCD:9:2,CD:9:2,B[68]:10,WEN:9:2,EN:9:2);
   Writeln(ListFile,B[69]:15,WTFAN:7:2,B[70]:7,WTFAB:7:2,B[71]:7,WTFOR:7:2);
   Writeln(ListFile,' ':45,B[72]:10,WFS:9:2,FS:9:2);
   Write(ListFile,B[73]:15,OXTOT:12:2,CTTOT:9:2,B[74]:19,WOL:9:2,OL:9:2);
   Writeln(ListFile,B[75]:15,TFAN:7:2,B[76]:7,TFAB:7:2,B[77]:7,TFOR:7:2);
   Writeln(ListFile,B[78]:15,NOH2O:12:2,B[79]:28,WFO:9:2,FO:9:2);
   Writeln(ListFile,' ':45,B[80]:10,WFA:9:2,FA:9:2,B[81]:35);
   Writeln(ListFile,' ':45,B[82]:10,WCS:9:2,CS:9:2);
   Writeln(ListFile,' ':45,B[83]:10,WMT:9:2,MT:9:2,B[84]:15,WAFMA:7:2,B[85]:7,WAFMF:7:2,B[86]:7,WAFMM:7:2);
   Writeln(ListFile,' ':45,B[87]:10,WCM:9:2,CM:9:2);
   Writeln(ListFile,' ':45,B[88]:10,WIL:9:2,IL:9:2,B[89]:15,AFMA:7:2,B[90]:7,AFMF:7:2,B[91]:7,AFMM:7:2);
   Writeln(ListFile,' ':45,B[92]:10,WHM:9:2,HM:9:2);
   Writeln(ListFile,' ':45,B[93]:10,WSP:9:2,SP:9:2,B[94]:29,WAGRAT:6:2,B[95]:11,AGRAT:6:2);
   Writeln(ListFile,' ':45,B[96]:10,WPF:9:2,PF:9:2);
   Writeln(ListFile,' ':45,B[97]:10,WRU:9:2,RU:9:2,B[98]:29,WFMRAT:6:2,B[99]:11,FMRAT:6:2);
   Writeln(ListFile,' ':45,B[100]:10,WAP:9:2,AP:9:2);
   Writeln(ListFile,' ':45,B[101]:10,WFL:9:2,FL:9:2,B[102]:29,WALRAT:6:2,B[103]:11,ALRAT:6:2);
   Writeln(ListFile,' ':45,B[104]:10,WPY:9:2,PY:9:2);
   Writeln(ListFile,' ':45,B[105]:10,WCC:9:2,CC:9:2,B[106]:29,WOXRAT:6:2,B[107]:11,OXRAT:6:2);
   Writeln(ListFile);
   Writeln(ListFile,' ':50,B[108]:5,WSALIC:9:2,SALIC:9:2,B[109]:27,WALKIN:8:2);
   Writeln(ListFile,' ':50,B[110]:5,WFEMIC:9:2,FEMIC:9:2,B[115]:27,WatM:8:2);
   Writeln(ListFile);
   Writeln(ListFile,' ':50,B[111]:5,WTOTAL:9:2,TOTAL:9:2);
   Writeln(ListFile,NormPrn);
   Writeln(ListFile,FF);
end;{proc CIPWPrint}



procedure MesoPrint;
begin
   Writeln(ListFile,CondPrn);
   Writeln(ListFile);
   Writeln(ListFile);
   Writeln(ListFile);
   Write(ListFile,' ':10,DataField.SampleNo:10,' ':4,DES:60,' ':8,'Meso-norm');
   Writeln(ListFile,' version '+LocalUserCode);
   Writeln(ListFile,StringOf(92,32),'Printed on ',DateString);
   if Toggle100=1 then
   begin
     Write(ListFile,StringOf(92,32),'Mineral wt % recalculated to 100%');
   end;
   Writeln(ListFile);
   Writeln(ListFile,A14);
   Writeln(ListFile,A15);
   Writeln(ListFile,A16);
   Writeln(ListFile);
   Writeln(ListFile, B[1]:15,OSI:12:2,CSI:9:2,SI:9:2, B[2]:10,WQZ:9:2,QZ:9:2, B[3]:19,WANPL:9:2, B[4]:13,ANPL:9:2);
   Writeln(ListFile, B[5]:15,OTI:12:2,CTI:9:2,TI:9:2, B[6]:10,WCO:9:2,CO:9:2);
   Writeln(ListFile, B[7]:15,OZR:12:2,CZR:9:2,ZR:9:2, B[8]:10,WZN:9:2,ZN:9:2, B[9]:19,WFAOL:9:2, B[10]:13,FAOL:9:2);
   Writeln(ListFile,B[11]:15,OAL:12:2,CAL:9:2,AL:9:2,B[12]:10,WOR:9:2,ORT:9:2);
   Writeln(ListFile,B[13]:15,OCR:12:2,CCR:9:2,NCR:9:2,B[14]:10,WPL:9:2,NPL:9:2,B[15]:19,WENHY:9:2,B[16]:13,ENHY:9:2);
   Writeln(ListFile,B[17]:15,OF3:12:2,CF3:9:2,F3:9:2,B[18]:10,WAB:9:2,AB:9:2);
   Writeln(ListFile,B[19]:15,OF2:12:2,CF2:9:2,F2:9:2,B[20]:10,WAN:9:2,AN:9:2,B[21]:19,WDIX:9:2,B[22]:13,DIX:9:2);
   Writeln(ListFile,B[23]:15,OMN:12:2,CMN:9:2,B[24]:19,WLC:9:2,LC:9:2);
   Writeln(ListFile,B[25]:15,ONI:12:2,CNI:9:2,B[26]:19,WNE:9:2,NE:9:2,B[27]:44);
   Writeln(ListFile,B[28]:15,OMG:12:2,CMG:9:2,MG:9:2,B[29]:10,WKP:9:2,KP:9:2);
   Write(ListFile,B[30]:15,OCA:12:2,CCA:9:2,CA:9:2,B[31]:10,WHL:9:2,HL:9:2);
   Writeln(ListFile,B[32]:15,WTQZ:7:2,B[33]:7,WTAB:7:2,B[34]:7,WTOR:7:2);
   Writeln(ListFile,B[35]:15,OSR:12:2,CSR:9:2,B[36]:19,WTH:9:2,TH:9:2);
   Writeln(ListFile,B[37]:15,OBA:12:2,CBA:9:2,B[38]:52,TQZ:7:2,B[39]:7,TTAB:7:2,B[40]:7,TOR:7:2);
   Writeln(ListFile,B[41]:15,ONA:12:2,CNA:9:2,NA:9:2,B[42]:10,WBI:9:2,BI:9:2);
   Writeln(ListFile,B[43]:15,OKO:12:2,CKO:9:2,KO:9:2,B[44]:10,WNS:9:2,NS:9:2,B[45]:37);
   Writeln(ListFile,B[46]:15,OPO:12:2,CPO:9:2,PO:9:2,B[47]:10,WKS:9:2,KS:9:2);
   Writeln(ListFile,B[48]:15,OHP:12:2,B[49]:28,WWO:9:2,WO:9:2,B[50]:15,WTTQZ:7:2,B[51]:7,WTNE:7:2,B[52]:7,WTKP:7:2);
   Writeln(ListFile,B[53]:15,OH2OM:12:2,B[54]:28,WDI:9:2,DI:9:2);
   Write(ListFile,B[55]:15,OSO:12:2,CSO:9:2,SO:9:2,B[56]:10,WWODI:9:2,WODI:9:2);
   Writeln(ListFile,B[57]:15,TTQZ:7:2,B[58]:7,TNE:7:2,B[59]:7,TKP:7:2);
   Writeln(ListFile,B[60]:15,OSU:12:2,CSU:9:2,SU:9:2,B[61]:10,WENDI:9:2,ENDI:9:2);
   Writeln(ListFile,B[62]:15,OCL:12:2,CCL:9:2,CL:9:2,B[63]:10,WFSDI:9:2,FSDI:9:2,B[64]:37);
   Writeln(ListFile,B[65]:15,OFU:12:2,CFU:9:2,FU:9:2,B[66]:10,WHY:9:2,HY:9:2);
   Write(ListFile,B[67]:15,OCD:12:2,CCD:9:2,CD:9:2,B[68]:10,WEN:9:2,EN:9:2);
   Writeln(ListFile,B[69]:15,WTFAN:7:2,B[70]:7,WTFAB:7:2,B[71]:7,WTFOR:7:2);
   Writeln(ListFile,' ':45,B[72]:10,WFS:9:2,FS:9:2);
   Write(ListFile,B[73]:15,OXTOT:12:2,CTTOT:9:2,B[74]:19,WOL:9:2,OL:9:2);
   Writeln(ListFile,B[75]:15,TFAN:7:2,B[76]:7,TFAB:7:2,B[77]:7,TFOR:7:2);
   Writeln(ListFile,B[78]:15,NOH2O:12:2,B[79]:28,WFO:9:2,FO:9:2);
   Writeln(ListFile,' ':45,B[80]:10,WFA:9:2,FA:9:2,B[81]:35);
   Writeln(ListFile,' ':45,B[82]:10,WHO:9:2,HO:9:2);
   Writeln(ListFile,' ':45,B[83]:10,WRI:9:2,RI:9:2,B[84]:15,WAFMA:7:2,B[85]:7,WAFMF:7:2,B[86]:7,WAFMM:7:2);
   Writeln(ListFile,' ':45,B[87]:10,WACT:9:2,ACT:9:2);
   Writeln(ListFile,' ':45,B[88]:10,WED:9:2,ED:9:2,B[89]:15,AFMA:7:2,B[90]:7,AFMF:7:2,B[91]:7,AFMM:7:2);
   Writeln(ListFile,' ':45,B[92]:10,WSPIN:9:2,SPIN:9:2);
   Writeln(ListFile,' ':45,B[93]:10,WCS:9:2,CS:9:2,B[94]:29,WAGRAT:6:2,B[95]:11,AGRAT:6:2);
   Writeln(ListFile,' ':45,B[96]:10,WMT:9:2,MT:9:2);
   Writeln(ListFile,' ':45,B[97]:10,WCM:9:2,CM:9:2,B[98]:29,WFMRAT:6:2,B[99]:11,FMRAT:6:2);
   Writeln(ListFile,' ':45,B[100]:10,WHM:9:2,HM:9:2);
   Writeln(ListFile,' ':45,B[101]:10,WSP:9:2,SP:9:2,B[102]:29,WALRAT:6:2,B[103]:11,ALRAT:6:2);
   Writeln(ListFile,' ':45,B[104]:10,WRU:9:2,RU:9:2);
   Writeln(ListFile,' ':45,B[105]:10,WAP:9:2,AP:9:2,B[106]:29,WOXRAT:6:2,B[107]:11,OXRAT:6:2);
   Writeln(ListFile,' ':45,B[108]:10,WFL:9:2,Fl:9:2);
   Writeln(ListFile,' ':45,B[109]:10,WPY:9:2,PY:9:2,B[110]:27,WALKIN:8:2);
   Writeln(ListFile,' ':45,B[111]:10,WCC:9:2,CC:9:2,B[115]:27,WatM:8:2);
   Writeln(ListFile);
   Writeln(ListFile,' ':50,B[112]:5,WSALIC:9:2,SALIC:9:2);
   Writeln(ListFile,' ':50,B[113]:5,WFEMIC:9:2,FEMIC:9:2);
   Writeln(ListFile);
   Writeln(ListFile,' ':50,B[114]:5,WTOTAL:9:2,TOTAL:9:2);
   Writeln(ListFile,NormPrn);
   Writeln(ListFile,FF);
end;{proc MesoPrint}


procedure Record_To_Memory;
begin
   WAFMA:=1.0;
   WAFMM:=1.0;
   AFMA:=1.0;
   AFMF:=1.0;
   AFMM:=1.0;
   WAGRAT:=1.0;
   AGRAT:=1.0;
   WFMRAT:=1.0;
   FMRAT:=1.0;
   WALRAT:=1.0;
   ALRAT:=1.0;
   WOXRAT:=1.0;
   OXRAT:=1.0;
   OSI:=DataField.FileOxData[ElementPos[1]]*OxFactor[1];
   OTI:=DataField.FileOxData[ElementPos[2]]*OxFactor[2];
   OZR:=DataField.FileOxData[ElementPos[3]]*OxFactor[3];
   OAL:=DataField.FileOxData[ElementPos[4]]*OxFactor[4];
   OCR:=DataField.FileOxData[ElementPos[5]]*OxFactor[5];
   OF3:=DataField.FileOxData[ElementPos[6]]*OxFactor[6];
   OF2:=DataField.FileOxData[ElementPos[7]]*OxFactor[7];
   OMN:=DataField.FileOxData[ElementPos[8]]*OxFactor[8];
   ONI:=DataField.FileOxData[ElementPos[9]]*OxFactor[9];
   OMG:=DataField.FileOxData[ElementPos[10]]*OxFactor[10];
   OCA:=DataField.FileOxData[ElementPos[11]]*OxFactor[11];
   OSR:=DataField.FileOxData[ElementPos[12]]*OxFactor[12];
   OBA:=DataField.FileOxData[ElementPos[13]]*OxFactor[13];
   ONA:=DataField.FileOxData[ElementPos[14]]*OxFactor[14];
   OKO:=DataField.FileOxData[ElementPos[15]]*OxFactor[15];
   OPO:=DataField.FileOxData[ElementPos[16]]*OxFactor[16];
   OHP:=DataField.FileOxData[ElementPos[17]]*OxFactor[17];
   OH2OM:=DataField.FileOxData[ElementPos[18]]*OxFactor[18];
   OSO:=DataField.FileOxData[ElementPos[19]]*OxFactor[19];
   OSU:=DataField.FileOxData[ElementPos[20]]*OxFactor[20];
   OCL:=DataField.FileOxData[ElementPos[21]]*OxFactor[21];
   OFU:=DataField.FileOxData[ElementPos[22]]*OxFactor[22];
   OCD:=DataField.FileOxData[ElementPos[23]]*OxFactor[23];
   OXTOT:=OSI+OTI+OAL+OF3+OF2+OMN+OMG+OCA+ONA+OKO+OPO+OHP+OH2OM+OCD;
   OXTOT:=OXTOT+OCR+OSR+OBA+OCL+OFU+ONI+OZR+OSU+OSO;
end;{proc Record_To_Memory}


procedure SundryRatios;
var
   ALCA            : real;
begin
   if OXTOT>10 then
   begin
      NOH2O:=OXTOT-OHP-OH2OM;
      ALK:=ONA+OKO;
      F23:=OF2+OF3;
      if F23>0 then begin
         WOXRAT:=OF2*100/F23;
         F23:=OF2+OF3*0.8998;
         SUM:=100.0/(ALK+F23+OMG);
         WAFMA:=ALK*SUM;
         WAFMF:=F23*SUM;
         WAFMM:=OMG*SUM;
      end
      else begin
        WOXRAT:=0.0;
        WAFMA:=0.0;
        WAFMF:=0.0;
        WAFMM:=0.0;
      END;
      if ALK>0 then WALRAT:=ONA*100/ALK else WALRAT:=0.0;
      ALK1:=ALK;
      WALKIN:=0.0;
      if ONA>0.0 then OKONA:=OKO/ONA
                 else OKONA:=0.0;
      if ALK>0 then
      begin
         if (OSI > 50.0) and ((OKONA) > 1.0) and ((OKONA) < 2.5) then ALK1:=2.0*ONA;
         ALCA:=OAL+OCA;
         WALKIN:=(ALCA+ALK1)/(ALCA-ALK1);
      end;{if}
   end;{if}
   if (CSI*CAL > 0.0) then
        WatM := (CNA+CKO+2.0*CCA)/(CSI*CAL) * 100.0;
   if WatM>1e5 then WatM:=1e5;
end;{proc SundryRatios}


procedure CatEquiv;
{
Compute cation equivalents
}
begin
   SI:=OSI/60.0848;
   TI:=OTI/79.89881;
   ZR:=OZR/123.2188;
   AL:=OAL/50.9806;
   NCR:=OCR/75.9951;
   F3:=OF3/79.8461;
   F2:=OF2/71.8464;
   MN:=OMN/70.9374;
   NI:=ONI/74.7094;
   MG:=OMG/40.3114;
   CA:=OCA/56.0794;
   SR:=OSR/103.6194;
   BA:=OBA/153.3394;
   NA:=ONA/30.9895;
   KO:=OKO/47.1017;
   PO:=OPO/70.9723;
   SO:=OSO/80.0622;
   SU:=OSU/32.064;
   CL:=OCL/35.453;
   FU:=OFU/18.9984;
   CD:=OCD/44.00995;
end;{proc CatEquiv}


procedure Agpaitic;
{
Compute agpaitic ratio, etc
}
var
   FEMN          : real;
begin
   if AL>0.0 then WAGRAT:=(NA*22.9898+KO*39.102)/(AL*0.269815);
   FEMN:=(F2+F3)*0.55847+MN*0.54938;
   if ((FEMN>0.0) or (MG>0.0)) then WFMRAT:=100.0*FEMN/(FEMN+MG*0.24312);
   ALK:=(NA+KO)*0.5;
   F23:=F2+F3*0.5;
   SUM:=(ALK+F23+MG);
   if SUM>0.0 then begin
     SUM:=100.0/SUM;
     AFMA:=ALK*SUM;
     AFMF:=F23*SUM;
     AFMM:=MG*SUM;
   end;
   if F23>0.0 then OXRAT:=F2*100.0/F23 else OXRAT:=0.0;
   if ALK>0 then ALRAT:=NA*50.0/ALK else ALRAT:=0.0;
   if AL>0.0 then AGRAT:=ALK*200.0/AL;
   FEMN:=F2+F3+MN;
   if ((FEMN+MG)>0.0) then FMRAT:=FEMN*100.0/(FEMN+MG);
end;{proc Agpaitic}


procedure CatEquivPerc;
{
Compute cation equivalent percentages
}
begin
   CTL:=SI+TI+ZR+AL+F3+F2+MN+NI+MG+CA+SR+BA+NA+KO;
   CTL:=CTL+PO+SO+SU+CL+FU+CD;
   if CTL<=0.0 then CTL:=1.0;
   F:=100.0/CTL;
   SI:=SI*F;
   CSI:=SI;
   TI:=TI*F;
   CTI:=TI;
   AL:=AL*F;
   CAL:=AL;
   ZR:=ZR*F;
   CZR:=ZR;
   NCR:=NCR*F;
   CCR:=NCR;
   F3:=F3*F;
   CF3:=F3;
   F2:=F2*F;
   CF2:=F2;
   CMN:=MN*F;
   CNI:=NI*F;
   F2:=CF2+CMN+CNI;
   MG:=MG*F;
   CMG:=MG;
   CCA:=CA*F;
   CSR:=SR*F;
   CBA:=BA*F;
   CA:=CCA+CSR+CBA;
   NA:=NA*F;
   CNA:=NA;
   KO:=KO*F;
   CKO:=KO;
   PO:=PO*F;
   CPO:=PO;
   SO:=SO*F;
   CSO:=SO;
   SU:=SU*F;
   CSU:=SU;
   CL:=CL*F;
   CCL:=CL;
   FU:=FU*F;
   CFU:=FU;
   CD:=CD*F;
   CCD:=CD;
   CTTOT:=SI+TI+ZR+AL+NCR+F3+F2+MG+CA+NA+KO+PO+SO+CD;
end;{proc CatEquivPerc}

procedure SetMinZero;
{------------------------------------------------------------
 Set minerals to 0
 ------------------------------------------------------------}
begin
   QZ:=0.0;
   CO:=0.0;
   ZN:=0.0;
   ORT:=0.0;
   AB:=0.0;
   AN:=0.0;
   LC:=0.0;
   NE:=0.0;
   KP:=0.0;
   HL:=0.0;
   TH:=0.0;
   AC:=0.0;
   NS:=0.0;
   KS:=0.0;
   WO:=0.0;
   DI:=0.0;
   HY:=0.0;
   OL:=0.0;
   CS:=0.0;
   MT:=0.0;
   CM:=0.0;
   IL:=0.0;
   HM:=0.0;
   SP:=0.0;
   PF:=0.0;
   RU:=0.0;
   RU:=0.0;
   AP:=0.0;
   FL:=0.0;
   PY:=0.0;
   CC:=0.0;
   PMG:=0.0;
   PF2:=0.0;
   HO:=0.0;
   ACT:=0.0;
   RI:=0.0;
   ED:=0.0;
   BI:=0.0;
   SPIN:=0.0;
   TOTAL:=0.0;

   WQZ:=0.0;
   WCO:=0.0;
   WZN:=0.0;
   WOR:=0.0;
   WAB:=0.0;
   WAN:=0.0;
   WLC:=0.0;
   WNE:=0.0;
   WKP:=0.0;
   WHL:=0.0;
   WTH:=0.0;
   WAC:=0.0;
   WNS:=0.0;
   WKS:=0.0;
   WWO:=0.0;
   WDI:=0.0;
   WHY:=0.0;
   WOL:=0.0;
   WCS:=0.0;
   WMT:=0.0;
   WCM:=0.0;
   WIL:=0.0;
   WHM:=0.0;
   WSP:=0.0;
   WPF:=0.0;
   WRU:=0.0;
   WAP:=0.0;
   WFL:=0.0;
   WPY:=0.0;
   WCC:=0.0;
   WHO:=0.0;
   WBI:=0.0;
   WACT:=0.0;
   WED:=0.0;
   WRI:=0.0;
   WSPIN:=0.0;
   WTOTAL:=0.0;
end;{proc SetMinZero}


procedure CatToWt;
{
Recalc. cation equiv. norm in wt. percent
}
begin
   F:=CTL*0.01;
   WQZ:=QZ*60.0848*F;
   WCO:=CO*50.9806*F;
   WZN:=ZN*91.6518*F;
   WOR:=ORT*55.6673*F;
   WAB:=AB*52.4449*F;
   WAN:=AN*55.642*F;
   WPL:=WAB+WAN;
   WLC:=LC*54.563*F;
   WNE:=NE*47.3516*F;
   WKP:=KP*52.7224*F;
   WHL:=HL*58.4428*F;
   WTH:=TH*47.3471*F;
   WAC:=AC*57.7513*F;
   WNS:=NS*40.687*F;
   WKS:=KS*51.4294*F;
   WWO:=WO*58.0821*F;
   WWODI:=DI*29.0411*F;
   WENDI:=DI*PMG*25.0991*F;
   WFSDI:=DI*PF2*32.9828*F;
   WDI:=WWODI+WENDI+WFSDI;
   WEN:=HY*PMG*50.1981*F;
   WFS:=HY*PF2*65.9656*F;
   WHY:=WEN+WFS;
   WFO:=OL*PMG*46.9025*F;
   WFA:=OL*PF2*67.92591*F;
   WOL:=WFO+WFA;
   WBI:=(BI*PMG*49.90886+BI*PF2*61.73449)*F;
   WACT:=(ACT*PMG*52.9596+ACT*PF2*63.47128)*F;
   WED:=(ED*PMG*51.01747+ED*PF2*60.8722)*F;
   WRI:=RI*61.19259*F;
   WSPIN:=SPIN*47.4242*F;
   WHO:=WACT+WED+WRI;
   WCS:=CS*57.4151*F;
   WMT:=MT*77.1795*F;
   WCM:=CM*74.6122*F;
   WIL:=IL*75.8726*F;
   WHM:=HM*79.8461*F;
   WSP:=SP*65.3543*F;
   WPF:=PF*67.9891*F;
   WRU:=RU*79.89881*F;
   if (OFU > 0.0) then WAP:=AP*63.0391*F
                  else WAP:=AP*61.6642*F;
   WFL:=FL*78.0768*F;
   WPY:=PY*119.975*F;
   WCC:=CC*50.044675*F;
   WODI:=0.5*DI;
   ENDI:=0.5*DI*PMG;
   FSDI:=WODI-ENDI;
   FO:=OL*PMG;
   FA:=OL-FO;
   EN:=HY*PMG;
   FS:=HY-EN;
   NPL:=AB+AN;
end;{proc CatToWt}


procedure MinComp;
begin
   ANPL:=0.0;
   FAOL:=0.0;
   ENHY:=0.0;
   WANPL:=0.0;
   WFAOL:=0.0;
   WENHY:=0.0;
   { An content of plag. }
   if NPL>0.0 then
   begin
      ANPL:=100.0*AN/NPL;
      WANPL:=100.0*WAN/WPL;
   end;{if}
   { Fa content of oliv. }
   if OL>0.0 then
   begin
      FAOL:=100.0*PF2;
      WFAOL:=100.0*WFA/WOL;
   end;{if}
   { En content of hypersth. }
   if HY>0.0 then
   begin
      ENHY:=100.0*PMG;
      WENHY:=100.0*WEN/WHY;
   end;{if}
end;{proc MinComp}


procedure DifIndex;
{
Thornton-Tuttle differ. index
}
begin
   DIX:=QZ+ORT+AB+NE+LC+KP;
   WDIX:=WQZ+WOR+WAB+WNE+WLC+WKP;
end;{proc DifIndex}


procedure Totals;
{
Totals for salic, femic and all minerals
}
begin
   SALIC:=DIX+CO+ZN+AN+HL+TH;
   FEMIC:=AC+NS+KS+WO+DI+HY+OL+BI+HO+SPIN+CS+MT+CM;
   FEMIC:=FEMIC+IL+HM+SP+PF+RU+AP+FL+PY+CC;
   TOTAL:=FEMIC+SALIC;
   WSALIC:=WDIX+WCO+WZN+WAN+WHL+WTH;
   WFEMIC:=WAC+WNS+WKS+WWO+WDI+WHY+WOL+WBI+WHO+WSPIN+WCS+WMT+WCM;
   WFEMIC:=WFEMIC+WIL+WHM+WSP+WPF+WRU+WAP+WFL+WPY+WCC;
   WTOTAL:=WFEMIC+WSALIC;
end;{proc Totals}


procedure CatToWt100;
{
Recalc. wt. % norm ito 100 wt. percent
}
begin
   F:=100.0/WTOTAL;
   WQZ:=WQZ*F;
   WCO:=WCO*F;
   WZN:=WZN*F;
   WOR:=WOR*F;
   WAB:=WAB*F;
   WAN:=WAN*F;
   WPL:=WAB+WAN;
   WLC:=WLC*F;
   WNE:=WNE*F;
   WKP:=WKP*F;
   WHL:=WHL*F;
   WTH:=WTH*F;
   WAC:=WAC*F;
   WNS:=WNS*F;
   WKS:=WKS*F;
   WWO:=WWO*F;
   WWODI:=WWODI*F;
   WENDI:=WENDI*F;
   WFSDI:=WFSDI*F;
   WDI:=WWODI+WENDI+WFSDI;
   WEN:=WEN*F;
   WFS:=WFS*F;
   WHY:=WEN+WFS;
   WFO:=WFO*F;
   WFA:=WFA*F;
   WOL:=WFO+WFA;
   WBI:=WBI*F;
   WACT:=WACT*F;
   WED:=WED*F;
   WRI:=WRI*F;
   WSPIN:=WSPIN*F;
   WHO:=WACT+WED+WRI;
   WCS:=WCS*F;
   WMT:=WMT*F;
   WCM:=WCM*F;
   WIL:=WIL*F;
   WHM:=WHM*F;
   WSP:=WSP*F;
   WPF:=WPF*F;
   WRU:=WRU*F;
   WAP:=WAP*F;
   WFL:=WFL*F;
   WPY:=WPY*F;
   WCC:=WCC*F;
   MinComp;
   DifIndex;
   WSALIC:=WDIX+WCO+WZN+WAN+WHL+WTH;
   WFEMIC:=WAC+WNS+WKS+WWO+WDI+WHY+WOL+WBI+WHO+WSPIN+WCS+WMT+WCM;
   WFEMIC:=WFEMIC+WIL+WHM+WSP+WPF+WRU+WAP+WFL+WPY+WCC;
   WTOTAL:=WFEMIC+WSALIC;
end;{proc CatToWt100}


procedure CalcCIPW;
begin
   Item:=0;
   Names;
   for I:=Nbeg TO Nfin do begin
      SetMinZero;      {set minerals to 0}
      {$I-}
      Seek(DataFile,I);        IOCheck;
      Read(DataFile,DataField);  IOCheck;
      {$I+}
      Record_To_Memory;
      DES:=DataField.SmpDescription+StringOf(64,32);
      Item:=Item+1;
      At(24,1);
      Write('CIPW-norm        Item ',Item:3,'             Record ',(I-2):3,
            '           Sample ',DataField.SampleNo:10);
      At(25,40);
      CatEquiv;
      Agpaitic;
      CatEquivPerc;
      SundryRatios;
      if OXTOT>10.0 then begin
         Apatite;
         Halite;
         Thenardite;
         Pyrite;
         Chromite;
         Ilmenite;
         Fluorite;
         Calcite;
         Zircon;
         Orthoclase;  {Orthoclase and Potassium metasilicate}
         Albite;
         AnCor;       {Anorthite and Corundum}
         Acmite;      {Acmite and Sodium metasilicate}
         SphRut;      {Sphene and Rutile}
         MtHm;        {Magnetite and Hematite}
         F2MgRatio;
         WollHyp;     {Wollastonite and Hypersthene}
         if HY>0.0 then Diopside;
         if SI < 0.0 then DesilHy;    {Test for Si deficiency}
         if SI < 0.0 then DesilSp;    {Test for Si deficiency}
         if SI < 0.0 then DesilAb;    {Test for Si deficiency}
         if SI < 0.0 then DesilOr;    {Test for Si deficiency}
         if SI < 0.0 then DesilWo;    {Test for Si deficiency}
         if SI < 0.0 then DesilDi;    {Test for Si deficiency}
         if SI < 0.0 then DesilLc;    {Test for Si deficiency}
         MakeQz;
         CatToWt;
         MinComp;
         DifIndex;
         Totals;
         if Toggle100=1 then
         begin
           CatToWt100;
         end;
         OrAbAn;
         QzNeKp;
         QzAbOr;
         Draw_Block( 9,15,11,65,1,Norm);
         if OutRoute='P' then
         begin
           Centre(10,'... Printing data ...',Norm);
           At(25,40);
           CheckPrinter;
           if not Prndone then
           begin
             Draw_Block( 9,15,11,65,1,Norm);
             Centre(10,'... Printing data ...',Norm);
             At(25,40);
             CIPWPrint;
           end;
         end;
         if OutRoute='F' then
         begin
           Centre(10,'... Generating file ...',Norm);
           At(25,40);
           OutToASCII(44);
         end;
      end;{if OXTOT}
   end;{for}
end;{proc CalcCIPW}


procedure CalcMeso;
begin
   Item:=0;
   Names;
   for I:=Nbeg TO Nfin do begin
      SetMinZero;      {set minerals to 0}
      {$I-}
      Seek(DataFile,I);        IOCheck;
      Read(DataFile,DataField);  IOCheck;
      {$I+}
      Record_To_Memory;
      Item:=Item+1;
      DES:=DataField.SmpDescription+StringOf(64,32);
      At(24,1);
      Write('Meso-norm        Item ',Item:3,'             Record ',(I-2):3,
            '           Sample ',DataField.SampleNo:10);
      At(25,40);
      CatEquiv;
      Agpaitic;
      CatEquivPerc;
      SundryRatios;
      if OXTOT>10.0 then begin
         Calcite;
         Apatite;
         Fluorite;
         Halite;
         Thenardite;
         Pyrite;
         Chromite;
         Zircon;
         SphRut;      {Sphene and Rutile}
         Orthoclase;  {Orthoclase and Potassium metasilicate}
         Albite;
         AnCor;       {Anorthite and Corundum}
         Riebeckite;  {Riebeckite and Sodium metasilicate}
         MtHm;        {Magnetite and Hematite}
         F2MgRatio;
         Biotite;
         if InOption='2' then
         begin
            ActHyWo;
         end
         else begin
            WollHyp;
            Diopside;
         end;{if}
         if SI < 0.0 then DesilAct;   {Test for Si deficiency}
         if SI < 0.0 then DesilHy;    {Test for Si deficiency}
         if SI < 0.0 then DesilOl;    {Test for Si deficiency}
         if SI < 0.0 then DesilAb;    {Test for Si deficiency}
         MakeQz;
         Hornblende;
         CatToWt;
         MinComp;
         DifIndex;
         Totals;
         if Toggle100=1 then
         begin
           CatToWt100;
         end;
         OrAbAn;
         QzNeKp;
         QzAbOr;
         Draw_Block( 9,15,11,65,1,Norm);
         if OutRoute='P' then
         begin
           Centre(10,'... Printing data ...',Norm);
           At(25,40);
           CheckPrinter;
           if not Prndone then
           begin
             Draw_Block( 9,15,11,65,1,Norm);
             Centre(10,'... Printing data ...',Norm);
             At(25,40);
             MesoPrint;
           end;
         end;
         if OutRoute='F' then
         begin
           Centre(10,'... Generating file ...',Norm);
           At(25,40);
           OutToASCII(44);
         end;
      end;{if OXTOT}
   end;{for}
end;{proc CalcMeso}

procedure ShowToggle100;
begin
  PrintS(24,10,StringOf(60,196),NormBld);
  PrintS(25,1,'F5 = toggle recalc. option',RevNorm);
  if OutRoute='P' then
    PrintS(25,27,' ³ Send to printer - F2 to change ',RevNorm)
  else
    PrintS(25,27,' ³ Send to file   - F2 to change  ',RevNorm);
  ChangeAtt(25,1,79,RevNorm);
  case Toggle100 of
    1 : begin
      At(24,1); ClrEOL;
      Centre(24,'Mineral wt % recalc. to 100 %',Norm);
    end;
    0 : begin
      At(24,1); ClrEOL;
      Centre(24,'Calculated mineral wt % not recalc. for 100 %',Norm);
    end;
  end;
end;

procedure ShowMenu1;
begin
      ClrScr;
      ShowToggle100;
      PrintS( 2,10,StringOf(60,205),NormBld);
      PrintS(18,10,StringOf(60,205),NormBld);
      Centre(2,'   Options Menu   ',RevNorm);
      PrintS( 4,10,'       [ Q ]  Terminate',NormBld);
      PrintS( 6,10,'       [ N ]  New file ',NormBld);
      PrintS( 8,10,'       [ P ]  Prepare NORMS file structure ',NormBld);
      PrintS(10,10,'       [ 0 ]  CIPW norm',NormBld);
      PrintS(12,10,'       [ 1 ]  Meso-norm with biotite and pyroxene  ',NormBld);
      PrintS(14,10,'       [ 2 ]                 biotite and hornblende',NormBld);
      PrintS(16,10,'    Enter your choice: ',NormBld);
end;


begin
  CheckSystem(BombChar,LocalYearStr,LocalMonthStr,LocalDayStr,LocalProgName+'.EXE',LocalUserCode);
  ProgramName:=LocalProgName;
  ProgVersion:=LocalUserCode;
   Toggle100:=0;
   DES:=' ';
   A14:='          Species     Input     Calc     Comp     Mineral    Wt       Cat     ';
   A15:='                       Wt        Cat     Error              Prcnt    Prcnt    ';
   A16:='                      Prcnt     Prcnt    Ctpcnt      ';
   done:=False;
   OpenBanner;
   SelectFile1;
   OutRoute:='P';
   repeat
      repeat
        ShowMenu1;
        At(16,35); GetKey(ComplexKey,KeyCode);
        { ReadKbd(InOption);
        InOption:=UpCase(InOption); }
        if not ComplexKey then
           InOption:=UpCase(Chr(KeyCode))
        else begin
          case KeyCode of
            f5key : begin
              if Toggle100=0 then Toggle100:=1
                             else Toggle100:=0;
            end;
            f2key : begin
              if OutRoute='P' then OutRoute:='F'
                              else OutRoute:='P';
              if OutRoute='P' then
              begin
                {$I-}
                Close(TextFile1); IOCheck;
                {$I+}
              end;
              if OutRoute='F' then
              begin
                GetTargetFileName;
                if TFileName='ABORT' then
                begin
                  OutRoute:='P';
                  ShowMenu1;
                end
                else begin
                  MineralNames;
                  HeaderOutToASCII(44);
                end;
              end;
            end;
          end;
          ClrScr;
          InOption:='?';
        end;
      until (InOption in ['Q','N','P','0','1','2']);
      if ((InOption='0') or (InOption='1') or (InOption='2')) then begin
         Nbeg:=1;
         Nfin:=TotalRecs;
         ClrScr;
         PrintS( 3,10,StringOf(60,205),NormBld);
         PrintS( 9,10,StringOf(60,205),NormBld);
         PromptInt( 5,10,'Record no. to begin   ###',Norm,Norm,Nbeg);
         PromptInt( 7,10,'Record no. to end     ###',Norm,Norm,Nfin);
         PrintS(25,1,'F10 - complete |',RevNorm);
         ChangeAtt(25,1,79,RevNorm);
         L:=1;
         LMin:=1;
         LMax:=2;
         repeat
           case L of
             1 : begin
               TCI:=0;
               CPos:=1;
               ReadInt( 5,32,'###',Norm,RevNorm,1,TotalRecs,
                        CPos,TCI,Nbeg);
             end;
             2 : begin
               TCI:=0;
               CPos:=1;
               ReadInt( 7,32,'###',Norm,RevNorm,0,TotalRecs,
                        CPos,TCI,Nfin);
             end;
           end;
           case TCI of
             arrowdnkey, enterkey : begin
               L:=L+1;
               if L>LMax then L:=LMin;
             end;
             arrowupkey : begin
               L:=L-1;
               if L<LMin then L:=LMax;
             end;
           end;
         until ((TCI = f10key) or (TCI = esckey));
         if TCI=esckey then InOption:='?';
         if Nfin<1 then Nfin:=TotalRecs;
         Nfin:=Nfin+2;
         if Nfin>RecCount then Nfin:=RecCount;
         Nbeg:=Nbeg+2;
      end;{if InOption}
      ClrScr;
      case InOption of
        'Q' : begin
                 done:=True;
                 {I-}
                 if FileName <> 'ABORT' then
                 begin
                   Close(DataFile);   IOCheck;
                 end;
                 {I+}
                 if OutRoute='F' then
                 begin
                   {$I-}
                   if FileName <> 'ABORT' then
                   begin
                     Close(TextFile1); IOCheck;
                   end;
                   {$I+}
                 end;
              end;
        'N' : begin
                 {I-}
                 Close(DataFile);   IOCheck;
                 {I+}
                 if OutRoute='F' then
                 begin
                   OutRoute:='P';
                   {$I-}
                   Close(TextFile1); IOCheck;
                   {$I+}
                 end;
                 SelectFile1;
              end;
        'P' : begin
                 PrepareNormsFile;
              end;
        '0' : begin
                 GetElementOrder;
                 if FilePrepared then
                 begin
                   Draw_Block(12,23,14,56,1,NormBld);
                   Centre(13,'... Busy processing data ...',NormFlik);
                   CalcCIPW;
                 end;
              end;
        '1','2' : begin
                 GetElementOrder;
                 if FilePrepared then
                 begin
                   Draw_Block(12,23,14,56,1,NormBld);
                   Centre(13,'... Busy processing data ...',NormFlik);
                   CalcMeso;
                 end;
              end;
      end;{case}
   until done;
   Terminate;
end.
