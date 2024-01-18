program Irvine;

uses
  crt, dos,
  printer3,
  sys_def,
  sys_def2,
  allsorts,
  xvideo,
  normvarb,
  kem_varb;

var
  IrvMg, IrvAl : real;
  TextFile     : text;

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

procedure CalcIrv;
var
  TotFe : real;
begin
   Item:=0;
   for I:=Nbeg TO Nfin do begin
      {$I-}
      Seek(DataFile,I);        IOCheck;
      Read(DataFile,DataField);  IOCheck;
      {$I+}
      Record_To_Memory;
      DES:=DataField.SmpDescription+StringOf(64,32);
      Item:=Item+1;
      At(24,1);
      Write('Irvine values    Item ',Item:3,'             Record ',(I-2):3,
            '           Sample ',DataField.SampleNo:10);
      At(25,40);
      CatEquiv;
      CatEquivPerc;
      (*
      if (F2 > 0.1) then
      begin
        IrvMg:=100.0*MG/(MG+F2);
        IrvAl:=100.0*Al/(MG+F2);
      end
      else begin
        IrvMg:=100.0*MG/(MG+F2+F3);
        IrvAl:=100.0*Al/(MG+F2+F3);
      end;
      *)
      TotFe:=F2+F3;
      F3:=Ti*0.5;
      F2:=TotFe-F3;
      IrvMg:=100.0*MG/(MG+F2);
      IrvAl:=100.0*Al/(MG+F2);
      Writeln(TextFile,'"#"   ',IrvMg:8:2,'  ',IrvAl:8:2,'     ',DataField.SampleNo:10);
   end;{for}
end;{proc CalcIrv}

begin
   done:=False;
   Assign(DataFile,'C:\bme\kd\equeefa.kem');
   Reset(DataFile);
   Assign(TextFile,'C:\g\equeefa\eq_25.dat');
   Rewrite(TextFile);
   Nbeg:=3;
   Nfin:=35;
   GetElementOrder;
   CalcIrv;
   Close(DataFile);
   Close(TextFile);
   Terminate;
end.



