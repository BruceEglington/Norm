unit Nor_Asci;

interface
uses
  crt, dos,
  sys_def, inputset,
  allsorts, xvideo,
  findfile, filselec,
  normvarb, kem_varb;

var
  MinName  : array[0..44] of string[6];

procedure GetTargetFileName;
procedure HeaderOutToASCII(numtoprint : byte);
procedure OutToASCII(numtoprint : byte);
procedure MineralNames;

implementation

var
  MinVal  : array[0..44] of real;

procedure GetTargetFileName;
begin
  TExtension:='.DAT';
  ClrScr;
  SelectFileT( 1,1,'Target File',10,'Y',Drive3,TFileName,TExtension,TextFile1);
  if ((TCI=esckey) or (TFileName='ABORT')) then begin
    TFileName:='ABORT';
  end;
  TCI:=0;
  ClrScr;
end;

procedure MineralNames;
begin
      MinName[1]:='Qz  ';
      MinName[2]:='Co  ';
      MinName[3]:='Z   ';
      MinName[4]:='Or  ';
      MinName[5]:='Pl  ';
      MinName[6]:='Pl(Ab)';
      MinName[7]:='Pl(An)';
      MinName[8]:='Lc  ';
      MinName[9]:='Ne  ';
      MinName[10]:='Kp  ';
      MinName[11]:='Hl  ';
      MinName[12]:='Th  ';
      MinName[13]:='Ac  ';
      MinName[14]:='Ns  ';
      MinName[15]:='Ks  ';
      MinName[16]:='Wo  ';
      MinName[17]:='Di  ';
      MinName[18]:='Di(Wo)';
      MinName[19]:='Di(En)';
      MinName[20]:='Di(Fs)';
      MinName[21]:='Hy  ';
      MinName[22]:='Hy(En)';
      MinName[23]:='Hy(Fs)';
      MinName[24]:='Ol  ';
      MinName[25]:='Ol(Fo)';
      MinName[26]:='Ol(Fa)';
      MinName[27]:='Cs  ';
      MinName[28]:='Mt  ';
      MinName[29]:='Cm  ';
      MinName[30]:='Il  ';
      MinName[31]:='Hm  ';
      MinName[32]:='Sp  ';
      MinName[33]:='Pf  ';
      MinName[34]:='Ru  ';
      MinName[35]:='Ap  ';
      MinName[36]:='Fl  ';
      MinName[37]:='Py  ';
      MinName[38]:='Cc  ';
      MinName[39]:='Bi  ';
      MinName[40]:='Ho  ';
      MinName[41]:='Ho(Act)';
      MinName[42]:='Ho(Ed)';
      MinName[43]:='Ho(Ri)';
      MinName[44]:='Spnl';
end;

procedure MineralPercent;
begin
      MinVal[1]:=WQz;
      MinVal[2]:=WCo;
      MinVal[3]:=WZn;
      MinVal[4]:=WOr;
      MinVal[5]:=WPl;
      MinVal[6]:=WAb;
      MinVal[7]:=WAn;
      MinVal[8]:=WLc;
      MinVal[9]:=WNe;
      MinVal[10]:=WKp;
      MinVal[11]:=WHl;
      MinVal[12]:=WTh;
      MinVal[13]:=WAc;
      MinVal[14]:=WNs;
      MinVal[15]:=WKs;
      MinVal[16]:=WWo;
      MinVal[17]:=WDi;
      MinVal[18]:=WWoDi;
      MinVal[19]:=WEnDi;
      MinVal[20]:=WFsDi;
      MinVal[21]:=WHy;
      MinVal[22]:=WEn;
      MinVal[23]:=WFs;
      MinVal[24]:=WOl;
      MinVal[25]:=WFo;
      MinVal[26]:=WFa;
      MinVal[27]:=WCs;
      MinVal[28]:=WMt;
      MinVal[29]:=WCm;
      MinVal[30]:=WIl;
      MinVal[31]:=WHm;
      MinVal[32]:=WSp;
      MinVal[33]:=WPf;
      MinVal[34]:=WRu;
      MinVal[35]:=WAp;
      MinVal[36]:=WFl;
      MinVal[37]:=WPy;
      MinVal[38]:=WCc;
      MinVal[39]:=WBi;
      MinVal[40]:=WHo;
      MinVal[41]:=WAct;
      MinVal[42]:=WEd;
      MinVal[43]:=WRi;
      MinVal[44]:=WSpin;
end;

procedure HeaderOutToASCII(numtoprint : byte);
var
  j        : byte;
  AnyStr   : string[10];
begin
   Write(TextFile1,'Sample#   ');
   for j:=2 to numtoprint+1 do
   begin
    AnyStr:=MinName[j-1];
    ClearNull(AnyStr);
    AnyStr:=AnyStr+StringOf(10-Length(AnyStr),32);
    Write(TextFile1,AnyStr);
    if (j mod 8)=0 then Writeln(TextFile1);
   end;
   if ((numtoprint mod 8)>0) or (numtoprint<8) then Writeln(TextFile1);
end;

procedure OutToASCII ( numtoprint : byte);
var
  j        : byte;
  AnyReal  : real;
begin
   MineralPercent;
   Write(TextFile1,DataField.SampleNo);
   Write(TextFile1,StringOf(10-Length(DataField.SampleNo),32));
   for j:=2 to numtoprint+1 do
   begin
     AnyReal:=MinVal[j-1];
     Write(TextFile1,AnyReal:10:2);
     if (j mod 8)=0 then Writeln(TextFile1);
   end;
   if ((numtoprint mod 8)>0) or (numtoprint<8) then Writeln(TextFile1);
end;


end.