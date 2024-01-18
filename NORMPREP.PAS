program NormPrep;
uses
  crt, dos, sys_def,
  xvideo, allsorts,
  readsys, inputset,
  kem_varb, filselec,
  findfile,
  getfield, protect,
  nor_prep;

const
   LocalProgName     = 'NORMPREP.EXE';
   LocalUserCode     = '2.0:000:301089';
   LocalYearStr      = '1990';
   LocalMonthStr     =   '12';
   LocalDayStr       =   '31';


var
  Num, I      : integer;



procedure FileBanner(HoldIt : boolean);
const
  ListLines = 14;
  StartLine = 5;
  List : array[1..ListLines] of string =
    ('           ',
    'A routine to set the positions of elements for NORMS ',
    '     ',
    'NORMPREP has the following features: ',
    '        þ programs such as NORMS need to know where ',
    '          specific elements are located in KEMFILE type',
    '          data files                           ',
    '        þ the program reads a .KUN custom file and',
    '          stores the relevant information in the chosen',
    '          data file.  ',
    '          ',
    '                                 ',
    '                                    ',
    '                                 ');
var
 i : byte;

begin
    ClrScr;
    OffScr;
    Draw_Block(2,30,4,50,2,NormBld);
    Centre(3,'NORMPREP',NormBld);
    Draw_Block(StartLine,5,ListLines+StartLine+2,75,1,NormBld);
    for i:=1 to ListLines do
      PrintS(i+StartLine,10,List[i],Norm);
    Centre(20,'Version '+ProgVersion,NormBld);
    Centre(25,'[Press Space Bar to Continue]',NormBld);
    OnScr;
    At(25,40);
    if HoldIt then
       begin
        repeat until Ord(ReadKey)=32; ClearBlock(25,10,25,60,true);
       end;
end;

procedure OpenBanner;
begin
  FileBanner(true);
end;{proc OpenBanner}



begin
  CheckSystem(LocalYearStr,LocalMonthStr,LocalDayStr,LocalProgName,LocalUserCode);
  ProgVersion:=LocalUserCode;
  OpenBanner;
  PrepareNormsFile;
end.
