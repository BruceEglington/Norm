unit Nor_Prep;

interface
uses
  crt,
  sys_def, inputset,
  findfile, filselec,
  xvideo, kem_varb;

procedure PrepareNormsFile;


implementation

procedure PrepareNormsFile;
var
  done  : boolean;
  TCI   : integer;
  I     : integer;
begin
  done:=false;
  SFileName:='        ';
  SExtension:='.KEM';
  TFileName:='         ';
  TExtension:='.KUN';
  ClrScr;
  SelectFile(1,1,'Data file',138,'N',Drive2,SFileName,SExtension,DataFile1);
  if (TCI=esckey) then begin
    done:=true;
    SFileName:='ABORT';
  end;
  if not done then
  begin
    SDirec:=Drive2;
    SelectFileT(1,1,'Custom file',138,'N',SDirec,TFileName,TExtension,TextFile1);
    if (TCI=esckey) then begin
      done:=true;
      TFileName:='ABORT';
    end;
  end;
  if not done then begin
        Seek(DataFile1,1);
        Read(DataFile1,DataField);
        for I:=1 to 25 do
        begin
          Readln(TextFile1,DataField.Pos25[I]);
        end;
        Datafield.NormFlag:=1;
        Seek(DataFile1,1);
        Write(DataFile1,DataField);
  end;
  if ((not done)) then begin
    Close(DataFile1);
    ClrScr;
    Draw_Block(11,10,13,70,2,RevNorm);
    Centre(12,'File '+SFileName+' prepared for NORMS',NormBld);
    At(25,40);
    Delay(1500);
    ClrScr;
  end;
end;


end.