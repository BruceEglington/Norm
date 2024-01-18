unit Nrm_ShtIm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles, 
  VCL.FlexCel.Core, FlexCel.XlsAdapter,
  System.Generics.Collections,
  Grids, DBGrids, DBCtrls, AxCtrls, Data.DB, System.ImageList, Vcl.ImgList,
  Vcl.VirtualImageList;

type
  TfmSheetImport = class(TForm)
    Panel1: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    bbCancel: TBitBtn;
    Splitter1: TSplitter;
    pData: TPanel;
    TabControl1: TTabControl;
    SheetData: TDrawGrid;
    Panel3: TPanel;
    gbDefineFields: TGroupBox;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    Memo1: TMemo;
    Panel2: TPanel;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label1: TLabel;
    Label4: TLabel;
    meFromRow: TEdit;
    meToRow: TEdit;
    bbImport: TBitBtn;
    VirtualImageList1: TVirtualImageList;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SheetDataDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure TabControl1Change(Sender: TObject);
  private
    { Private declarations }
    MyActiveSheetNum : integer;
    procedure OpenFile(const FileName: string);
    procedure FillTabControl;
    procedure FillGrid;
    function GetCellValue(const aCol, aRow: integer): string;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure ConvertDataToOxides;
    procedure GetElementOrder;
    procedure MatchElementsInFile;
  public
    { Public declarations }
  end;

var
  fmSheetImport: TfmSheetImport;

implementation

{$R *.DFM}

uses
  AllSorts, normvarb, Nrm_dm_acs;

var
  iRec, iRecCount      : integer;

procedure TfmSheetImport.bbOpenSheetClick(Sender: TObject);
var
  pFileType : smallint;
  pBuf      : string;
  pTitle    : string;
  tmpStr    : string[3];
begin
  OpenDialogSprdSheet.InitialDir := DataPath;
  if not OpenDialogSprdSheet.Execute then Exit;
  DataPath := ExtractFilePath(OpenDialogSprdSheet.FileName);
  OpenFile(OpenDialogSprdSheet.FileName);
end;

function TfmSheetImport.ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
    Result := 0;
    if (length(tmpStr) = 2) then
    begin
      tmpChar := tmpStr[1];
      itmp := (ord(tmpChar)-64)*26;
      tmpChar := tmpStr[2];
      Result := itmp+(ord(tmpChar)-64);
    end else
    begin
      tmpChar := tmpStr[1];
      Result := (ord(tmpChar)-64);
    end;
end;

procedure TfmSheetImport.OpenFile(const FileName: string);
var
  //StartOpen: TDateTime;
  //EndOpen: TDateTime;
  //StartSheetSelect, EndSheetSelect: TDateTime;
  Xls: TExcelFile;
  CellReader: TCellReader;
  Formatted : boolean;
begin
  pData.Visible := true;
  Formatted := false;
  //Open the Excel file.
  Xls := TXlsFile.Create(true);
  try
    xls.IgnoreFormulaText := true; //bme - hard code this for this situation since just reading cell values
    xls.VirtualMode := true;
    CellReader := TCellReader.Create(ShowOnly50Rows, CellData, Formatted);
    try
      xls.VirtualCellStartReading := CellReader.OnStartReading;
      xls.VirtualCellRead := CellReader.OnCellRead;
      xls.Open(FileName);
    finally
      CellReader.Free;
    end;
    FillTabControl;
    TabControl1.TabIndex := xls.ActiveSheet - 1;
    FillGrid;
  finally
    Xls.Free;
  end;
  gbDefineRows.Visible := true;
  bbImport.Visible := true;
  gbDefineFields.Visible := true;
end;

procedure TfmSheetImport.SheetDataDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  SheetData.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, GetCellValue(ACol, ARow));
end;

procedure TfmSheetImport.TabControl1Change(Sender: TObject);
begin
  FillGrid;
  //ShowMessage('Active sheet is number '+Int2Str(MyActiveSheetNum));
end;

procedure TfmSheetImport.FillTabControl;
var
  i: Integer;
begin
  TabControl1.Tabs.Clear;
  for i := 0 to CellData.Count - 1 do
  begin
    TabControl1.Tabs.Add(CellData[i].SheetName);
  end;
end;

procedure TfmSheetImport.bbImportClick(Sender: TObject);
var
  j     : integer;
  iCode : integer;
  i : integer;
  FromRow, ToRow : integer;
  tmpStr : string;
  Xls: TExcelFile;
  Formatted : boolean;
  v : TCellValue;
  tmpDataStr : string;
  tmpData : double;
  WasSuccessful : boolean;
  tmpDataValue : double;
begin
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
      tmpStr := meToRow.Text;
      Val(tmpStr, ToRow, iCode);
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
    if (iCode = 0) then
    begin
      if (ToRow >= FromRow) then iCode := 0
                            else iCode := -1;
    end else
    begin
      ShowMessage('Incorrect value entered for To row');
      Exit;
    end;
    if (iCode <> 0)
      then begin
        ShowMessage('Incorrect values entered for rows to import');
        Exit;
      end;
  until (iCode = 0);
  Xls := TXlsFile.Create(false);
  try
    xls.IgnoreFormulaText := true; //bme - hard code this for this situation
    xls.VirtualMode := false;
    try
      xls.Open(OpenDialogSprdSheet.FileName);
    finally
    end;
    with dmNrm do
    begin
      cdsNormsFac.First;
    end;
    MatchElementsInFile;
    GetElementOrder;
    with dmNrm do
    begin
      {
      cdsNormsFac.First;
      repeat
        tmpStr := NormsFacColumn.AsString;
        cdsNormsFac.Edit;
        cdsNormsFacColumnNo.AsInteger := ConvertCol2Int(tmpStr);
        cdsNormsFac.Post;
        cdsNormsFac.Next;
      until cdsNormsFac.EOF;
      }
      j := 1;
      for i := FromRow to ToRow do
      begin
        cdsNormChem.Append;
        for j:= 1 to 24 do
        begin
          //SprdSheet.Row := i;
          if (ElementPos[j] > 0) then
          begin
            //SprdSheet.Col := ElementPos[j];
            v := Xls.GetCellValue(i,ElementPos[j]);
            tmpDataStr := v.ToString;
            //if (j = 1) then
            //  cdsNormChemGroupName.AsString := tmpDataStr;
            if (j = 1) then
              cdsNormChemSampleNum.AsString := tmpDataStr;
            if (j > 1) then
            begin
                if (tmpDataStr <> '') then
                begin
                  try
                    cdsNormChem.Fields[j-1].AsVariant := tmpDataStr
                  except
                    cdsNormChem.Fields[j-1].AsString := '';
                  end;
                end else
                begin
                  cdsNormChem.Fields[j-1].AsVariant := 0.0;
                end;
                if (cdsNormChem.Fields[j-1].AsFloat < 0.0) then
                begin
                  cdsNormChem.Fields[j-1].AsFloat := (Abs(cdsNormChem.Fields[j-1].AsFloat))/2.0;
                end;
            end;
          end else
          begin
            cdsNormChem.Fields[j-1].AsVariant := 0.0;
          end;
        end;
        cdsNormChem.Post;
        cdsNormChem.Edit;
        ConvertDataToOxides;
        cdsNormChem.Post;
      end;
      cdsNormChem.First;
    end;
  finally
    Xls.Free;
  end;
end;

procedure TfmSheetImport.FormCreate(Sender: TObject);
begin
  CellData := TObjectList<TSparseCellArray>.Create;
end;

procedure TfmSheetImport.FormDestroy(Sender: TObject);
begin
  CellData.Free;
end;

procedure TfmSheetImport.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  bbImport.Visible := false;
  gbDefineFields.Visible := false;
  //SprdSheet.Visible := false;
  gbDefineRows.Visible := false;
  {
  with dmNrm do
  begin
    NormsFac.Open;
  end;
  }
end;


procedure TfmSheetImport.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;

procedure TfmSheetImport.ConvertDataToOxides;
begin
   dmNrm.cdsNormChemSIO2.AsFloat:=dmNrm.cdsNormChemSIO2.AsFloat*OxFactor[1];
   dmNrm.cdsNormChemTIO2.AsFloat:=dmNrm.cdsNormChemTIO2.AsFloat*OxFactor[2];
   dmNrm.cdsNormChemZRO2.AsFloat:=dmNrm.cdsNormChemZRO2.AsFloat*OxFactor[3];
   dmNrm.cdsNormChemAL2O3.AsFloat:=dmNrm.cdsNormChemAL2O3.AsFloat*OxFactor[4];
   dmNrm.cdsNormChemCR2O3.AsFloat:=dmNrm.cdsNormChemCR2O3.AsFloat*OxFactor[5];
   dmNrm.cdsNormChemFE2O3.AsFloat:=dmNrm.cdsNormChemFE2O3.AsFloat*OxFactor[6];
   dmNrm.cdsNormChemFEO.AsFloat:=dmNrm.cdsNormChemFEO.AsFloat*OxFactor[7];
   dmNrm.cdsNormChemMNO.AsFloat:=dmNrm.cdsNormChemMNO.AsFloat*OxFactor[8];
   dmNrm.cdsNormChemNIO.AsFloat:=dmNrm.cdsNormChemNIO.AsFloat*OxFactor[9];
   dmNrm.cdsNormChemMGO.AsFloat:=dmNrm.cdsNormChemMGO.AsFloat*OxFactor[10];
   dmNrm.cdsNormChemCAO.AsFloat:=dmNrm.cdsNormChemCAO.AsFloat*OxFactor[11];
   dmNrm.cdsNormChemSRO.AsFloat:=dmNrm.cdsNormChemSRO.AsFloat*OxFactor[12];
   dmNrm.cdsNormChemBAO.AsFloat:=dmNrm.cdsNormChemBAO.AsFloat*OxFactor[13];
   dmNrm.cdsNormChemNA2O.AsFloat:=dmNrm.cdsNormChemNA2O.AsFloat*OxFactor[14];
   dmNrm.cdsNormChemK2O.AsFloat:=dmNrm.cdsNormChemK2O.AsFloat*OxFactor[15];
   dmNrm.cdsNormChemP2O5.AsFloat:=dmNrm.cdsNormChemP2O5.AsFloat*OxFactor[16];
   dmNrm.cdsNormChemLOI.AsFloat:=dmNrm.cdsNormChemLOI.AsFloat*OxFactor[17];
   dmNrm.cdsNormChemH2OM.AsFloat:=dmNrm.cdsNormChemH2OM.AsFloat*OxFactor[18];
   dmNrm.cdsNormChemSO3.AsFloat:=dmNrm.cdsNormChemSO3.AsFloat*OxFactor[19];
   dmNrm.cdsNormChemS.AsFloat:=dmNrm.cdsNormChemS.AsFloat*OxFactor[20];
   dmNrm.cdsNormChemCL.AsFloat:=dmNrm.cdsNormChemCL.AsFloat*OxFactor[21];
   dmNrm.cdsNormChemF.AsFloat:=dmNrm.cdsNormChemF.AsFloat*OxFactor[22];
   dmNrm.cdsNormChemCO2.AsFloat:=dmNrm.cdsNormChemCO2.AsFloat*OxFactor[23];
end;

procedure TfmSheetImport.GetElementOrder;
begin
   (*
   OxFactor[1]:=1.0;  {Si}
   OxFactor[2]:=1.0;  {Ti}
   OxFactor[3]:=0.000135;   {Zr}
   OxFactor[4]:=1.0;  {Al}
   OxFactor[5]:=0.0001461;  {Cr}
   OxFactor[6]:=1.0;  {Fe3}
   OxFactor[7]:=1.0;  {Fe2}
   OxFactor[8]:=1.0;  {Mn}
   OxFactor[9]:=0.0001272;  {Ni}
   OxFactor[10]:=1.0; {Mg}
   OxFactor[11]:=1.0; {Ca}
   OxFactor[12]:=0.0001182; {Sr}
   OxFactor[13]:=0.0001116; {Ba}
   OxFactor[14]:=1.0; {Na}
   OxFactor[15]:=1.0; {K }
   OxFactor[16]:=1.0; {P }
   OxFactor[17]:=1.0; {H+}
   OxFactor[18]:=1.0;    {H-}
   OxFactor[19]:=1.0;    {SO3}
   OxFactor[20]:=1.0;    {S }
   OxFactor[21]:=1.0;    {Cl}
   OxFactor[22]:=1.0;    {F }
   OxFactor[23]:=1.0;    {CO2}
   *)
   with dmNrm do
   begin
     cdsNormsFac.First;
     repeat
      ElementPos[cdsNormsFacPos.AsInteger+1] := cdsNormsFacColumnNo.AsInteger;
      if (cdsNormsFacPos.AsInteger > 0) then
      begin
        OxFactor[cdsNormsFacPos.AsInteger] := cdsNormsFacFactor.AsFloat;
      end;
       cdsNormsFac.Next;
     until cdsNormsFac.EOF;
     cdsNormsFac.First;
   end;
end;{proc GetElementOrder}

procedure TfmSheetImport.MatchElementsInFile;
var
  tmpStr : string;
begin
  with dmNrm do
  begin
    cdsNormsFac.First;
    repeat
      tmpStr := UpperCase(cdsNormsFacColumn.AsString);
      ClearNull(tmpStr);
      cdsNormsFac.Edit;
      cdsNormsFacColumn.AsString := tmpStr;
      cdsNormsFac.Post;
      cdsNormsFac.Edit;
      if (cdsNormsFacColumn.AsString >= 'A') then
      begin
        cdsNormsFacColumnNo.AsInteger := ConvertCol2Int(tmpStr);
      end else
      begin
        cdsNormsFacColumnNo.AsInteger := 0;
      end;
      cdsNormsFac.Post;
      cdsNormsFac.Next;
    until cdsNormsFac.EOF;
    cdsNormsFac.First;
  end;
end;



function TfmSheetImport.GetCellValue(const aCol, aRow: integer): string;
begin
if ACol = 0 then
  begin
    if ARow = 0 then exit('');
    exit (IntToStr(aRow));
  end;
  if ARow = 0 then exit(TCellAddress.EncodeColumn(aCol));
  if (TabControl1.TabIndex < 0) or (TabControl1.TabIndex >= CellData.Count) then exit('');
  if CellData[TabControl1.TabIndex] = nil then exit('');
  exit(CellData[TabControl1.TabIndex].GetValue(ARow, aCol));
end;

procedure TfmSheetImport.FillGrid;
var
  sheet: integer;
begin
  sheet := TabControl1.TabIndex;
  if (sheet < 0) or (sheet >= CellData.Count) then
  begin
    SheetData.ColCount := 1;
    SheetData.RowCount := 1;
    SheetData.Invalidate;
    exit;
  end;
  if CellData[sheet] <> nil then
  begin
    SheetData.ColCount := CellData[sheet].ColCount + 1;
    SheetData.RowCount := CellData[sheet].RowCount + 1;
  end
  else
  begin
    SheetData.ColCount := 1;
    SheetData.RowCount := 1;
  end;
  if (SheetData.ColCount > 1) and (SheetData.RowCount > 1) then
  begin
    SheetData.FixedRows := 1;
    SheetData.FixedCols := 1;
  end;
  MyActiveSheetNum := sheet+1;
  SheetData.Invalidate;
end;


end.
