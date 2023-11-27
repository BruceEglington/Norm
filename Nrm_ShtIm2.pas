unit Nrm_ShtIm2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls,
  VCL.FlexCel.Core, FlexCel.Render, FlexCel.Preview,
  FlexCel.XlsAdapter,
  Vcl.Tabs, Data.DB, System.ImageList, Vcl.ImgList, Vcl.VirtualImageList;

type
  TfmSheetImport = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    meFromRow: TEdit;
    meToRow: TEdit;
    bbImport: TBitBtn;
    Memo1: TMemo;
    Label4: TLabel;
    pDefineRows: TPanel;
    Splitter1: TSplitter;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    sbFindLastRow: TSpeedButton;
    pDefinitions: TPanel;
    pDefineFields: TPanel;
    TabControl: TTabControl;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    pSpreadSheet: TPanel;
    lFilePath: TLabel;
    VirtualImageList1: TVirtualImageList;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure SheetDataSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure TabsChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
  private
    { Private declarations }
    Xls : TXlsFile;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
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
  Normvarb, allsorts, Nrm_dm_acs;

var
  iRec, iRecCount      : integer;

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

function TfmSheetImport.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport.SheetDataSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  //SelectedCell(aCol, aRow);
  CanSelect := true;
end;

procedure TfmSheetImport.TabsChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin
  //
end;

procedure TfmSheetImport.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string;
  i         : integer;
begin
  TabControl.Tabs.Clear;
  cbSheetname.Items.Clear;
  OpenDialogSprdSheet.InitialDir := DataPath;
  if OpenDialogSprdSheet.Execute then
  begin
    DataPath := ExtractFilePath(OpenDialogSprdSheet.FileName);
    lFilePath.Caption := OpenDialogSprdSheet.FileName;
    //Open the Excel file.
    if Xls = nil then Xls := TXlsFile.Create(false);
    xls.Open(OpenDialogSprdSheet.FileName);
    FillTabs;
    Tabs.TabIndex := Xls.ActiveSheet - 1;
    cbSheetName.ItemIndex := Xls.ActiveSheet - 1;
    FillGrid(true);
    SheetData.Row := 1;
    SheetData.Col := 1;
    bbImport.Visible := true;
    pDefinitions.Visible := true;
    Splitter1.Visible := true;
    TabControl.Visible := true;
    gbDefineFields.Visible := true;
    gbDefineRows.Visible := true;
    pDefineRows.Visible := true;
    try
      sbFindLastRowClick(Sender);
    except
    end;
  end;
end;

procedure TfmSheetImport.bbImportClick(Sender: TObject);
var
  j, k     : integer;
  iCode : integer;
  i : integer;
  FromRow, ToRow : integer;
  tmpStr : string;
  RefnumCol,
  SampleCol,
  DataCol, VariableCol : integer;
  tmpSampleNo, tmpDataValueStr : string;
  WasSuccessful : boolean;
  AreVariablesCorrect : boolean;
  tSampleNoStr : string;
  tmpDataStr : string;
  tmpData : double;
begin
  //tRefNumStr := 'RefNum';
  tSampleNoStr := 'SampleNo';
  //tFracStr := 'Frac';
  //tZoneIDstr := 'ZoneID';
  //tTechAbrStr := 'TechAbr';
  //tMaterialAbrStr := 'MaterialAbr';
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
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
  try
    with dmNrm do
    begin
      cdsNormsFac.First;
    end;
    MatchElementsInFile;
    GetElementOrder;
    with dmNrm do
    begin
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
            DataCol := ElementPos[j];
            tmpDataValueStr := Trim(Xls.GetStringFromCell(i,DataCol));
            tmpDataStr := String(tmpDataValueStr);
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
    //Xls.Free;
  end;
  sbSheet.Panels[1].Text := 'Finished importing all data';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  TabControl.Visible := false;
  Splitter1.Visible := false;
  pDefinitions.Visible := false;
  meFromRow.Text := '2';
  meToRow.Text := '3';
  bbImport.Enabled := true;
  pDefineRows.Visible := false;
  gbDefineFields.Visible := false;
  gbDefineRows.Visible := false;
  bbOpenSheetClick(Sender);
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

procedure TfmSheetImport.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
  SampleNameColstr : string;
  tmpDataValue : double;
  tmpDataStr : string;
begin
  dmNrm.cdsNormsFac.First;
  SampleNameColstr := UpperCase(dmNrm.cdsNormsFacColumn.AsString);
  ImportSheetNumber := cbSheetName.ItemIndex;
  meToRow.Text := '';
  ToRow := 0;
  //dmNrm.cdsImportSpecVariables.First;
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
  until (iCode = 0);
  //ShowMessage(tmpStr);
  try
    i := FromRow;
    j := ConvertCol2Int('A');
    try
      j := ConvertCol2Int(SampleNameColstr);
    except
      j := ConvertCol2Int('A');
    end;
    ToRow := 0;
    repeat
      i := i + 1;
      if (i > ToRow) then ToRow := i-1;
      meToRow.Text := IntToStr(ToRow);
      //ShowMessage(meToRow.Text);
      try
        tmpStr := Trim(Xls.GetStringFromCell(i,j));
        //ShowMessage(tmpStr);
      except
        tmpStr := '0.0';
      end;
    until (tmpStr = '');
  except
    //MessageDlg('Error reading data in column '+IntToStr(Data.Col),mtwarning,[mbOK],0);
  end;
  //ShowMessage('Finished at '+ meToRow.Text);
  meToRow.Text := IntToStr(ToRow);
  RowCount[ImportSheetNumber] := ToRow + 1;
  SheetData.Row := 1;
end;

procedure TfmSheetImport.cbSheetNameChange(Sender: TObject);
begin
  Tabs.TabIndex := cbSheetname.ItemIndex;
  ClearGrid;
  FillGrid(true);
end;

procedure TfmSheetImport.FillTabs;
var
  s: integer;
begin
  Tabs.Tabs.Clear;
  cbSheetname.Items.Clear;
  for s := 1 to Xls.SheetCount do
  begin
    Tabs.Tabs.Add(Xls.GetSheetName(s));
    cbSheetname.Items.Add(Xls.GetSheetName(s));
  end;
end;

procedure TfmSheetImport.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport.FillGrid(const Formatted: boolean);
var
  r, c, cIndex: Integer;
  v: TCellValue;
begin
  if Xls = nil then exit;

  if (Tabs.TabIndex + 1 <= Xls.SheetCount) and (Tabs.TabIndex >= 0) then Xls.ActiveSheet := Tabs.TabIndex + 1 else Xls.ActiveSheet := 1;
  //Clear data in previous grid
  ClearGrid;
  SheetData.RowCount := 1;
  SheetData.ColCount := 1;
  //FmtBox.Text := '';

  SheetData.RowCount := Xls.RowCount + 1; //Include fixed row
  SheetData.ColCount := Xls.ColCount + 1; //Include fixed col. NOTE THAT COLCOUNT IS SLOW. We use it here because we really need it. See the Performance.pdf doc.

  if (SheetData.ColCount > 1) then SheetData.FixedCols := 1; //it is deleted when we set the width to 1.
  if (SheetData.RowCount > 1) then SheetData.FixedRows := 1;

  for r := 1 to Xls.RowCount do
  begin
    //Instead of looping in all the columns, we will just loop in the ones that have data. This is much faster.
    for cIndex := 1 to Xls.ColCountInRow(r) do
    begin
      c := Xls.ColFromIndex(r, cIndex); //The real column.
      if Formatted then
      begin
        SheetData.Cells[c, r] := Xls.GetStringFromCell(r, c);
      end
      else
      begin
        v := Xls.GetCellValue(r, c);
        SheetData.Cells[c, r] := v.ToString;
      end;
    end;
  end;

  //Fill the row headers
  for r := 1 to SheetData.RowCount - 1 do
  begin
    SheetData.Cells[0, r] := IntToStr(r);
    SheetData.RowHeights[r] := Round(Xls.GetRowHeight(r) / TExcelMetrics.RowMultDisplay(Xls));
  end;

  //Fill the column headers
  for c := 1 to SheetData.ColCount - 1 do
  begin
    SheetData.Cells[c, 0] := TCellAddress.EncodeColumn(c);
    SheetData.ColWidths[c] := Round(Xls.GetColWidth(c) / TExcelMetrics.ColMult(Xls));
  end;

  //SelectedCell(1,1);

end;

end.
