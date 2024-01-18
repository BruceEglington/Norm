unit Nrm_shtmin;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, FileCtrl, Grids,
  DBGrids, DB, AxCtrls;

type
  TfmNrmMinSheet = class(TForm)
    Panel1: TPanel;
    sbClose: TSpeedButton;
    sbSheet: TStatusBar;
    SaveDialogSprdSheet: TSaveDialog;
    gb3: TGroupBox;
    bbSaveSheet: TBitBtn;
    //SprdSheet: TF1Book6;
    ds1: TDataSource;
    procedure sbCloseClick(Sender: TObject);
    procedure bbSaveSheetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    MaxIsData : Integer;
    FileSavePath : string;
    FileSaveName : string;
    FileOpenPathAndName : string;
    IsData : array[1..200] of boolean;
    procedure FillSheet;
  end;

var
  fmNrmMinSheet: TfmNrmMinSheet;

implementation

uses Nrm_dm_acs, normvarb;

{$R *.DFM}

procedure TfmNrmMinSheet.sbCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfmNrmMinSheet.bbSaveSheetClick(Sender: TObject);
//const
  //Excel5Type = 4;
  //Excel97Type = 11;
  //VisualComponentType = 5;
  //FormulaOne6Type = 12;
var
  pFileType : smallint;
  pBuf      : string;
  pTitle    : string;
  tmpStr    : string[3];
begin
  SaveDialogSprdSheet.InitialDir := FileSavePath;
  SaveDialogSprdSheet.FileName := FileSaveName;
  if SaveDialogSprdSheet.Execute then
  begin
    FileSavePath := ExtractFilePath(SaveDialogSprdSheet.FileName);
    MineralTablePath := FileSavePath;
    {
    pFileType := Excel97Type;
    case SaveDialogSprdSheet.FilterIndex of
      1 : pFileType := Excel97Type;
      2 : pFileType := Excel5Type;
    end;
    pBuf := SaveDialogSprdSheet.FileName;
    SprdSheet.Write(pBuf,pFileType);
    }
  end;
end;

procedure TfmNrmMinSheet.FillSheet;
var
  i, j : integer;
begin
  ds1.DataSet.DisableControls;
  try
    i := 1;
    ds1.DataSet.First;
    //SprdSheet.Row := i;
    for j := 0 to ds1.DataSet.FieldCount - 1 do
    begin
      //SprdSheet.Col := j+1;
      //SprdSheet.Text := ds1.DataSet.Fields[j].FieldName;
    end;
    for i := 1 to ds1.DataSet.RecordCount do
    begin
      //SprdSheet.Row := i+1;
      for j := 0 to ds1.DataSet.FieldCount - 1 do
      begin
        //SprdSheet.Col := j+1;
        if ((j+1 <= MaxIsData) and (IsData[j+1] = true)) then
        begin
          //SprdSheet.NumberFormatLocal := '##0.00';
          //SprdSheet.Number := ds1.DataSet.Fields[j].AsVariant;
        end else
        begin
          //SprdSheet.Text := ds1.DataSet.Fields[j].AsString;
        end;
      end;
      ds1.DataSet.Next;
    end;
    //SprdSheet.Row := 1;
    ds1.DataSet.First;
  finally
    ds1.DataSet.EnableControls;
  end;
  //SprdSheet.MaxCol := ds1.DataSet.FieldCount + 2;
  //SprdSheet.MaxRow := ds1.DataSet.RecordCount + 2;
  //SprdSheet.Row := 1;
  //SprdSheet.Col := 1;
  //SprdSheet.ShowActiveCell;
end;

procedure TfmNrmMinSheet.FormCreate(Sender: TObject);
begin
  FileSavePath := '';
  FileOpenPathAndName := '';
  FileSaveName := '';
  FileSavePath := MineralTablePath;
  MaxIsData := 200;
end;

procedure TfmNrmMinSheet.FormShow(Sender: TObject);
begin
  gb3.enabled := true;
  bbSaveSheet.Enabled := true;
  FillSheet;
end;


end.
