unit Nrm_dm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables;

type
  TdmNrm = class(TDataModule)
    NormChem: TTable;
    dsNormChem: TDataSource;
    NormsMin: TTable;
    dsNormsMin: TDataSource;
    NormsFac: TTable;
    dsNormsFac: TDataSource;
    NormsFacPos: TSmallintField;
    NormsFacFactor: TFloatField;
    NormsFacColumn: TStringField;
    NormsFacColumnNo: TSmallintField;
    NormsFacRequired: TStringField;
    NormsFacEntered: TStringField;
    NormsMinLinked: TTable;
    dsNormsMinLinked: TDataSource;
    NormsCat: TTable;
    dsNormsCat: TDataSource;
    NormChemGROUPNAME: TStringField;
    NormChemSAMPLENUM: TStringField;
    NormChemSIO2: TFloatField;
    NormChemTIO2: TFloatField;
    NormChemZRO2: TFloatField;
    NormChemAL2O3: TFloatField;
    NormChemCR2O3: TFloatField;
    NormChemFE2O3: TFloatField;
    NormChemFEO: TFloatField;
    NormChemMNO: TFloatField;
    NormChemNIO: TFloatField;
    NormChemMGO: TFloatField;
    NormChemCAO: TFloatField;
    NormChemSRO: TFloatField;
    NormChemBAO: TFloatField;
    NormChemNA2O: TFloatField;
    NormChemK2O: TFloatField;
    NormChemP2O5: TFloatField;
    NormChemLOI: TFloatField;
    NormChemH2OM: TFloatField;
    NormChemSO3: TFloatField;
    NormChemS: TFloatField;
    NormChemCL: TFloatField;
    NormChemF: TFloatField;
    NormChemCO2: TFloatField;
    NormChemTOTAL: TFloatField;
    NormsCatSAMPLENUM: TStringField;
    NormsCatSi: TFloatField;
    NormsCatTi: TFloatField;
    NormsCatZr: TFloatField;
    NormsCatAl: TFloatField;
    NormsCatCr: TFloatField;
    NormsCatFe3: TFloatField;
    NormsCatFe2: TFloatField;
    NormsCatMn: TFloatField;
    NormsCatNi: TFloatField;
    NormsCatMg: TFloatField;
    NormsCatCa: TFloatField;
    NormsCatSr: TFloatField;
    NormsCatBa: TFloatField;
    NormsCatNa: TFloatField;
    NormsCatK: TFloatField;
    NormsCatP: TFloatField;
    NormsCatS: TFloatField;
    NormsCatSO: TFloatField;
    NormsCatCL: TFloatField;
    NormsCatF: TFloatField;
    NormsCatC: TFloatField;
    NormsCatTotal: TFloatField;
    NormsMinNORMTYPE: TStringField;
    NormsMinGROUPNAME: TStringField;
    NormsMinSAMPLENUM: TStringField;
    NormsMinQz: TFloatField;
    NormsMinCo: TFloatField;
    NormsMinZ: TFloatField;
    NormsMinOr: TFloatField;
    NormsMinPl: TFloatField;
    NormsMinPlAb: TFloatField;
    NormsMinPlAn: TFloatField;
    NormsMinLc: TFloatField;
    NormsMinNe: TFloatField;
    NormsMinKp: TFloatField;
    NormsMinHl: TFloatField;
    NormsMinTh: TFloatField;
    NormsMinAc: TFloatField;
    NormsMinNs: TFloatField;
    NormsMinKs: TFloatField;
    NormsMinWo: TFloatField;
    NormsMinDi: TFloatField;
    NormsMinDiWo: TFloatField;
    NormsMinDiEn: TFloatField;
    NormsMinDiFs: TFloatField;
    NormsMinHy: TFloatField;
    NormsMinHyEn: TFloatField;
    NormsMinHyFs: TFloatField;
    NormsMinOl: TFloatField;
    NormsMinOlFo: TFloatField;
    NormsMinOlFa: TFloatField;
    NormsMinCs: TFloatField;
    NormsMinMt: TFloatField;
    NormsMinCm: TFloatField;
    NormsMinIl: TFloatField;
    NormsMinHm: TFloatField;
    NormsMinSp: TFloatField;
    NormsMinPf: TFloatField;
    NormsMinRu: TFloatField;
    NormsMinAp: TFloatField;
    NormsMinFl: TFloatField;
    NormsMinPy: TFloatField;
    NormsMinCc: TFloatField;
    NormsMinBi: TFloatField;
    NormsMinHo: TFloatField;
    NormsMinHoAct: TFloatField;
    NormsMinHoEd: TFloatField;
    NormsMinHoRi: TFloatField;
    NormsMinSpnl: TFloatField;
    NormsMinSALIC: TFloatField;
    NormsMinFEMIC: TFloatField;
    NormsMinTOTAL: TFloatField;
    NormsMinANPL: TFloatField;
    NormsMinFAOL: TFloatField;
    NormsMinENHY: TFloatField;
    NormsMinWQz: TFloatField;
    NormsMinWCo: TFloatField;
    NormsMinWZ: TFloatField;
    NormsMinWOr: TFloatField;
    NormsMinWPl: TFloatField;
    NormsMinWPlAb: TFloatField;
    NormsMinWPlAn: TFloatField;
    NormsMinWLc: TFloatField;
    NormsMinWNe: TFloatField;
    NormsMinWKp: TFloatField;
    NormsMinWHl: TFloatField;
    NormsMinWTh: TFloatField;
    NormsMinWAc: TFloatField;
    NormsMinWNs: TFloatField;
    NormsMinWKs: TFloatField;
    NormsMinWWo: TFloatField;
    NormsMinWDi: TFloatField;
    NormsMinWDiWo: TFloatField;
    NormsMinWDiEn: TFloatField;
    NormsMinWDiFs: TFloatField;
    NormsMinWHy: TFloatField;
    NormsMinWHyEn: TFloatField;
    NormsMinWHyFs: TFloatField;
    NormsMinWOl: TFloatField;
    NormsMinWOlFo: TFloatField;
    NormsMinWOlFa: TFloatField;
    NormsMinWCs: TFloatField;
    NormsMinWMt: TFloatField;
    NormsMinWCm: TFloatField;
    NormsMinWIl: TFloatField;
    NormsMinWHm: TFloatField;
    NormsMinWSp: TFloatField;
    NormsMinWPf: TFloatField;
    NormsMinWRu: TFloatField;
    NormsMinWAp: TFloatField;
    NormsMinWFl: TFloatField;
    NormsMinWPy: TFloatField;
    NormsMinWCc: TFloatField;
    NormsMinWBi: TFloatField;
    NormsMinWHo: TFloatField;
    NormsMinWHoAct: TFloatField;
    NormsMinWHoEd: TFloatField;
    NormsMinWHoRi: TFloatField;
    NormsMinWSpnl: TFloatField;
    NormsMinWANPL: TFloatField;
    NormsMinWFAOL: TFloatField;
    NormsMinWENHY: TFloatField;
    NormsMinWSALIC: TFloatField;
    NormsMinWFEMIC: TFloatField;
    NormsMinWTOTAL: TFloatField;
    NormsMinQzAbOrTQZ: TFloatField;
    NormsMinQzAbOrTAB: TFloatField;
    NormsMinQzAbOrTOR: TFloatField;
    NormsMinQzAbOrWTQZ: TFloatField;
    NormsMinQzAbOrWTAB: TFloatField;
    NormsMinQzAbOrWTOR: TFloatField;
    NormsMinQzNeKpTQZ: TFloatField;
    NormsMinQzNeKpTNE: TFloatField;
    NormsMinQzNeKpTKP: TFloatField;
    NormsMinQzNeKpWTQZ: TFloatField;
    NormsMinQzNeKpWTNE: TFloatField;
    NormsMinQzNeKpWTKP: TFloatField;
    NormsMinOrAbAnTAN: TFloatField;
    NormsMinOrAbAnTOR: TFloatField;
    NormsMinOrAbAnTAB: TFloatField;
    NormsMinOrAbAnWTAN: TFloatField;
    NormsMinOrAbAnWTAB: TFloatField;
    NormsMinOrAbAnWTOR: TFloatField;
    NormsMinWAFMA: TFloatField;
    NormsMinWAFMF: TFloatField;
    NormsMinWAFMM: TFloatField;
    NormsMinAFMA: TFloatField;
    NormsMinAFMF: TFloatField;
    NormsMinAFMM: TFloatField;
    NormsMinWAgpaitic: TFloatField;
    NormsMinAgpaitic: TFloatField;
    NormsMinWFeMgRat: TFloatField;
    NormsMinFeMgRat: TFloatField;
    NormsMinWAlkRat: TFloatField;
    NormsMinAlkRat: TFloatField;
    NormsMinOxidationRat: TFloatField;
    NormsMinWOxidationRat: TFloatField;
    NormsMinWrightsAlkInd: TFloatField;
    NormsMinTotalAlkalis: TFloatField;
    NormsMinDiffInd: TFloatField;
    NormsMinWDiffInd: TFloatField;
    NormsMinWatsonM: TFloatField;
    NormsMinLinkedNORMTYPE: TStringField;
    NormsMinLinkedGROUPNAME: TStringField;
    NormsMinLinkedSAMPLENUM: TStringField;
    NormsMinLinkedQz: TFloatField;
    NormsMinLinkedCo: TFloatField;
    NormsMinLinkedZ: TFloatField;
    NormsMinLinkedOr: TFloatField;
    NormsMinLinkedPl: TFloatField;
    NormsMinLinkedPlAb: TFloatField;
    NormsMinLinkedPlAn: TFloatField;
    NormsMinLinkedLc: TFloatField;
    NormsMinLinkedNe: TFloatField;
    NormsMinLinkedKp: TFloatField;
    NormsMinLinkedHl: TFloatField;
    NormsMinLinkedTh: TFloatField;
    NormsMinLinkedAc: TFloatField;
    NormsMinLinkedNs: TFloatField;
    NormsMinLinkedKs: TFloatField;
    NormsMinLinkedWo: TFloatField;
    NormsMinLinkedDi: TFloatField;
    NormsMinLinkedDiWo: TFloatField;
    NormsMinLinkedDiEn: TFloatField;
    NormsMinLinkedDiFs: TFloatField;
    NormsMinLinkedHy: TFloatField;
    NormsMinLinkedHyEn: TFloatField;
    NormsMinLinkedHyFs: TFloatField;
    NormsMinLinkedOl: TFloatField;
    NormsMinLinkedOlFo: TFloatField;
    NormsMinLinkedOlFa: TFloatField;
    NormsMinLinkedCs: TFloatField;
    NormsMinLinkedMt: TFloatField;
    NormsMinLinkedCm: TFloatField;
    NormsMinLinkedIl: TFloatField;
    NormsMinLinkedHm: TFloatField;
    NormsMinLinkedSp: TFloatField;
    NormsMinLinkedPf: TFloatField;
    NormsMinLinkedRu: TFloatField;
    NormsMinLinkedAp: TFloatField;
    NormsMinLinkedFl: TFloatField;
    NormsMinLinkedPy: TFloatField;
    NormsMinLinkedCc: TFloatField;
    NormsMinLinkedBi: TFloatField;
    NormsMinLinkedHo: TFloatField;
    NormsMinLinkedHoAct: TFloatField;
    NormsMinLinkedHoEd: TFloatField;
    NormsMinLinkedHoRi: TFloatField;
    NormsMinLinkedSpnl: TFloatField;
    NormsMinLinkedSALIC: TFloatField;
    NormsMinLinkedFEMIC: TFloatField;
    NormsMinLinkedTOTAL: TFloatField;
    NormsMinLinkedANPL: TFloatField;
    NormsMinLinkedFAOL: TFloatField;
    NormsMinLinkedENHY: TFloatField;
    NormsMinLinkedWQz: TFloatField;
    NormsMinLinkedWCo: TFloatField;
    NormsMinLinkedWZ: TFloatField;
    NormsMinLinkedWOr: TFloatField;
    NormsMinLinkedWPl: TFloatField;
    NormsMinLinkedWPlAb: TFloatField;
    NormsMinLinkedWPlAn: TFloatField;
    NormsMinLinkedWLc: TFloatField;
    NormsMinLinkedWNe: TFloatField;
    NormsMinLinkedWKp: TFloatField;
    NormsMinLinkedWHl: TFloatField;
    NormsMinLinkedWTh: TFloatField;
    NormsMinLinkedWAc: TFloatField;
    NormsMinLinkedWNs: TFloatField;
    NormsMinLinkedWKs: TFloatField;
    NormsMinLinkedWWo: TFloatField;
    NormsMinLinkedWDi: TFloatField;
    NormsMinLinkedWDiWo: TFloatField;
    NormsMinLinkedWDiEn: TFloatField;
    NormsMinLinkedWDiFs: TFloatField;
    NormsMinLinkedWHy: TFloatField;
    NormsMinLinkedWHyEn: TFloatField;
    NormsMinLinkedWHyFs: TFloatField;
    NormsMinLinkedWOl: TFloatField;
    NormsMinLinkedWOlFo: TFloatField;
    NormsMinLinkedWOlFa: TFloatField;
    NormsMinLinkedWCs: TFloatField;
    NormsMinLinkedWMt: TFloatField;
    NormsMinLinkedWCm: TFloatField;
    NormsMinLinkedWIl: TFloatField;
    NormsMinLinkedWHm: TFloatField;
    NormsMinLinkedWSp: TFloatField;
    NormsMinLinkedWPf: TFloatField;
    NormsMinLinkedWRu: TFloatField;
    NormsMinLinkedWAp: TFloatField;
    NormsMinLinkedWFl: TFloatField;
    NormsMinLinkedWPy: TFloatField;
    NormsMinLinkedWCc: TFloatField;
    NormsMinLinkedWBi: TFloatField;
    NormsMinLinkedWHo: TFloatField;
    NormsMinLinkedWHoAct: TFloatField;
    NormsMinLinkedWHoEd: TFloatField;
    NormsMinLinkedWHoRi: TFloatField;
    NormsMinLinkedWSpnl: TFloatField;
    NormsMinLinkedWANPL: TFloatField;
    NormsMinLinkedWFAOL: TFloatField;
    NormsMinLinkedWENHY: TFloatField;
    NormsMinLinkedWSALIC: TFloatField;
    NormsMinLinkedWFEMIC: TFloatField;
    NormsMinLinkedWTOTAL: TFloatField;
    NormsMinLinkedQzAbOrTQZ: TFloatField;
    NormsMinLinkedQzAbOrTAB: TFloatField;
    NormsMinLinkedQzAbOrTOR: TFloatField;
    NormsMinLinkedQzAbOrWTQZ: TFloatField;
    NormsMinLinkedQzAbOrWTAB: TFloatField;
    NormsMinLinkedQzAbOrWTOR: TFloatField;
    NormsMinLinkedQzNeKpTQZ: TFloatField;
    NormsMinLinkedQzNeKpTNE: TFloatField;
    NormsMinLinkedQzNeKpTKP: TFloatField;
    NormsMinLinkedQzNeKpWTQZ: TFloatField;
    NormsMinLinkedQzNeKpWTNE: TFloatField;
    NormsMinLinkedQzNeKpWTKP: TFloatField;
    NormsMinLinkedOrAbAnTAN: TFloatField;
    NormsMinLinkedOrAbAnTOR: TFloatField;
    NormsMinLinkedOrAbAnTAB: TFloatField;
    NormsMinLinkedOrAbAnWTAN: TFloatField;
    NormsMinLinkedOrAbAnWTAB: TFloatField;
    NormsMinLinkedOrAbAnWTOR: TFloatField;
    NormsMinLinkedWAFMA: TFloatField;
    NormsMinLinkedWAFMF: TFloatField;
    NormsMinLinkedWAFMM: TFloatField;
    NormsMinLinkedAFMA: TFloatField;
    NormsMinLinkedAFMF: TFloatField;
    NormsMinLinkedAFMM: TFloatField;
    NormsMinLinkedWAgpaitic: TFloatField;
    NormsMinLinkedAgpaitic: TFloatField;
    NormsMinLinkedWFeMgRat: TFloatField;
    NormsMinLinkedFeMgRat: TFloatField;
    NormsMinLinkedWAlkRat: TFloatField;
    NormsMinLinkedAlkRat: TFloatField;
    NormsMinLinkedOxidationRat: TFloatField;
    NormsMinLinkedWOxidationRat: TFloatField;
    NormsMinLinkedWrightsAlkInd: TFloatField;
    NormsMinLinkedTotalAlkalis: TFloatField;
    NormsMinLinkedDiffInd: TFloatField;
    NormsMinLinkedWDiffInd: TFloatField;
    NormsMinLinkedWatsonM: TFloatField;
    NormsMinR1: TFloatField;
    NormsMinR2: TFloatField;
    NormsMinLinkedR1: TFloatField;
    NormsMinLinkedR2: TFloatField;
    NormsMinChemicalIndexOfAlteration: TFloatField;
    NormsMinLinkedChemicalIndexOfAlteration: TFloatField;
    NormsMinRoserKorschD1: TFloatField;
    NormsMinRoserKorschD2: TFloatField;
    NormsMinRoserKorschD3: TFloatField;
    NormsMinRoserKorschD4: TFloatField;
    NormsMinLinkedRoserKorschD1: TFloatField;
    NormsMinLinkedRoserKorschD2: TFloatField;
    NormsMinLinkedRoserKorschD3: TFloatField;
    NormsMinLinkedRoserKorschD4: TFloatField;
    NormsMinPeraluminousIndex: TFloatField;
    NormsMinLinkedPeraluminousIndex: TFloatField;
    NormsMinDebonLefortA: TFloatField;
    NormsMinDebonLefortB: TFloatField;
    NormsMinLinkedDebonLefortA: TFloatField;
    NormsMinLinkedDebonLefortB: TFloatField;
    procedure NormsCatPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
    procedure NormsMinPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
    procedure NormsMinLinkedPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
    procedure NormChemPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dmNrm: TdmNrm;

implementation

{$R *.DFM}
uses
  ErrCodes;

procedure TdmNrm.NormsCatPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
var
  iDBIError: Integer;
begin
  if (E is EDBEngineError) then
  begin
    iDBIError := (E as EDBEngineError).Errors[0].Errorcode;
    case iDBIError of
      eRequiredFieldMissing:
        begin
        end;
      eKeyViol:
        begin
          dmNrm.NormsCat.Delete;
        end;
    end;
  end;
end;

procedure TdmNrm.NormsMinPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
var
  iDBIError: Integer;
begin
  if (E is EDBEngineError) then
  begin
    iDBIError := (E as EDBEngineError).Errors[0].Errorcode;
    case iDBIError of
      eRequiredFieldMissing:
        begin
        end;
      eKeyViol:
        begin
          MessageDlg('Key violation - duplicate combination of Norm type, Group and Sample',
            mtWarning,[mbOK],0);
          dmNrm.NormsMin.Delete;
        end;
    end;
  end;
end;

procedure TdmNrm.NormsMinLinkedPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
var
  iDBIError: Integer;
begin
  if (E is EDBEngineError) then
  begin
    iDBIError := (E as EDBEngineError).Errors[0].Errorcode;
    case iDBIError of
      eRequiredFieldMissing:
        begin
        end;
      eKeyViol:
        begin
          MessageDlg('Key violation - duplicate combination of Norm type, Group and Sample',
            mtWarning,[mbOK],0);
          dmNrm.NormsMinLinked.Delete;
        end;
    end;
  end;
end;

procedure TdmNrm.NormChemPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
var
  iDBIError: Integer;
begin
  if (E is EDBEngineError) then
  begin
    iDBIError := (E as EDBEngineError).Errors[0].Errorcode;
    case iDBIError of
      eRequiredFieldMissing:
        begin
        end;
      eKeyViol:
        begin
          MessageDlg('Key violation - probably duplicate combination of Group and Sample',
            mtWarning,[mbOK],0);
          dmNrm.NormChem.Delete;
        end;
    end;
  end;
end;

end.
