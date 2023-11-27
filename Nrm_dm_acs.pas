unit Nrm_dm_acs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, Provider, Vcl.BaseImageCollection,
  Vcl.ImageCollection, midaslib;

type
  TdmNrm = class(TDataModule)
    dsNormChem: TDataSource;
    dsNormsMin: TDataSource;
    dsNormsFac: TDataSource;
    dsNormsMinLinked: TDataSource;
    dsNormsCat: TDataSource;
    cdsNormsFac: TClientDataSet;
    cdsNormsFacPOS: TFloatField;
    cdsNormsFacREQUIRED: TWideStringField;
    cdsNormsFacENTERED: TWideStringField;
    cdsNormsFacFACTOR: TFloatField;
    cdsNormsFacCOLUMN: TWideStringField;
    cdsNormsFacCOLUMNNO: TFloatField;
    cdsNormsCat: TClientDataSet;
    cdsNormsCatSAMPLENUM: TWideStringField;
    cdsNormsCatSI: TFloatField;
    cdsNormsCatTI: TFloatField;
    cdsNormsCatZR: TFloatField;
    cdsNormsCatAL: TFloatField;
    cdsNormsCatCR: TFloatField;
    cdsNormsCatFE3: TFloatField;
    cdsNormsCatFE2: TFloatField;
    cdsNormsCatMN: TFloatField;
    cdsNormsCatNI: TFloatField;
    cdsNormsCatMG: TFloatField;
    cdsNormsCatCA: TFloatField;
    cdsNormsCatSR: TFloatField;
    cdsNormsCatBA: TFloatField;
    cdsNormsCatNA: TFloatField;
    cdsNormsCatK: TFloatField;
    cdsNormsCatP: TFloatField;
    cdsNormsCatS: TFloatField;
    cdsNormsCatSO: TFloatField;
    cdsNormsCatCL: TFloatField;
    cdsNormsCatF: TFloatField;
    cdsNormsCatC: TFloatField;
    cdsNormsCatTOTAL: TFloatField;
    cdsNormsMin: TClientDataSet;
    cdsNormsMinNORMTYPE: TWideStringField;
    cdsNormsMinSAMPLENUM: TWideStringField;
    cdsNormsMinQZ: TFloatField;
    cdsNormsMinCO: TFloatField;
    cdsNormsMinZ: TFloatField;
    cdsNormsMinOR: TFloatField;
    cdsNormsMinPL: TFloatField;
    cdsNormsMinPLAB: TFloatField;
    cdsNormsMinPLAN: TFloatField;
    cdsNormsMinLC: TFloatField;
    cdsNormsMinNE: TFloatField;
    cdsNormsMinKP: TFloatField;
    cdsNormsMinHL: TFloatField;
    cdsNormsMinTH: TFloatField;
    cdsNormsMinAC: TFloatField;
    cdsNormsMinNS: TFloatField;
    cdsNormsMinKS: TFloatField;
    cdsNormsMinWO: TFloatField;
    cdsNormsMinDI: TFloatField;
    cdsNormsMinDIWO: TFloatField;
    cdsNormsMinDIEN: TFloatField;
    cdsNormsMinDIFS: TFloatField;
    cdsNormsMinHY: TFloatField;
    cdsNormsMinHYEN: TFloatField;
    cdsNormsMinHYFS: TFloatField;
    cdsNormsMinOL: TFloatField;
    cdsNormsMinOLFO: TFloatField;
    cdsNormsMinOLFA: TFloatField;
    cdsNormsMinCS: TFloatField;
    cdsNormsMinMT: TFloatField;
    cdsNormsMinCM: TFloatField;
    cdsNormsMinIL: TFloatField;
    cdsNormsMinHM: TFloatField;
    cdsNormsMinSP: TFloatField;
    cdsNormsMinPF: TFloatField;
    cdsNormsMinRU: TFloatField;
    cdsNormsMinAP: TFloatField;
    cdsNormsMinFL: TFloatField;
    cdsNormsMinPY: TFloatField;
    cdsNormsMinCC: TFloatField;
    cdsNormsMinBI: TFloatField;
    cdsNormsMinHO: TFloatField;
    cdsNormsMinHOACT: TFloatField;
    cdsNormsMinHOED: TFloatField;
    cdsNormsMinHORI: TFloatField;
    cdsNormsMinSPNL: TFloatField;
    cdsNormsMinSALIC: TFloatField;
    cdsNormsMinFEMIC: TFloatField;
    cdsNormsMinTOTAL: TFloatField;
    cdsNormsMinANPL: TFloatField;
    cdsNormsMinFAOL: TFloatField;
    cdsNormsMinENHY: TFloatField;
    cdsNormsMinWQZ: TFloatField;
    cdsNormsMinWCO: TFloatField;
    cdsNormsMinWZ: TFloatField;
    cdsNormsMinWOR: TFloatField;
    cdsNormsMinWPL: TFloatField;
    cdsNormsMinWPLAB: TFloatField;
    cdsNormsMinWPLAN: TFloatField;
    cdsNormsMinWLC: TFloatField;
    cdsNormsMinWNE: TFloatField;
    cdsNormsMinWKP: TFloatField;
    cdsNormsMinWHL: TFloatField;
    cdsNormsMinWTH: TFloatField;
    cdsNormsMinWAC: TFloatField;
    cdsNormsMinWNS: TFloatField;
    cdsNormsMinWKS: TFloatField;
    cdsNormsMinWWO: TFloatField;
    cdsNormsMinWDI: TFloatField;
    cdsNormsMinWDIWO: TFloatField;
    cdsNormsMinWDIEN: TFloatField;
    cdsNormsMinWDIFS: TFloatField;
    cdsNormsMinWHY: TFloatField;
    cdsNormsMinWHYEN: TFloatField;
    cdsNormsMinWHYFS: TFloatField;
    cdsNormsMinWOL: TFloatField;
    cdsNormsMinWOLFO: TFloatField;
    cdsNormsMinWOLFA: TFloatField;
    cdsNormsMinWCS: TFloatField;
    cdsNormsMinWMT: TFloatField;
    cdsNormsMinWCM: TFloatField;
    cdsNormsMinWIL: TFloatField;
    cdsNormsMinWHM: TFloatField;
    cdsNormsMinWSP: TFloatField;
    cdsNormsMinWPF: TFloatField;
    cdsNormsMinWRU: TFloatField;
    cdsNormsMinWAP: TFloatField;
    cdsNormsMinWFL: TFloatField;
    cdsNormsMinWPY: TFloatField;
    cdsNormsMinWCC: TFloatField;
    cdsNormsMinWBI: TFloatField;
    cdsNormsMinWHO: TFloatField;
    cdsNormsMinWHOACT: TFloatField;
    cdsNormsMinWHOED: TFloatField;
    cdsNormsMinWHORI: TFloatField;
    cdsNormsMinWSPNL: TFloatField;
    cdsNormsMinWANPL: TFloatField;
    cdsNormsMinWFAOL: TFloatField;
    cdsNormsMinWENHY: TFloatField;
    cdsNormsMinWSALIC: TFloatField;
    cdsNormsMinWFEMIC: TFloatField;
    cdsNormsMinWTOTAL: TFloatField;
    cdsNormsMinQZABORTQZ: TFloatField;
    cdsNormsMinQZABORTAB: TFloatField;
    cdsNormsMinQZABORTOR: TFloatField;
    cdsNormsMinQZABORWTQZ: TFloatField;
    cdsNormsMinQZABORWTAB: TFloatField;
    cdsNormsMinQZABORWTOR: TFloatField;
    cdsNormsMinQZNEKPTQZ: TFloatField;
    cdsNormsMinQZNEKPTNE: TFloatField;
    cdsNormsMinQZNEKPTKP: TFloatField;
    cdsNormsMinQZNEKPWTQZ: TFloatField;
    cdsNormsMinQZNEKPWTNE: TFloatField;
    cdsNormsMinQZNEKPWTKP: TFloatField;
    cdsNormsMinORABANTAN: TFloatField;
    cdsNormsMinORABANTOR: TFloatField;
    cdsNormsMinORABANTAB: TFloatField;
    cdsNormsMinORABANWTAN: TFloatField;
    cdsNormsMinORABANWTAB: TFloatField;
    cdsNormsMinORABANWTOR: TFloatField;
    cdsNormsMinWAFMA: TFloatField;
    cdsNormsMinWAFMF: TFloatField;
    cdsNormsMinWAFMM: TFloatField;
    cdsNormsMinAFMA: TFloatField;
    cdsNormsMinAFMF: TFloatField;
    cdsNormsMinAFMM: TFloatField;
    cdsNormsMinWAGPAITIC: TFloatField;
    cdsNormsMinAGPAITIC: TFloatField;
    cdsNormsMinWFEMGRAT: TFloatField;
    cdsNormsMinFEMGRAT: TFloatField;
    cdsNormsMinWALKRAT: TFloatField;
    cdsNormsMinALKRAT: TFloatField;
    cdsNormsMinOXIDRAT: TFloatField;
    cdsNormsMinWOXIDRAT: TFloatField;
    cdsNormsMinWRIGHTSALK: TFloatField;
    cdsNormsMinTOTALALK: TFloatField;
    cdsNormsMinDIFFIND: TFloatField;
    cdsNormsMinWDIFFIND: TFloatField;
    cdsNormsMinWATSONM: TFloatField;
    cdsNormsMinR1: TFloatField;
    cdsNormsMinR2: TFloatField;
    cdsNormsMinCHEMINDALT: TFloatField;
    cdsNormsMinROSKOR_D1: TFloatField;
    cdsNormsMinROSKOR_D2: TFloatField;
    cdsNormsMinROSKOR_D3: TFloatField;
    cdsNormsMinROSKOR_D4: TFloatField;
    cdsNormsMinPERALUMIND: TFloatField;
    cdsNormsMinDEBLEFOR_A: TFloatField;
    cdsNormsMinDEBLEFOR_B: TFloatField;
    cdsNormChem: TClientDataSet;
    cdsNormChemSAMPLENUM: TWideStringField;
    cdsNormChemSIO2: TFloatField;
    cdsNormChemTIO2: TFloatField;
    cdsNormChemZRO2: TFloatField;
    cdsNormChemAL2O3: TFloatField;
    cdsNormChemCR2O3: TFloatField;
    cdsNormChemFE2O3: TFloatField;
    cdsNormChemFEO: TFloatField;
    cdsNormChemMNO: TFloatField;
    cdsNormChemNIO: TFloatField;
    cdsNormChemMGO: TFloatField;
    cdsNormChemCAO: TFloatField;
    cdsNormChemSRO: TFloatField;
    cdsNormChemBAO: TFloatField;
    cdsNormChemNA2O: TFloatField;
    cdsNormChemK2O: TFloatField;
    cdsNormChemP2O5: TFloatField;
    cdsNormChemLOI: TFloatField;
    cdsNormChemH2OM: TFloatField;
    cdsNormChemSO3: TFloatField;
    cdsNormChemS: TFloatField;
    cdsNormChemCL: TFloatField;
    cdsNormChemF: TFloatField;
    cdsNormChemCO2: TFloatField;
    cdsNormChemTOTAL: TFloatField;
    cdsNormsMinLinked: TClientDataSet;
    cdsNormsMinLinkedNORMTYPE: TWideStringField;
    cdsNormsMinLinkedSAMPLENUM: TWideStringField;
    cdsNormsMinLinkedQZ: TFloatField;
    cdsNormsMinLinkedCO: TFloatField;
    cdsNormsMinLinkedZ: TFloatField;
    cdsNormsMinLinkedOR: TFloatField;
    cdsNormsMinLinkedPL: TFloatField;
    cdsNormsMinLinkedPLAB: TFloatField;
    cdsNormsMinLinkedPLAN: TFloatField;
    cdsNormsMinLinkedLC: TFloatField;
    cdsNormsMinLinkedNE: TFloatField;
    cdsNormsMinLinkedKP: TFloatField;
    cdsNormsMinLinkedHL: TFloatField;
    cdsNormsMinLinkedTH: TFloatField;
    cdsNormsMinLinkedAC: TFloatField;
    cdsNormsMinLinkedNS: TFloatField;
    cdsNormsMinLinkedKS: TFloatField;
    cdsNormsMinLinkedWO: TFloatField;
    cdsNormsMinLinkedDI: TFloatField;
    cdsNormsMinLinkedDIWO: TFloatField;
    cdsNormsMinLinkedDIEN: TFloatField;
    cdsNormsMinLinkedDIFS: TFloatField;
    cdsNormsMinLinkedHY: TFloatField;
    cdsNormsMinLinkedHYEN: TFloatField;
    cdsNormsMinLinkedHYFS: TFloatField;
    cdsNormsMinLinkedOL: TFloatField;
    cdsNormsMinLinkedOLFO: TFloatField;
    cdsNormsMinLinkedOLFA: TFloatField;
    cdsNormsMinLinkedCS: TFloatField;
    cdsNormsMinLinkedMT: TFloatField;
    cdsNormsMinLinkedCM: TFloatField;
    cdsNormsMinLinkedIL: TFloatField;
    cdsNormsMinLinkedHM: TFloatField;
    cdsNormsMinLinkedSP: TFloatField;
    cdsNormsMinLinkedPF: TFloatField;
    cdsNormsMinLinkedRU: TFloatField;
    cdsNormsMinLinkedAP: TFloatField;
    cdsNormsMinLinkedFL: TFloatField;
    cdsNormsMinLinkedPY: TFloatField;
    cdsNormsMinLinkedCC: TFloatField;
    cdsNormsMinLinkedBI: TFloatField;
    cdsNormsMinLinkedHO: TFloatField;
    cdsNormsMinLinkedHOACT: TFloatField;
    cdsNormsMinLinkedHOED: TFloatField;
    cdsNormsMinLinkedHORI: TFloatField;
    cdsNormsMinLinkedSPNL: TFloatField;
    cdsNormsMinLinkedSALIC: TFloatField;
    cdsNormsMinLinkedFEMIC: TFloatField;
    cdsNormsMinLinkedTOTAL: TFloatField;
    cdsNormsMinLinkedANPL: TFloatField;
    cdsNormsMinLinkedFAOL: TFloatField;
    cdsNormsMinLinkedENHY: TFloatField;
    cdsNormsMinLinkedWQZ: TFloatField;
    cdsNormsMinLinkedWCO: TFloatField;
    cdsNormsMinLinkedWZ: TFloatField;
    cdsNormsMinLinkedWOR: TFloatField;
    cdsNormsMinLinkedWPL: TFloatField;
    cdsNormsMinLinkedWPLAB: TFloatField;
    cdsNormsMinLinkedWPLAN: TFloatField;
    cdsNormsMinLinkedWLC: TFloatField;
    cdsNormsMinLinkedWNE: TFloatField;
    cdsNormsMinLinkedWKP: TFloatField;
    cdsNormsMinLinkedWHL: TFloatField;
    cdsNormsMinLinkedWTH: TFloatField;
    cdsNormsMinLinkedWAC: TFloatField;
    cdsNormsMinLinkedWNS: TFloatField;
    cdsNormsMinLinkedWKS: TFloatField;
    cdsNormsMinLinkedWWO: TFloatField;
    cdsNormsMinLinkedWDI: TFloatField;
    cdsNormsMinLinkedWDIWO: TFloatField;
    cdsNormsMinLinkedWDIEN: TFloatField;
    cdsNormsMinLinkedWDIFS: TFloatField;
    cdsNormsMinLinkedWHY: TFloatField;
    cdsNormsMinLinkedWHYEN: TFloatField;
    cdsNormsMinLinkedWHYFS: TFloatField;
    cdsNormsMinLinkedWOL: TFloatField;
    cdsNormsMinLinkedWOLFO: TFloatField;
    cdsNormsMinLinkedWOLFA: TFloatField;
    cdsNormsMinLinkedWCS: TFloatField;
    cdsNormsMinLinkedWMT: TFloatField;
    cdsNormsMinLinkedWCM: TFloatField;
    cdsNormsMinLinkedWIL: TFloatField;
    cdsNormsMinLinkedWHM: TFloatField;
    cdsNormsMinLinkedWSP: TFloatField;
    cdsNormsMinLinkedWPF: TFloatField;
    cdsNormsMinLinkedWRU: TFloatField;
    cdsNormsMinLinkedWAP: TFloatField;
    cdsNormsMinLinkedWFL: TFloatField;
    cdsNormsMinLinkedWPY: TFloatField;
    cdsNormsMinLinkedWCC: TFloatField;
    cdsNormsMinLinkedWBI: TFloatField;
    cdsNormsMinLinkedWHO: TFloatField;
    cdsNormsMinLinkedWHOACT: TFloatField;
    cdsNormsMinLinkedWHOED: TFloatField;
    cdsNormsMinLinkedWHORI: TFloatField;
    cdsNormsMinLinkedWSPNL: TFloatField;
    cdsNormsMinLinkedWANPL: TFloatField;
    cdsNormsMinLinkedWFAOL: TFloatField;
    cdsNormsMinLinkedWENHY: TFloatField;
    cdsNormsMinLinkedWSALIC: TFloatField;
    cdsNormsMinLinkedWFEMIC: TFloatField;
    cdsNormsMinLinkedWTOTAL: TFloatField;
    cdsNormsMinLinkedQZABORTQZ: TFloatField;
    cdsNormsMinLinkedQZABORTAB: TFloatField;
    cdsNormsMinLinkedQZABORTOR: TFloatField;
    cdsNormsMinLinkedQZABORWTQZ: TFloatField;
    cdsNormsMinLinkedQZABORWTAB: TFloatField;
    cdsNormsMinLinkedQZABORWTOR: TFloatField;
    cdsNormsMinLinkedQZNEKPTQZ: TFloatField;
    cdsNormsMinLinkedQZNEKPTNE: TFloatField;
    cdsNormsMinLinkedQZNEKPTKP: TFloatField;
    cdsNormsMinLinkedQZNEKPWTQZ: TFloatField;
    cdsNormsMinLinkedQZNEKPWTNE: TFloatField;
    cdsNormsMinLinkedQZNEKPWTKP: TFloatField;
    cdsNormsMinLinkedORABANTAN: TFloatField;
    cdsNormsMinLinkedORABANTOR: TFloatField;
    cdsNormsMinLinkedORABANTAB: TFloatField;
    cdsNormsMinLinkedORABANWTAN: TFloatField;
    cdsNormsMinLinkedORABANWTAB: TFloatField;
    cdsNormsMinLinkedORABANWTOR: TFloatField;
    cdsNormsMinLinkedWAFMA: TFloatField;
    cdsNormsMinLinkedWAFMF: TFloatField;
    cdsNormsMinLinkedWAFMM: TFloatField;
    cdsNormsMinLinkedAFMA: TFloatField;
    cdsNormsMinLinkedAFMF: TFloatField;
    cdsNormsMinLinkedAFMM: TFloatField;
    cdsNormsMinLinkedWAGPAITIC: TFloatField;
    cdsNormsMinLinkedAGPAITIC: TFloatField;
    cdsNormsMinLinkedWFEMGRAT: TFloatField;
    cdsNormsMinLinkedFEMGRAT: TFloatField;
    cdsNormsMinLinkedWALKRAT: TFloatField;
    cdsNormsMinLinkedALKRAT: TFloatField;
    cdsNormsMinLinkedOXIDRAT: TFloatField;
    cdsNormsMinLinkedWOXIDRAT: TFloatField;
    cdsNormsMinLinkedWRIGHTSALK: TFloatField;
    cdsNormsMinLinkedTOTALALK: TFloatField;
    cdsNormsMinLinkedDIFFIND: TFloatField;
    cdsNormsMinLinkedWDIFFIND: TFloatField;
    cdsNormsMinLinkedWATSONM: TFloatField;
    cdsNormsMinLinkedR1: TFloatField;
    cdsNormsMinLinkedR2: TFloatField;
    cdsNormsMinLinkedCHEMINDALT: TFloatField;
    cdsNormsMinLinkedROSKOR_D1: TFloatField;
    cdsNormsMinLinkedROSKOR_D2: TFloatField;
    cdsNormsMinLinkedROSKOR_D3: TFloatField;
    cdsNormsMinLinkedROSKOR_D4: TFloatField;
    cdsNormsMinLinkedPERALUMIND: TFloatField;
    cdsNormsMinLinkedDEBLEFOR_A: TFloatField;
    cdsNormsMinLinkedDEBLEFOR_B: TFloatField;
    ImageCollection1: TImageCollection;
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
    ChosenStyle : string;
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
  {
  if (E is EDBEngineError) then
  begin
    iDBIError := (E as EDBEngineError).Errors[0].Errorcode;
    case iDBIError of
      eRequiredFieldMissing:
        begin
        end;
      eKeyViol:
        begin
          dmNrm.cdsNormsCat.Delete;
        end;
    end;
  end;
  }
end;

procedure TdmNrm.NormsMinPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
var
  iDBIError: Integer;
begin
  {
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
          dmNrm.cdsNormsMin.Delete;
        end;
    end;
  end;
  }
end;

procedure TdmNrm.NormsMinLinkedPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
var
  iDBIError: Integer;
begin
  {
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
          dmNrm.cdsNormsMinLinked.Delete;
        end;
    end;
  end;
  }
end;

procedure TdmNrm.NormChemPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
var
  iDBIError: Integer;
begin
  {
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
          dmNrm.cdsNormChem.Delete;
        end;
    end;
  end;
  }
end;

end.
