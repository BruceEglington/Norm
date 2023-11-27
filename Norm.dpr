program Norm;

uses
  Forms,
  Nrm_mn in 'Nrm_mn.pas' {fmNormMain},
  Nrm_dm_acs in 'Nrm_dm_acs.pas' {dmNrm: TDataModule},
  Normvarb in 'Normvarb.pas',
  Norm_min in 'Norm_min.pas',
  NORMDESI in 'NORMDESI.PAS',
  Normtern in 'Normtern.pas',
  About in 'About.pas' {AboutBox},
  Vcl.Themes,
  Vcl.Styles,
  Allsorts in '..\Eglington Delphi common code items\Allsorts.pas',
  ErrCodes in '..\Eglington Delphi common code items\ErrCodes.pas',
  Nrm_ShtIm2 in 'Nrm_ShtIm2.pas' {fmSheetImport},
  Nrm_ShtImTemplate2 in 'Nrm_ShtImTemplate2.pas' {fmSheetImportTemplate};

{$R *.RES}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Windows10');
  Application.Title := 'Norm';
  Application.CreateForm(TdmNrm, dmNrm);
  Application.CreateForm(TfmNormMain, fmNormMain);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TfmSheetImport, fmSheetImport);
  Application.CreateForm(TfmSheetImportTemplate, fmSheetImportTemplate);
  Application.Run;
end.
