unit Nrm_mn;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, ToolWin, ComCtrls,
  System.IOUtils, System.UITypes, VCL.Themes,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, FlexCel.Report,
  Printers, Menus, Mask, DBCtrls, Db, IniFiles, DBClient, System.ImageList,
  Vcl.ImgList, Vcl.VirtualImageList;

type
  TfmNormMain = class(TForm)
    ToolBar1: TToolBar;
    sbMain: TStatusBar;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    bbExit: TBitBtn;
    N1: TMenuItem;
    Import1: TMenuItem;
    Export1: TMenuItem;
    dsNormsMin: TDataSource;
    Help: TMenuItem;
    About1: TMenuItem;
    Printersetup1: TMenuItem;
    PrinterSetupDialog1: TPrinterSetupDialog;
    N2: TMenuItem;
    PrintDialog1: TPrintDialog;
    Options1: TMenuItem;
    Mineraltablepath1: TMenuItem;
    pc1: TPageControl;
    tsControl: TTabSheet;
    tsChemistry: TTabSheet;
    Panel7: TPanel;
    Label10: TLabel;
    Label11: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label79: TLabel;
    Label80: TLabel;
    Label81: TLabel;
    Label82: TLabel;
    Label83: TLabel;
    Label84: TLabel;
    Label85: TLabel;
    Label86: TLabel;
    Label87: TLabel;
    Label88: TLabel;
    Label89: TLabel;
    Label90: TLabel;
    Label91: TLabel;
    Label92: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    Label95: TLabel;
    Label96: TLabel;
    Label97: TLabel;
    Label107: TLabel;
    Label119: TLabel;
    Label120: TLabel;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Label122: TLabel;
    Label124: TLabel;
    Label98: TLabel;
    Label99: TLabel;
    Label100: TLabel;
    Label101: TLabel;
    Label102: TLabel;
    Label103: TLabel;
    Label104: TLabel;
    Label105: TLabel;
    Label106: TLabel;
    Label108: TLabel;
    Label109: TLabel;
    Label110: TLabel;
    Label116: TLabel;
    Label117: TLabel;
    Label118: TLabel;
    Label121: TLabel;
    Label125: TLabel;
    Label126: TLabel;
    Label127: TLabel;
    Label128: TLabel;
    Label135: TLabel;
    Label136: TLabel;
    Label137: TLabel;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    DBEdit27: TDBEdit;
    DBEdit28: TDBEdit;
    DBEdit70: TDBEdit;
    DBEdit73: TDBEdit;
    DBEdit74: TDBEdit;
    DBEdit75: TDBEdit;
    DBEdit76: TDBEdit;
    DBEdit77: TDBEdit;
    DBEdit78: TDBEdit;
    DBEdit79: TDBEdit;
    DBEdit80: TDBEdit;
    DBEdit81: TDBEdit;
    DBEdit82: TDBEdit;
    DBEdit83: TDBEdit;
    DBEdit84: TDBEdit;
    DBEdit85: TDBEdit;
    DBEdit86: TDBEdit;
    DBEdit87: TDBEdit;
    DBEdit88: TDBEdit;
    DBNavigator1: TDBNavigator;
    DBEdit98: TDBEdit;
    DBEdit100: TDBEdit;
    DBEdit101: TDBEdit;
    DBEdit112: TDBEdit;
    DBEdit89: TDBEdit;
    DBEdit90: TDBEdit;
    DBEdit91: TDBEdit;
    DBEdit92: TDBEdit;
    DBEdit93: TDBEdit;
    DBEdit94: TDBEdit;
    DBEdit95: TDBEdit;
    DBEdit96: TDBEdit;
    DBEdit97: TDBEdit;
    DBEdit102: TDBEdit;
    DBEdit103: TDBEdit;
    DBEdit104: TDBEdit;
    DBEdit105: TDBEdit;
    DBEdit106: TDBEdit;
    DBEdit107: TDBEdit;
    DBEdit108: TDBEdit;
    DBEdit109: TDBEdit;
    DBEdit110: TDBEdit;
    DBEdit126: TDBEdit;
    DBEdit127: TDBEdit;
    DBEdit128: TDBEdit;
    tsMineralogy: TTabSheet;
    Panel4: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Label54: TLabel;
    DBText1: TDBText;
    Label111: TLabel;
    Label112: TLabel;
    Label113: TLabel;
    Label130: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit12: TDBEdit;
    DBEdit13: TDBEdit;
    DBEdit14: TDBEdit;
    DBEdit15: TDBEdit;
    DBEdit16: TDBEdit;
    DBEdit17: TDBEdit;
    DBEdit18: TDBEdit;
    DBEdit19: TDBEdit;
    DBEdit20: TDBEdit;
    DBEdit21: TDBEdit;
    DBEdit22: TDBEdit;
    DBEdit23: TDBEdit;
    DBEdit24: TDBEdit;
    DBEdit25: TDBEdit;
    DBEdit26: TDBEdit;
    DBEdit29: TDBEdit;
    DBEdit30: TDBEdit;
    DBEdit31: TDBEdit;
    DBEdit32: TDBEdit;
    DBEdit33: TDBEdit;
    DBEdit34: TDBEdit;
    DBEdit35: TDBEdit;
    DBEdit36: TDBEdit;
    dbNavMin: TDBNavigator;
    DBEdit37: TDBEdit;
    DBEdit39: TDBEdit;
    DBEdit40: TDBEdit;
    DBEdit41: TDBEdit;
    DBEdit42: TDBEdit;
    DBEdit43: TDBEdit;
    DBEdit44: TDBEdit;
    DBEdit45: TDBEdit;
    DBEdit46: TDBEdit;
    DBEdit47: TDBEdit;
    DBEdit48: TDBEdit;
    DBEdit49: TDBEdit;
    DBEdit50: TDBEdit;
    DBEdit119: TDBEdit;
    DBEdit120: TDBEdit;
    DBEdit121: TDBEdit;
    DBEdit114: TDBEdit;
    dbNavChemistry: TDBNavigator;
    tsProjections: TTabSheet;
    Panel5: TPanel;
    Label67: TLabel;
    DBText2: TDBText;
    GroupBox1: TGroupBox;
    Label53: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label141: TLabel;
    DBEdit51: TDBEdit;
    DBEdit52: TDBEdit;
    DBEdit53: TDBEdit;
    DBEdit133: TDBEdit;
    DBEdit134: TDBEdit;
    DBEdit135: TDBEdit;
    GroupBox2: TGroupBox;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label142: TLabel;
    DBEdit54: TDBEdit;
    DBEdit55: TDBEdit;
    DBEdit56: TDBEdit;
    DBEdit136: TDBEdit;
    DBEdit137: TDBEdit;
    DBEdit138: TDBEdit;
    GroupBox3: TGroupBox;
    Label63: TLabel;
    Label64: TLabel;
    Label65: TLabel;
    Label66: TLabel;
    Label143: TLabel;
    DBEdit57: TDBEdit;
    DBEdit58: TDBEdit;
    DBEdit59: TDBEdit;
    DBEdit139: TDBEdit;
    DBEdit140: TDBEdit;
    DBEdit141: TDBEdit;
    DBEdit60: TDBEdit;
    DBNavigator2: TDBNavigator;
    GroupBox4: TGroupBox;
    Label75: TLabel;
    Label76: TLabel;
    Label77: TLabel;
    Label78: TLabel;
    Label144: TLabel;
    DBEdit69: TDBEdit;
    DBEdit71: TDBEdit;
    DBEdit72: TDBEdit;
    DBEdit142: TDBEdit;
    DBEdit143: TDBEdit;
    DBEdit144: TDBEdit;
    Panel8: TPanel;
    Label129: TLabel;
    Label131: TLabel;
    Label132: TLabel;
    Label133: TLabel;
    Label134: TLabel;
    DBEdit111: TDBEdit;
    DBEdit113: TDBEdit;
    DBEdit115: TDBEdit;
    DBEdit116: TDBEdit;
    DBEdit117: TDBEdit;
    DBEdit118: TDBEdit;
    DBNavigator4: TDBNavigator;
    tsOther: TTabSheet;
    Panel9: TPanel;
    Label145: TLabel;
    DBText3: TDBText;
    DBEdit163: TDBEdit;
    DBNavigator7: TDBNavigator;
    Panel10: TPanel;
    Label162: TLabel;
    Label163: TLabel;
    Label164: TLabel;
    Label165: TLabel;
    Label166: TLabel;
    Label167: TLabel;
    Label168: TLabel;
    Label169: TLabel;
    Label170: TLabel;
    Label171: TLabel;
    Label172: TLabel;
    Label173: TLabel;
    Label147: TLabel;
    Label148: TLabel;
    Label72: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    DBEdit165: TDBEdit;
    DBEdit166: TDBEdit;
    DBEdit167: TDBEdit;
    DBEdit168: TDBEdit;
    DBEdit169: TDBEdit;
    DBEdit170: TDBEdit;
    DBEdit171: TDBEdit;
    DBEdit172: TDBEdit;
    DBEdit173: TDBEdit;
    DBEdit174: TDBEdit;
    DBEdit175: TDBEdit;
    DBEdit176: TDBEdit;
    DBEdit177: TDBEdit;
    DBEdit178: TDBEdit;
    DBEdit179: TDBEdit;
    DBEdit147: TDBEdit;
    DBEdit66: TDBEdit;
    DBEdit67: TDBEdit;
    DBEdit68: TDBEdit;
    DBNavigator8: TDBNavigator;
    Panel6: TPanel;
    Label9: TLabel;
    Label69: TLabel;
    Label70: TLabel;
    Label71: TLabel;
    DBEdit62: TDBEdit;
    DBEdit63: TDBEdit;
    DBEdit64: TDBEdit;
    DBEdit65: TDBEdit;
    PanelMain: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Splitter1: TSplitter;
    Panel1: TPanel;
    rbCIPW: TRadioButton;
    rbMesoPx: TRadioButton;
    rbMesoHb: TRadioButton;
    bbCalculate: TBitBtn;
    cbRecalculate100: TCheckBox;
    cbPrint: TCheckBox;
    Panel13: TPanel;
    Panel2: TPanel;
    dbgChemistry: TDBGrid;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Splitter2: TSplitter;
    Panel18: TPanel;
    Panel17: TPanel;
    Panel3: TPanel;
    dbgMinerals: TDBGrid;
    Panel19: TPanel;
    bbEmptyNormsMin: TBitBtn;
    dbnMineralogy: TDBNavigator;
    Panel20: TPanel;
    Panel21: TPanel;
    Panel22: TPanel;
    Panel23: TPanel;
    dbnChemistry: TDBNavigator;
    bbEmptyNormChem: TBitBtn;
    Panel24: TPanel;
    SaveDialogSprdSheet: TSaveDialog;
    N3: TMenuItem;
    ImportTemplate1: TMenuItem;
    tsTemplate: TTabSheet;
    Panel25: TPanel;
    DBGrid1: TDBGrid;
    VirtualImageList1: TVirtualImageList;
    Styles1: TMenuItem;
    Test1: TMenuItem;
    procedure bbCalculateClick(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure bbEmptyNormsMinClick(Sender: TObject);
    procedure Export1Click(Sender: TObject);
    procedure Import1Click(Sender: TObject);
    procedure bbEmptyNormChemClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Printersetup1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Mineraltablepath1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StyleClick(Sender: TObject);
    procedure Test1Click(Sender: TObject);
    procedure ImportTemplate1Click(Sender: TObject);
  private
    { Private declarations }
    procedure CalcCIPW;
    procedure CalcMeso;
    procedure Record_To_Memory;
    procedure Memory_To_Record;
    procedure ExportNormativeMinerals;
  public
    { Public declarations }
    procedure GetIniFile;
    procedure SetIniFile;
  end;

var
  fmNormMain: TfmNormMain;

implementation

uses normvarb, norm_min, normdesi, normtern,
     Nrm_ShtIm2, About, Nrm_dm_acs, Nrm_ShtImTemplate2;

{$R *.DFM}
var
  ImportForm : TfmSheetImport;
  ImportFormTemplate : TfmSheetImportTemplate;

procedure Names;
begin
            B[1]:='SiO2 ';
            B[2]:=' Qz ';
            B[3]:='An/Pl  Wt pct';
            B[4]:='Mol pct';
            B[5]:='TiO2 ';
            B[6]:=' Co ';
            B[7]:='ZrO2 ';
            B[8]:=' Z  ';
            B[9]:='Fa/Ol  Wt pct';
            B[10]:='Mol pct';
            B[11]:='Al2O3';
            B[12]:=' Or ';
            B[13]:='Cr2O3';
            B[14]:=' Pl ';
            B[15]:='En/Hy  Wt pct';
            B[16]:='Mol pct';
            B[17]:='Fe2O3';
            B[18]:='(Ab)';
            B[19]:='FeO  ';
            B[20]:='(An)';
            B[21]:='Difndx Wt pct';
            B[22]:='Cat pct';
            B[23]:='MnO  ';
            B[24]:=' Lc ';
            B[25]:='NiO  ';
            B[26]:=' Ne ';
            B[27]:='QZ-Ab-Or diagram       ';
            B[28]:='MgO  ';
            B[29]:=' Ks ';
            B[30]:='CaO  ';
            B[31]:=' Hl ';
            B[32]:='Wt pct  Qz';
            B[33]:=' Ab ';
            B[34]:=' Or ';
            B[35]:='SrO  ';
            B[36]:=' Th ';
            B[37]:='BaO  ';
            B[38]:='Mol pct Qz';
            B[39]:=' Ab ';
            B[40]:=' Or ';
            B[41]:='Na2O ';

            B[43]:='K2O  ';
            B[44]:=' Ns ';
            B[45]:='Qz-Ne-Ks diagram';
            B[46]:='P2O5 ';
            B[47]:=' Ks ';
            B[48]:='H2O+ ';
            B[49]:=' Wo ';
            B[50]:='Wt pct  Qz';
            B[51]:=' Ne ';
            B[52]:=' Ks ';
            B[53]:='H2O- ';
            B[54]:=' Di ';
            B[55]:='SO3  ';
            B[56]:='(Wo)';
            B[57]:='Mol pct Qz';
            B[58]:=' Ne ';
            B[59]:=' Ks ';
            B[60]:='S    ';
            B[61]:='(En)';
            B[62]:='Cl   ';
            B[63]:='(Fs)';
            B[64]:='An-Ab-Or diagram';
            B[65]:='F    ';
            B[66]:=' Hy ';
            B[67]:='CO2  ';
            B[68]:='(En)';
            B[69]:='Wt pct  An';
            B[70]:=' Ab ';
            B[71]:=' Or ';
            B[72]:='(Fs)';
            B[73]:='Total';
            B[74]:=' Ol ';
            B[75]:='Mol pct An';
            B[76]:=' Ab ';
            B[77]:=' Or ';
            B[78]:='-H2O ';
            B[79]:='(Fo)';
            B[80]:='(Fa)';
            B[81]:=' A-F-M diagram';
            B[115]:='Watson & Harrison  M     ';
            B[116]:='De la Roche R1           ';
            B[117]:='De la Roche R2           ';
            B[118]:='Chem. Index Alteration   ';
   case InOption of
     '0' : begin
            B[42]:=' Ac ';

            B[82]:=' Cs ';
            B[83]:=' Mt ';
            B[84]:='Wt pct Alk';
            B[85]:=' Fe ';
            B[86]:=' Mg ';
            B[87]:=' Cm ';
            B[88]:=' Il ';
            B[89]:='Mol pct Alk';
            B[90]:=' Fe ';
            B[91]:=' Mg ';
            B[92]:=' Hm ';
            B[93]:=' Sp ';
            B[94]:='(Na+K)/Al          Atm wt %';
            B[95]:='Grmatm %  ';
            B[96]:=' Pf ';
            B[97]:=' Ru ';
            B[98]:='(Fe+Mn)/(Fe+Mn+Mg) Atm wt %';
            B[99]:='Grmatm %  ';
            B[100]:=' Ap ';
            B[101]:=' Fl ';
            B[102]:='Na2O/(Na2O+K2O)        wt %';
            B[103]:='   Mol %  ';
            B[104]:=' Py ';
            B[105]:=' Cc ';
            B[106]:='FeO/(FeO+Fe2O3)        wt %';
            B[107]:='   Mol %  ';
            B[108]:='Salic';
            B[109]:='Wright`s alkalinity index';
            B[110]:='Femic';
            B[111]:='Total';
         end;
   '1','2' : begin
            B[42]:=' Bi ';

            B[82]:=' Ho ';
            B[83]:='(Ri)';
            B[84]:='Wt pct Alk';
            B[85]:=' Fe ';
            B[86]:=' Mg ';
            B[87]:='(Act)';
            B[88]:='(Ed)';
            B[89]:='Mol pct Alk';
            B[90]:=' Fe ';
            B[91]:=' Mg ';
            B[92]:='Spnl';
            B[93]:=' Cs ';
            B[94]:='(Na+K)/Al          Atm wt %';
            B[95]:='Grmatm %  ';
            B[96]:=' Mt ';
            B[97]:=' Cm ';
            B[98]:='(Fe+Mn)/(Fe+Mn+Mg) Atm wt %';
            B[99]:='Grmatm %  ';
            B[100]:=' Hm ';
            B[101]:=' Sp ';
            B[102]:='Na2O/(Na2O+K2O)        wt %';
            B[103]:='   Mol %  ';
            B[104]:=' Ru ';
            B[105]:=' Ap ';
            B[106]:='FeO/(FeO+Fe2O3)        wt %';
            B[107]:='   Mol %  ';
            B[108]:=' Fl ';
            B[109]:=' Py ';
            B[110]:='Wright`s alkalinity index';
            B[111]:=' Cc ';
            B[112]:='Salic';
            B[113]:='Femic';
            B[114]:='Total';
         end;
   end;{case}
end;{proc Names}


procedure CIPWPrint;
var
  Lst : textfile;
begin
   AssignPrn(Lst);
   Rewrite(Lst);
   Printer.Canvas.Font.Name := 'Courier';
   Printer.Canvas.Font.Style := [fsBold];
   Printer.Canvas.Font.Size := 7;
   Writeln(Lst,' ');
   Writeln(Lst,'');
   Writeln(Lst,'');
   Write(Lst,' ':10,dmNrm.cdsNormChemSampleNum.AsString+' ':4,DES:60,' ':8,'CIPW-norm');
   Writeln(Lst,'  ');
   Printer.Canvas.Font.Style := [];
   {
   Writeln(Lst,StringOf(92,32),'Printed on ',DateString);
   }
   if Toggle100=1 then
   begin
     Write(Lst,'Mineral wt % recalculated to 100%');
   end;
   Writeln(Lst,'');
   Writeln(Lst,A14);
   Writeln(Lst,A15);
   Writeln(Lst,A16);
   Writeln(Lst,'');
   Writeln(Lst, B[1]:15,OSI:12:2,CSI:9:2,SI:9:2, B[2]:10,WQZ:9:2,QZ:9:2, B[3]:19,WANPL:9:2, B[4]:13,ANPL:9:2);
   Writeln(Lst, B[5]:15,OTI:12:2,CTI:9:2,TI:9:2, B[6]:10,WCO:9:2,CO:9:2);
   Writeln(Lst, B[7]:15,OZR:12:2,CZR:9:2,ZR:9:2, B[8]:10,WZN:9:2,ZN:9:2, B[9]:19,WFAOL:9:2, B[10]:13,FAOL:9:2);
   Writeln(Lst,B[11]:15,OAL:12:2,CAL:9:2,AL:9:2,B[12]:10,WOR:9:2,ORT:9:2);
   Writeln(Lst,B[13]:15,OCR:12:2,CCR:9:2,NCR:9:2,B[14]:10,WPL:9:2,NPL:9:2,B[15]:19,WENHY:9:2,B[16]:13,ENHY:9:2);
   Writeln(Lst,B[17]:15,OF3:12:2,CF3:9:2,F3:9:2,B[18]:10,WAB:9:2,AB:9:2);
   Writeln(Lst,B[19]:15,OF2:12:2,CF2:9:2,F2:9:2,B[20]:10,WAN:9:2,AN:9:2,B[21]:19,WDIX:9:2,B[22]:13,DIX:9:2);
   Writeln(Lst,B[23]:15,OMN:12:2,CMN:9:2,B[24]:19,WLC:9:2,LC:9:2);
   Writeln(Lst,B[25]:15,ONI:12:2,CNI:9:2,B[26]:19,WNE:9:2,NE:9:2,B[27]:44);
   Writeln(Lst,B[28]:15,OMG:12:2,CMG:9:2,MG:9:2,B[29]:10,WKP:9:2,KP:9:2);
   Write(Lst,B[30]:15,OCA:12:2,CCA:9:2,CA:9:2,B[31]:10,WHL:9:2,HL:9:2);
   Writeln(Lst,B[32]:15,WTQZ:7:2,B[33]:7,WTAB:7:2,B[34]:7,WTOR:7:2);
   Writeln(Lst,B[35]:15,OSR:12:2,CSR:9:2,B[36]:19,WTH:9:2,TH:9:2);
   Writeln(Lst,B[37]:15,OBA:12:2,CBA:9:2,B[38]:52,TQZ:7:2,B[39]:7,TTAB:7:2,B[40]:7,TOR:7:2);
   Writeln(Lst,B[41]:15,ONA:12:2,CNA:9:2,NA:9:2,B[42]:10,WAC:9:2,AC:9:2);
   Writeln(Lst,B[43]:15,OKO:12:2,CKO:9:2,KO:9:2,B[44]:10,WNS:9:2,NS:9:2,B[45]:37);
   Writeln(Lst,B[46]:15,OPO:12:2,CPO:9:2,PO:9:2,B[47]:10,WKS:9:2,KS:9:2);
   Writeln(Lst,B[48]:15,OHP:12:2,B[49]:28,WWO:9:2,WO:9:2,B[50]:15,WTTQZ:7:2,B[51]:7,WTNE:7:2,B[52]:7,WTKP:7:2);
   Writeln(Lst,B[53]:15,OH2OM:12:2,B[54]:28,WDI:9:2,DI:9:2);
   Write(Lst,B[55]:15,OSO:12:2,CSO:9:2,SO:9:2,B[56]:10,WWODI:9:2,WODI:9:2);
   Writeln(Lst,B[57]:15,TTQZ:7:2,B[58]:7,TNE:7:2,B[59]:7,TKP:7:2);
   Writeln(Lst,B[60]:15,OSU:12:2,CSU:9:2,SU:9:2,B[61]:10,WENDI:9:2,ENDI:9:2);
   Writeln(Lst,B[62]:15,OCL:12:2,CCL:9:2,CL:9:2,B[63]:10,WFSDI:9:2,FSDI:9:2,B[64]:37);
   Writeln(Lst,B[65]:15,OFU:12:2,CFU:9:2,FU:9:2,B[66]:10,WHY:9:2,HY:9:2);
   Write(Lst,B[67]:15,OCD:12:2,CCD:9:2,CD:9:2,B[68]:10,WEN:9:2,EN:9:2);
   Writeln(Lst,B[69]:15,WTFAN:7:2,B[70]:7,WTFAB:7:2,B[71]:7,WTFOR:7:2);
   Writeln(Lst,' ':45,B[72]:10,WFS:9:2,FS:9:2);
   Write(Lst,B[73]:15,OXTOT:12:2,CTTOT:9:2,B[74]:19,WOL:9:2,OL:9:2);
   Writeln(Lst,B[75]:15,TFAN:7:2,B[76]:7,TFAB:7:2,B[77]:7,TFOR:7:2);
   Writeln(Lst,B[78]:15,NOH2O:12:2,B[79]:28,WFO:9:2,FO:9:2);
   Writeln(Lst,' ':45,B[80]:10,WFA:9:2,FA:9:2,B[81]:35);
   Writeln(Lst,' ':45,B[82]:10,WCS:9:2,CS:9:2);
   Writeln(Lst,' ':45,B[83]:10,WMT:9:2,MT:9:2,B[84]:15,WAFMA:7:2,B[85]:7,WAFMF:7:2,B[86]:7,WAFMM:7:2);
   Writeln(Lst,' ':45,B[87]:10,WCM:9:2,CM:9:2);
   Writeln(Lst,' ':45,B[88]:10,WIL:9:2,IL:9:2,B[89]:15,AFMA:7:2,B[90]:7,AFMF:7:2,B[91]:7,AFMM:7:2);
   Writeln(Lst,' ':45,B[92]:10,WHM:9:2,HM:9:2);
   Writeln(Lst,' ':45,B[93]:10,WSP:9:2,SP:9:2,B[94]:29,WAGRAT:6:2,B[95]:11,AGRAT:6:2);
   Writeln(Lst,' ':45,B[96]:10,WPF:9:2,PF:9:2);
   Writeln(Lst,' ':45,B[97]:10,WRU:9:2,RU:9:2,B[98]:29,WFMRAT:6:2,B[99]:11,FMRAT:6:2);
   Writeln(Lst,' ':45,B[100]:10,WAP:9:2,AP:9:2);
   Writeln(Lst,' ':45,B[101]:10,WFL:9:2,FL:9:2,B[102]:29,WALRAT:6:2,B[103]:11,ALRAT:6:2);
   Writeln(Lst,' ':45,B[104]:10,WPY:9:2,PY:9:2);
   Writeln(Lst,' ':45,B[105]:10,WCC:9:2,CC:9:2,B[106]:29,WOXRAT:6:2,B[107]:11,OXRAT:6:2);
   Writeln(Lst,'');
   Writeln(Lst,' ':50,B[108]:5,WSALIC:9:2,SALIC:9:2);
   Writeln(Lst,' ':50,B[110]:5,WFEMIC:9:2,FEMIC:9:2);
   Writeln(Lst);
   Writeln(Lst,' ':50,B[111]:5,WTOTAL:9:2,TOTAL:9:2);
   Writeln(Lst);
   Writeln(Lst);
   Writeln(Lst,' ':73,B[109]:27,WALKIN:8:2);
   Writeln(Lst,' ':73,B[115]:27,WatM:8:2);
   Writeln(Lst,' ':73,B[116]:27,R1:8:2);
   Writeln(Lst,' ':73,B[117]:27,R2:8:2);
   Writeln(Lst,' ':73,B[118]:27,CIA:8:2);
   Writeln(Lst,char(13));
   System.CloseFile(Lst);
end;{proc CIPWPrint}



procedure MesoPrint;
var
  Lst : textfile;
begin
   AssignPrn(Lst);
   Rewrite(Lst);
   Printer.Canvas.Font.Name := 'Courier';
   Printer.Canvas.Font.Style := [fsBold];
   Printer.Canvas.Font.Size := 7;
   Writeln(Lst);
   Writeln(Lst);
   Writeln(Lst);
   Write(Lst,' ':10,dmNrm.cdsNormChemSampleNum.AsString+' ':4,DES:60,' ':8,'Meso-norm');
   Writeln(Lst,' ');
   Printer.Canvas.Font.Style := [];
   {
   Writeln(Lst,StringOf(92,32),'Printed on ',DateString);
   }
   if Toggle100=1 then
   begin
     Write(Lst,'Mineral wt % recalculated to 100%');
   end;
   Writeln(Lst);
   Writeln(Lst,A14);
   Writeln(Lst,A15);
   Writeln(Lst,A16);
   Writeln(Lst);
   Writeln(Lst, B[1]:15,OSI:12:2,CSI:9:2,SI:9:2, B[2]:10,WQZ:9:2,QZ:9:2, B[3]:19,WANPL:9:2, B[4]:13,ANPL:9:2);
   Writeln(Lst, B[5]:15,OTI:12:2,CTI:9:2,TI:9:2, B[6]:10,WCO:9:2,CO:9:2);
   Writeln(Lst, B[7]:15,OZR:12:2,CZR:9:2,ZR:9:2, B[8]:10,WZN:9:2,ZN:9:2, B[9]:19,WFAOL:9:2, B[10]:13,FAOL:9:2);
   Writeln(Lst,B[11]:15,OAL:12:2,CAL:9:2,AL:9:2,B[12]:10,WOR:9:2,ORT:9:2);
   Writeln(Lst,B[13]:15,OCR:12:2,CCR:9:2,NCR:9:2,B[14]:10,WPL:9:2,NPL:9:2,B[15]:19,WENHY:9:2,B[16]:13,ENHY:9:2);
   Writeln(Lst,B[17]:15,OF3:12:2,CF3:9:2,F3:9:2,B[18]:10,WAB:9:2,AB:9:2);
   Writeln(Lst,B[19]:15,OF2:12:2,CF2:9:2,F2:9:2,B[20]:10,WAN:9:2,AN:9:2,B[21]:19,WDIX:9:2,B[22]:13,DIX:9:2);
   Writeln(Lst,B[23]:15,OMN:12:2,CMN:9:2,B[24]:19,WLC:9:2,LC:9:2);
   Writeln(Lst,B[25]:15,ONI:12:2,CNI:9:2,B[26]:19,WNE:9:2,NE:9:2,B[27]:44);
   Writeln(Lst,B[28]:15,OMG:12:2,CMG:9:2,MG:9:2,B[29]:10,WKP:9:2,KP:9:2);
   Write(Lst,B[30]:15,OCA:12:2,CCA:9:2,CA:9:2,B[31]:10,WHL:9:2,HL:9:2);
   Writeln(Lst,B[32]:15,WTQZ:7:2,B[33]:7,WTAB:7:2,B[34]:7,WTOR:7:2);
   Writeln(Lst,B[35]:15,OSR:12:2,CSR:9:2,B[36]:19,WTH:9:2,TH:9:2);
   Writeln(Lst,B[37]:15,OBA:12:2,CBA:9:2,B[38]:52,TQZ:7:2,B[39]:7,TTAB:7:2,B[40]:7,TOR:7:2);
   Writeln(Lst,B[41]:15,ONA:12:2,CNA:9:2,NA:9:2,B[42]:10,WBI:9:2,BI:9:2);
   Writeln(Lst,B[43]:15,OKO:12:2,CKO:9:2,KO:9:2,B[44]:10,WNS:9:2,NS:9:2,B[45]:37);
   Writeln(Lst,B[46]:15,OPO:12:2,CPO:9:2,PO:9:2,B[47]:10,WKS:9:2,KS:9:2);
   Writeln(Lst,B[48]:15,OHP:12:2,B[49]:28,WWO:9:2,WO:9:2,B[50]:15,WTTQZ:7:2,B[51]:7,WTNE:7:2,B[52]:7,WTKP:7:2);
   Writeln(Lst,B[53]:15,OH2OM:12:2,B[54]:28,WDI:9:2,DI:9:2);
   Write(Lst,B[55]:15,OSO:12:2,CSO:9:2,SO:9:2,B[56]:10,WWODI:9:2,WODI:9:2);
   Writeln(Lst,B[57]:15,TTQZ:7:2,B[58]:7,TNE:7:2,B[59]:7,TKP:7:2);
   Writeln(Lst,B[60]:15,OSU:12:2,CSU:9:2,SU:9:2,B[61]:10,WENDI:9:2,ENDI:9:2);
   Writeln(Lst,B[62]:15,OCL:12:2,CCL:9:2,CL:9:2,B[63]:10,WFSDI:9:2,FSDI:9:2,B[64]:37);
   Writeln(Lst,B[65]:15,OFU:12:2,CFU:9:2,FU:9:2,B[66]:10,WHY:9:2,HY:9:2);
   Write(Lst,B[67]:15,OCD:12:2,CCD:9:2,CD:9:2,B[68]:10,WEN:9:2,EN:9:2);
   Writeln(Lst,B[69]:15,WTFAN:7:2,B[70]:7,WTFAB:7:2,B[71]:7,WTFOR:7:2);
   Writeln(Lst,' ':45,B[72]:10,WFS:9:2,FS:9:2);
   Write(Lst,B[73]:15,OXTOT:12:2,CTTOT:9:2,B[74]:19,WOL:9:2,OL:9:2);
   Writeln(Lst,B[75]:15,TFAN:7:2,B[76]:7,TFAB:7:2,B[77]:7,TFOR:7:2);
   Writeln(Lst,B[78]:15,NOH2O:12:2,B[79]:28,WFO:9:2,FO:9:2);
   Writeln(Lst,' ':45,B[80]:10,WFA:9:2,FA:9:2,B[81]:35);
   Writeln(Lst,' ':45,B[82]:10,WHO:9:2,HO:9:2);
   Writeln(Lst,' ':45,B[83]:10,WRI:9:2,RI:9:2,B[84]:15,WAFMA:7:2,B[85]:7,WAFMF:7:2,B[86]:7,WAFMM:7:2);
   Writeln(Lst,' ':45,B[87]:10,WACT:9:2,ACT:9:2);
   Writeln(Lst,' ':45,B[88]:10,WED:9:2,ED:9:2,B[89]:15,AFMA:7:2,B[90]:7,AFMF:7:2,B[91]:7,AFMM:7:2);
   Writeln(Lst,' ':45,B[92]:10,WSPIN:9:2,SPIN:9:2);
   Writeln(Lst,' ':45,B[93]:10,WCS:9:2,CS:9:2,B[94]:29,WAGRAT:6:2,B[95]:11,AGRAT:6:2);
   Writeln(Lst,' ':45,B[96]:10,WMT:9:2,MT:9:2);
   Writeln(Lst,' ':45,B[97]:10,WCM:9:2,CM:9:2,B[98]:29,WFMRAT:6:2,B[99]:11,FMRAT:6:2);
   Writeln(Lst,' ':45,B[100]:10,WHM:9:2,HM:9:2);
   Writeln(Lst,' ':45,B[101]:10,WSP:9:2,SP:9:2,B[102]:29,WALRAT:6:2,B[103]:11,ALRAT:6:2);
   Writeln(Lst,' ':45,B[104]:10,WRU:9:2,RU:9:2);
   Writeln(Lst,' ':45,B[105]:10,WAP:9:2,AP:9:2,B[106]:29,WOXRAT:6:2,B[107]:11,OXRAT:6:2);
   Writeln(Lst,' ':45,B[108]:10,WFL:9:2,Fl:9:2);
   Writeln(Lst,' ':45,B[109]:10,WPY:9:2,PY:9:2);
   Writeln(Lst,' ':45,B[111]:10,WCC:9:2,CC:9:2);
   Writeln(Lst);
   Writeln(Lst,' ':50,B[112]:5,WSALIC:9:2,SALIC:9:2);
   Writeln(Lst,' ':50,B[113]:5,WFEMIC:9:2,FEMIC:9:2);
   Writeln(Lst);
   Writeln(Lst,' ':50,B[114]:5,WTOTAL:9:2,TOTAL:9:2);
   Writeln(Lst);
   Writeln(Lst);
   Writeln(Lst,' ':73,B[110]:27,WALKIN:8:2);
   Writeln(Lst,' ':73,B[115]:27,WatM:8:2);
   Writeln(Lst,' ':73,B[116]:27,R1:8:2);
   Writeln(Lst,' ':73,B[117]:27,R2:8:2);
   Writeln(Lst,' ':73,B[118]:27,CIA:8:2);
   Writeln(Lst,char(13));
   System.CloseFile(Lst);
end;{proc MesoPrint}


procedure TfmNormMain.Record_To_Memory;
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
   OSI:=dmNrm.cdsNormChemSIO2.AsFloat;
   OTI:=dmNrm.cdsNormChemTIO2.AsFloat;
   OZR:=dmNrm.cdsNormChemZRO2.AsFloat;
   OAL:=dmNrm.cdsNormChemAL2O3.AsFloat;
   OCR:=dmNrm.cdsNormChemCR2O3.AsFloat;
   OF3:=dmNrm.cdsNormChemFE2O3.AsFloat;
   OF2:=dmNrm.cdsNormChemFEO.AsFloat;
   OMN:=dmNrm.cdsNormChemMNO.AsFloat;
   ONI:=dmNrm.cdsNormChemNIO.AsFloat;
   OMG:=dmNrm.cdsNormChemMGO.AsFloat;
   OCA:=dmNrm.cdsNormChemCAO.AsFloat;
   OSR:=dmNrm.cdsNormChemSRO.AsFloat;
   OBA:=dmNrm.cdsNormChemBAO.AsFloat;
   ONA:=dmNrm.cdsNormChemNA2O.AsFloat;
   OKO:=dmNrm.cdsNormChemK2O.AsFloat;
   OPO:=dmNrm.cdsNormChemP2O5.AsFloat;
   OHP:=dmNrm.cdsNormChemLOI.AsFloat;
   OH2OM:=dmNrm.cdsNormChemH2OM.AsFloat;
   OSO:=dmNrm.cdsNormChemSO3.AsFloat;
   OSU:=dmNrm.cdsNormChemS.AsFloat;
   OCL:=dmNrm.cdsNormChemCL.AsFloat;
   OFU:=dmNrm.cdsNormChemF.AsFloat;
   OCD:=dmNrm.cdsNormChemCO2.AsFloat;
   OXTOT:=OSI+OTI+OAL+OF3+OF2+OMN+OMG+OCA+ONA+OKO+OPO+OHP+OH2OM+OCD;
   OXTOT:=OXTOT+OCR+OSR+OBA+OCL+OFU+ONI+OZR+OSU+OSO;
   dmNrm.cdsNormChem.Edit;
   dmNrm.cdsNormChemTOTAL.AsFloat := OXTOT;
   dmNrm.cdsNormChem.Post;
end;{proc Record_To_Memory}


procedure TfmNormMain.Memory_To_Record;
{------------------------------------------------------------
 Output minerals to datafile
 ------------------------------------------------------------}
begin
   dmNrm.cdsNormsMinQZ.AsFloat := QZ;
   dmNrm.cdsNormsMinCO.AsFloat := CO;
   dmNrm.cdsNormsMinZ.AsFloat := ZN;
   dmNrm.cdsNormsMinOR.AsFloat := ORT;
   dmNrm.cdsNormsMinPL.AsFloat := AB+AN;
   dmNrm.cdsNormsMinPLAB.AsFloat := AB;
   dmNrm.cdsNormsMinPLAN.AsFloat := AN;
   dmNrm.cdsNormsMinLC.AsFloat := LC;
   dmNrm.cdsNormsMinNE.AsFloat := NE;
   dmNrm.cdsNormsMinKP.AsFloat := KP;
   dmNrm.cdsNormsMinHL.AsFloat := HL;
   dmNrm.cdsNormsMinTH.AsFloat := TH;
   dmNrm.cdsNormsMinAC.AsFloat := AC;
   dmNrm.cdsNormsMinNS.AsFloat := NS;
   dmNrm.cdsNormsMinKS.AsFloat := KS;
   dmNrm.cdsNormsMinWO.AsFloat := WO;
   dmNrm.cdsNormsMinDI.AsFloat := DI;
   dmNrm.cdsNormsMinDIEN.AsFloat := ENDI;
   dmNrm.cdsNormsMinDIFS.AsFloat := FSDI;
   dmNrm.cdsNormsMinDIWO.AsFloat := WODI;
   dmNrm.cdsNormsMinHY.AsFloat := HY;
   dmNrm.cdsNormsMinHYEN.AsFloat := EN;
   dmNrm.cdsNormsMinHYFS.AsFloat := FS;
   dmNrm.cdsNormsMinOL.AsFloat := OL;
   dmNrm.cdsNormsMinOLFO.AsFloat := FO;
   dmNrm.cdsNormsMinOLFA.AsFloat := FA;
   dmNrm.cdsNormsMinCS.AsFloat := CS;
   dmNrm.cdsNormsMinMT.AsFloat := MT;
   dmNrm.cdsNormsMinCM.AsFloat := CM;
   dmNrm.cdsNormsMinIL.AsFloat := IL;
   dmNrm.cdsNormsMinHM.AsFloat := HM;
   dmNrm.cdsNormsMinSP.AsFloat := SP;
   dmNrm.cdsNormsMinPF.AsFloat := PF;
   dmNrm.cdsNormsMinRU.AsFloat := RU;
   dmNrm.cdsNormsMinAP.AsFloat := AP;
   dmNrm.cdsNormsMinFL.AsFloat := FL;
   dmNrm.cdsNormsMinPY.AsFloat := PY;
   dmNrm.cdsNormsMinCC.AsFloat := CC;
   dmNrm.cdsNormsMinHO.AsFloat := HO;
   dmNrm.cdsNormsMinHOACT.AsFloat := ACT;
   dmNrm.cdsNormsMinHORI.AsFloat := RI;
   dmNrm.cdsNormsMinHOED.AsFloat := ED;
   dmNrm.cdsNormsMinBI.AsFloat := BI;
   dmNrm.cdsNormsMinSPNL.AsFloat := SPIN;
   dmNrm.cdsNormsMinSALIC.AsFloat := SALIC;
   dmNrm.cdsNormsMinFEMIC.AsFloat := FEMIC;
   dmNrm.cdsNormsMinTOTAL.AsFloat := TOTAL;

   dmNrm.cdsNormsMinWQZ.AsFloat := WQZ;
   dmNrm.cdsNormsMinWCO.AsFloat := WCO;
   dmNrm.cdsNormsMinWZ.AsFloat := WZN;
   dmNrm.cdsNormsMinWOR.AsFloat := WOR;
   dmNrm.cdsNormsMinWPL.AsFloat := WAB+WAN;
   dmNrm.cdsNormsMinWPLAB.AsFloat := WAB;
   dmNrm.cdsNormsMinWPLAN.AsFloat := WAN;
   dmNrm.cdsNormsMinWLC.AsFloat := WLC;
   dmNrm.cdsNormsMinWNE.AsFloat := WNE;
   dmNrm.cdsNormsMinWKP.AsFloat := WKP;
   dmNrm.cdsNormsMinWHL.AsFloat := WHL;
   dmNrm.cdsNormsMinWTH.AsFloat := WTH;
   dmNrm.cdsNormsMinWAC.AsFloat := WAC;
   dmNrm.cdsNormsMinWNS.AsFloat := WNS;
   dmNrm.cdsNormsMinWKS.AsFloat := WKS;
   dmNrm.cdsNormsMinWWO.AsFloat := WWO;
   dmNrm.cdsNormsMinWDI.AsFloat := WDI;
   dmNrm.cdsNormsMinWDIEN.AsFloat := WENDI;
   dmNrm.cdsNormsMinWDIFS.AsFloat := WFSDI;
   dmNrm.cdsNormsMinWDIWO.AsFloat := WWODI;
   dmNrm.cdsNormsMinWHY.AsFloat := WHY;
   dmNrm.cdsNormsMinWHYEN.AsFloat := WEN;
   dmNrm.cdsNormsMinWHYFS.AsFloat := WFS;
   dmNrm.cdsNormsMinWOL.AsFloat := WOL;
   dmNrm.cdsNormsMinWOLFO.AsFloat := WFO;
   dmNrm.cdsNormsMinWOLFA.AsFloat := WFA;
   dmNrm.cdsNormsMinWCS.AsFloat := WCS;
   dmNrm.cdsNormsMinWMT.AsFloat := WMT;
   dmNrm.cdsNormsMinWCM.AsFloat := WCM;
   dmNrm.cdsNormsMinWIL.AsFloat := WIL;
   dmNrm.cdsNormsMinWHM.AsFloat := WHM;
   dmNrm.cdsNormsMinWSP.AsFloat := WSP;
   dmNrm.cdsNormsMinWPF.AsFloat := WPF;
   dmNrm.cdsNormsMinWRU.AsFloat := WRU;
   dmNrm.cdsNormsMinWAP.AsFloat := WAP;
   dmNrm.cdsNormsMinWFL.AsFloat := WFL;
   dmNrm.cdsNormsMinWPY.AsFloat := WPY;
   dmNrm.cdsNormsMinWCC.AsFloat := WCC;
   dmNrm.cdsNormsMinWHO.AsFloat := WHO;
   dmNrm.cdsNormsMinWBI.AsFloat := WBI;
   dmNrm.cdsNormsMinWHOACT.AsFloat := WACT;
   dmNrm.cdsNormsMinWHOED.AsFloat := WED;
   dmNrm.cdsNormsMinWHORI.AsFloat := WRI;
   dmNrm.cdsNormsMinWSPNL.AsFloat := WSPIN;
   dmNrm.cdsNormsMinWSALIC.AsFloat := WSALIC;
   dmNrm.cdsNormsMinWFEMIC.AsFloat := WFEMIC;
   dmNrm.cdsNormsMinWTOTAL.AsFloat := WTOTAL;

   {Qz-Ab-Or}
   dmNrm.cdsNormsMinQzAbOrTQZ.AsFloat := TQZ;
   dmNrm.cdsNormsMinQzAbOrTAB.AsFloat := TTAB;
   dmNrm.cdsNormsMinQzAbOrTOR.AsFloat := TOR;
   dmNrm.cdsNormsMinQzAbOrWTQZ.AsFloat := WTQZ;
   dmNrm.cdsNormsMinQzAbOrWTAB.AsFloat := WTAB;
   dmNrm.cdsNormsMinQzAbOrWTOR.AsFloat := WTOR;
   {Qz-Ne-Kp}
   dmNrm.cdsNormsMinQzNeKpTQZ.AsFloat := TTQZ;
   dmNrm.cdsNormsMinQzNeKpTNE.AsFloat := TNE;
   dmNrm.cdsNormsMinQzNeKpTKP.AsFloat := TKP;
   dmNrm.cdsNormsMinQzNeKpWTQZ.AsFloat := WTTQZ;
   dmNrm.cdsNormsMinQzNeKpWTNE.AsFloat := WTNE;
   dmNrm.cdsNormsMinQzNeKpWTKP.AsFloat := WTKP;
   {Or-Ab-An}
   dmNrm.cdsNormsMinOrAbAnTOR.AsFloat := TFOR;
   dmNrm.cdsNormsMinOrAbAnTAB.AsFloat := TFAB;
   dmNrm.cdsNormsMinOrAbAnTAN.AsFloat := TFAN;
   dmNrm.cdsNormsMinOrAbAnWTOR.AsFloat := WTFOR;
   dmNrm.cdsNormsMinOrAbAnWTAB.AsFloat := WTFAB;
   dmNrm.cdsNormsMinOrAbAnWTAN.AsFloat := WTFAN;

   dmNrm.cdsNormsMinAFMA.AsFloat := AFMA;
   dmNrm.cdsNormsMinAFMF.AsFloat := AFMF;
   dmNrm.cdsNormsMinAFMM.AsFloat := AFMM;
   dmNrm.cdsNormsMinWAFMA.AsFloat := WAFMA;
   dmNrm.cdsNormsMinWAFMF.AsFloat := WAFMF;
   dmNrm.cdsNormsMinWAFMM.AsFloat := WAFMM;
   dmNrm.cdsNormsMinAgpaitic.AsFloat := AGRAT;
   dmNrm.cdsNormsMinWAgpaitic.AsFloat := WAGRAT;
   dmNrm.cdsNormsMinFeMgRat.AsFloat := FMRAT;
   dmNrm.cdsNormsMinWFeMgRat.AsFloat := WFMRAT;
   dmNrm.cdsNormsMinAlkRat.AsFloat := ALRAT;
   dmNrm.cdsNormsMinWAlkRat.AsFloat := WALRAT;
   dmNrm.cdsNormsMinOxidRat.AsFloat := OXRAT;
   dmNrm.cdsNormsMinWOxidRat.AsFloat := WOXRAT;
   dmNrm.cdsNormsMinWrightsAlk.AsFloat := WALKIN;
   dmNrm.cdsNormsMinTotalAlk.AsFloat := ALK;
   dmNrm.cdsNormsMinDiffInd.AsFloat := DIX;
   dmNrm.cdsNormsMinWDiffInd.AsFloat := WDIX;
   dmNrm.cdsNormsMinWatsonM.AsFloat := WATM;
   dmNrm.cdsNormsMinWANPL.AsFloat := WANPL;
   dmNrm.cdsNormsMinANPL.AsFloat := ANPL;
   dmNrm.cdsNormsMinWFAOL.AsFloat := WFAOL;
   dmNrm.cdsNormsMinFAOL.AsFloat := FAOL;
   dmNrm.cdsNormsMinWENHY.AsFloat := WENHY;
   dmNrm.cdsNormsMinENHY.AsFloat := ENHY;
   dmNrm.cdsNormsMinR1.AsFloat := R1;
   dmNrm.cdsNormsMinR2.AsFloat := R2;
   dmNrm.cdsNormsMinChemIndAlt.AsFloat := CIA;
   dmNrm.cdsNormsMinRosKor_D1.AsFloat := RoserKorschD1;
   dmNrm.cdsNormsMinRosKor_D2.AsFloat := RoserKorschD2;
   dmNrm.cdsNormsMinRosKor_D3.AsFloat := RoserKorschD3;
   dmNrm.cdsNormsMinRosKor_D4.AsFloat := RoserKorschD4;
   dmNrm.cdsNormsMinPeralumInd.AsFloat := ACNK;
   dmNrm.cdsNormsMinDebLefor_A.AsFloat := DebonLefortA;
   dmNrm.cdsNormsMinDebLefor_B.AsFloat := DebonLefortB;
end;{proc Memory_To_Record}


procedure SundryRatios;
var
   ALCA            : double;
begin
   if OXTOT>10 then
   begin
      NOH2O:=OXTOT-OHP-OH2OM;
      ALK:=ONA+OKO;
      F23:=OF2+OF3;
      if F23>0 then begin
         WOXRAT:=OF2*100/F23;
         F23:=OF2+OF3*0.8998;
         SUM:=100.0/(ALK+F23+OMG);
         WAFMA:=ALK*SUM;
         WAFMF:=F23*SUM;
         WAFMM:=OMG*SUM;
      end
      else begin
        WOXRAT:=0.0;
        WAFMA:=0.0;
        WAFMF:=0.0;
        WAFMM:=0.0;
      end;
      if ALK>0 then WALRAT:=ONA*100/ALK else WALRAT:=0.0;
      ALK1:=ALK;
      WALKIN:=0.0;
      if ONA>0.0 then OKONA:=OKO/ONA
                 else OKONA:=0.0;
      if ALK>0 then
      begin
         if (OSI > 50.0) and ((OKONA) > 1.0) and ((OKONA) < 2.5) then ALK1:=2.0*ONA;
         ALCA:=OAL+OCA;
         WALKIN:=(ALCA+ALK1)/(ALCA-ALK1);
      end;{if}
      RoserKorschD1 := 0.0;
      RoserKorschD2 := 0.0;
      RoserKorschD3 := 0.0;
      RoserKorschD4 := 0.0;
      if OAL > 0.0 then
      begin
        RoserKorschD1 := -1.773*OTI + 0.607*OAL +0.76*(OF3+OF2/0.8998) - 1.5*OMG
                        + 0.616*OCA + 0.509*ONA - 1.224*OKO - 9.09;
        RoserKorschD2 := 0.445*OTI + 0.07*OAL - 0.25*(OF3+OF2/0.8998) - 1.142*OMG
                        + 0.438*OCA + 1.475*ONA + 1.426*OKO -6.861;
        RoserKorschD3 := 30.638*OTI/OAL - 12.541*(OF3+OF2/0.8998)/OAL
                        + 7.329*OMG/OAL + 12.031*ONA/OAL + 35.402*OKO/OAL - 6.382;
        RoserKorschD4 := 56.500*OTI/OAL - 10.879*(OF3+OF2/0.8998)/OAL
                        + 30.875*OMG/OAL - 5.404*ONA/OAL + 11.12*OKO/OAL - 3.89;
      end;
   end;{if}
   if (CSI*CAL > 0.0) then
        WatM := (CNA+CKO+2.0*CCA)/(CSI*CAL) * 100.0;
   if WatM>1e5 then WatM:=1e5;
end;{proc SundryRatios}


procedure OtherRatios;
begin
  {Chemical index of alteration, calculated in molecular proportions}
  {Nesbitt and Young, Nature, 299, 715-717}
  CIA := 0.0;
  if OXTOT>10 then
  begin
    if ((CAL+CCA-CalciumInNonsilicates+CNA+CKO) > 0.0) then
    begin
      CIA := 100.0*0.5*CAL/(0.5*CAL+CCA-CalciumInNonsilicates+0.5*CNA+0.5*CKO);
    end;
  end;{if}
  {Molar Al/(2Ca+Na+K) peraluminous index -
   Pichavant et al., 1992, Geochimica Cosmochimica Acta, 56, 3855-3861}
  {equivalent to the Aluminium Saturation Index of Zen, 1986, -
   J Petrol, 27, 1095-1117}
  ACNK := 0.0;
  if OXTOT>10 then
  begin
    if ((CCA+CNA+CKO) > 0.0) then
    begin
      ACNK := 0.5*CAL/(CCA+0.5*CNA+0.5*CKO);
    end;
  end;{if}
  {Debon and Lefort variables A and B -
   Debon and Lefort, 1983}
  DebonLefortA := 0.0;
  DebonLefortB := 0.0;
  if OXTOT>10 then
  begin
    DebonLefortA := (mcAL-(mcNA+mcK+2.0*mcCA));
    DebonLefortB := ((mcFE3+mcFE2)+mcMG+mcTI);
  end;{if}
end;{proc OtherRatios}


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
   KO:=OKO/47.0975;    {was 45.1017}
   PO:=OPO/70.9723;
   SO:=OSO/80.0622;
   SU:=OSU/32.064;
   CL:=OCL/35.453;
   FU:=OFU/18.9984;
   CD:=OCD/44.00995;
   R1 := (4.0*SI-11.0*(NA+KO)-2.0*(F2+F3+TI))*1000.0;
   R2 := (AL+2.0*MG+6.0*CA)*1000.0;
   mcSI := SI*1000.0;
   mcTI := TI*1000.0;
   mcAL := AL*1000.0;
   mcFE3 := F3*1000.0;
   mcFE2 := F2*1000.0;
   mcMN := MN*1000.0;
   mcMG := MG*1000.0;
   mcCA := CA*1000.0;
   mcNA := NA*1000.0;
   mcK := KO*1000.0;
end;{proc CatEquiv}


procedure Agpaitic;
{
Compute agpaitic ratio, etc
}
var
   FEMN          : double;
begin
   if AL>0.0 then WAGRAT:=(NA*22.9898+KO*39.102)/(AL*0.269815);
   FEMN:=(F2+F3)*0.55847+MN*0.54938;
   if ((FEMN>0.0) or (MG>0.0)) then WFMRAT:=100.0*FEMN/(FEMN+MG*0.24312);
   ALK:=(NA+KO)*0.5;
   F23:=F2+F3*0.5;
   SUM:=(ALK+F23+MG);
   if SUM>0.0 then begin
     SUM:=100.0/SUM;
     AFMA:=ALK*SUM;
     AFMF:=F23*SUM;
     AFMM:=MG*SUM;
   end;
   if F23>0.0 then OXRAT:=F2*100.0/F23 else OXRAT:=0.0;
   if ALK>0 then ALRAT:=NA*50.0/ALK else ALRAT:=0.0;
   if AL>0.0 then AGRAT:=ALK*200.0/AL;
   FEMN:=F2+F3+MN;
   if ((FEMN+MG)>0.0) then FMRAT:=FEMN*100.0/(FEMN+MG);
end;{proc Agpaitic}


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
   if (dmNrm.cdsNormsCat.RecordCount > 0)
     then dmNrm.cdsNormsCat.Edit
     else dmNrm.cdsNormsCat.Append;
   dmNrm.cdsNormsCatSI.AsFloat := CSI;
   dmNrm.cdsNormsCatTI.AsFloat := CTI;
   dmNrm.cdsNormsCatZR.AsFloat := CZR;
   dmNrm.cdsNormsCatAL.AsFloat := CAL;
   dmNrm.cdsNormsCatCR.AsFloat := CCR;
   dmNrm.cdsNormsCatFE3.AsFloat := CF3;
   dmNrm.cdsNormsCatFE2.AsFloat := CF2;
   dmNrm.cdsNormsCatMN.AsFloat := CMN;
   dmNrm.cdsNormsCatNI.AsFloat := CNI;
   dmNrm.cdsNormsCatMG.AsFloat := CMG;
   dmNrm.cdsNormsCatCA.AsFloat := CCA;
   dmNrm.cdsNormsCatSR.AsFloat := CSR;
   dmNrm.cdsNormsCatBA.AsFloat := CBA;
   dmNrm.cdsNormsCatNA.AsFloat := CNA;
   dmNrm.cdsNormsCatK.AsFloat := CKO;
   dmNrm.cdsNormsCatP.AsFloat := CPO;
   dmNrm.cdsNormsCatSO.AsFloat := CSO;
   dmNrm.cdsNormsCatS.AsFloat := CSU;
   dmNrm.cdsNormsCatCL.AsFloat := CCL;
   dmNrm.cdsNormsCatF.AsFloat := CFU;
   dmNrm.cdsNormsCatC.AsFloat := CCD;
   dmNrm.cdsNormsCatTotal.AsFloat := CTTOT;
   dmNrm.cdsNormsCat.Post;
end;{proc CatEquivPerc}

procedure SetMinZero;
{------------------------------------------------------------
 Set minerals to 0
 ------------------------------------------------------------}
begin
   QZ:=0.0;
   CO:=0.0;
   ZN:=0.0;
   ORT:=0.0;
   AB:=0.0;
   AN:=0.0;
   LC:=0.0;
   NE:=0.0;
   KP:=0.0;
   HL:=0.0;
   TH:=0.0;
   AC:=0.0;
   NS:=0.0;
   KS:=0.0;
   WO:=0.0;
   DI:=0.0;
   HY:=0.0;
   OL:=0.0;
   CS:=0.0;
   MT:=0.0;
   CM:=0.0;
   IL:=0.0;
   HM:=0.0;
   SP:=0.0;
   PF:=0.0;
   RU:=0.0;
   RU:=0.0;
   AP:=0.0;
   FL:=0.0;
   PY:=0.0;
   CC:=0.0;
   PMG:=0.0;
   PF2:=0.0;
   HO:=0.0;
   ACT:=0.0;
   RI:=0.0;
   ED:=0.0;
   BI:=0.0;
   SPIN:=0.0;
   TOTAL:=0.0;

   WQZ:=0.0;
   WCO:=0.0;
   WZN:=0.0;
   WOR:=0.0;
   WAB:=0.0;
   WAN:=0.0;
   WLC:=0.0;
   WNE:=0.0;
   WKP:=0.0;
   WHL:=0.0;
   WTH:=0.0;
   WAC:=0.0;
   WNS:=0.0;
   WKS:=0.0;
   WWO:=0.0;
   WDI:=0.0;
   WHY:=0.0;
   WOL:=0.0;
   WCS:=0.0;
   WMT:=0.0;
   WCM:=0.0;
   WIL:=0.0;
   WHM:=0.0;
   WSP:=0.0;
   WPF:=0.0;
   WRU:=0.0;
   WAP:=0.0;
   WFL:=0.0;
   WPY:=0.0;
   WCC:=0.0;
   WHO:=0.0;
   WBI:=0.0;
   WACT:=0.0;
   WED:=0.0;
   WRI:=0.0;
   WSPIN:=0.0;
   WTOTAL:=0.0;
   CIA := 0.0;
end;{proc SetMinZero}


procedure CatToWt;
{
Recalc. cation equiv. norm in wt. percent
}
begin
   F:=CTL*0.01;
   WQZ:=QZ*60.0848*F;
   WCO:=CO*50.9806*F;
   WZN:=ZN*91.6518*F;
   WOR:=ORT*55.6673*F;
   WAB:=AB*52.4449*F;
   WAN:=AN*55.642*F;
   WPL:=WAB+WAN;
   WLC:=LC*54.563*F;
   WNE:=NE*47.3516*F;
   WKP:=KP*52.7224*F;
   WHL:=HL*58.4428*F;
   WTH:=TH*47.3471*F;
   WAC:=AC*57.7513*F;
   WNS:=NS*40.687*F;
   WKS:=KS*51.4294*F;
   WWO:=WO*58.0821*F;
   WWODI:=DI*29.0411*F;
   WENDI:=DI*PMG*25.0991*F;
   WFSDI:=DI*PF2*32.9828*F;
   WDI:=WWODI+WENDI+WFSDI;
   WEN:=HY*PMG*50.1981*F;
   WFS:=HY*PF2*65.9656*F;
   WHY:=WEN+WFS;
   WFO:=OL*PMG*46.9025*F;
   WFA:=OL*PF2*67.92591*F;
   WOL:=WFO+WFA;
   WBI:=(BI*PMG*49.90886+BI*PF2*61.73449)*F;
   WACT:=(ACT*PMG*52.9596+ACT*PF2*63.47128)*F;
   WED:=(ED*PMG*51.01747+ED*PF2*60.8722)*F;
   WRI:=RI*61.19259*F;
   WSPIN:=SPIN*47.4242*F;
   WHO:=WACT+WED+WRI;
   WCS:=CS*57.4151*F;
   WMT:=MT*77.1795*F;
   WCM:=CM*74.6122*F;
   WIL:=IL*75.8726*F;
   WHM:=HM*79.8461*F;
   WSP:=SP*65.3543*F;
   WPF:=PF*67.9891*F;
   WRU:=RU*79.89881*F;
   if (OFU > 0.0) then WAP:=AP*63.0391*F
                  else WAP:=AP*61.6642*F;
   WFL:=FL*78.0768*F;
   WPY:=PY*119.975*F;
   WCC:=CC*50.044675*F;
   WODI:=0.5*DI;
   ENDI:=0.5*DI*PMG;
   FSDI:=WODI-ENDI;
   FO:=OL*PMG;
   FA:=OL-FO;
   EN:=HY*PMG;
   FS:=HY-EN;
   NPL:=AB+AN;
end;{proc CatToWt}


procedure MinComp;
begin
   ANPL:=0.0;
   FAOL:=0.0;
   ENHY:=0.0;
   WANPL:=0.0;
   WFAOL:=0.0;
   WENHY:=0.0;
   { An content of plag. }
   if NPL>0.0 then
   begin
      ANPL:=100.0*AN/NPL;
      WANPL:=100.0*WAN/WPL;
   end;{if}
   { Fa content of oliv. }
   if OL>0.0 then
   begin
      FAOL:=100.0*PF2;
      WFAOL:=100.0*WFA/WOL;
   end;{if}
   { En content of hypersth. }
   if HY>0.0 then
   begin
      ENHY:=100.0*PMG;
      WENHY:=100.0*WEN/WHY;
   end;{if}
end;{proc MinComp}


procedure DifIndex;
{
Thornton-Tuttle differ. index
}
begin
   DIX:=QZ+ORT+AB+NE+LC+KP;
   WDIX:=WQZ+WOR+WAB+WNE+WLC+WKP;
end;{proc DifIndex}


procedure Totals;
{
Totals for salic, femic and all minerals
}
begin
   SALIC:=DIX+CO+ZN+AN+HL+TH;
   FEMIC:=AC+NS+KS+WO+DI+HY+OL+BI+HO+SPIN+CS+MT+CM;
   FEMIC:=FEMIC+IL+HM+SP+PF+RU+AP+FL+PY+CC;
   TOTAL:=FEMIC+SALIC;
   WSALIC:=WDIX+WCO+WZN+WAN+WHL+WTH;
   WFEMIC:=WAC+WNS+WKS+WWO+WDI+WHY+WOL+WBI+WHO+WSPIN+WCS+WMT+WCM;
   WFEMIC:=WFEMIC+WIL+WHM+WSP+WPF+WRU+WAP+WFL+WPY+WCC;
   WTOTAL:=WFEMIC+WSALIC;
end;{proc Totals}


procedure CatToWt100;
{
Recalc. wt. % norm ito 100 wt. percent
}
begin
   F:=100.0/WTOTAL;
   WQZ:=WQZ*F;
   WCO:=WCO*F;
   WZN:=WZN*F;
   WOR:=WOR*F;
   WAB:=WAB*F;
   WAN:=WAN*F;
   WPL:=WAB+WAN;
   WLC:=WLC*F;
   WNE:=WNE*F;
   WKP:=WKP*F;
   WHL:=WHL*F;
   WTH:=WTH*F;
   WAC:=WAC*F;
   WNS:=WNS*F;
   WKS:=WKS*F;
   WWO:=WWO*F;
   WWODI:=WWODI*F;
   WENDI:=WENDI*F;
   WFSDI:=WFSDI*F;
   WDI:=WWODI+WENDI+WFSDI;
   WEN:=WEN*F;
   WFS:=WFS*F;
   WHY:=WEN+WFS;
   WFO:=WFO*F;
   WFA:=WFA*F;
   WOL:=WFO+WFA;
   WBI:=WBI*F;
   WACT:=WACT*F;
   WED:=WED*F;
   WRI:=WRI*F;
   WSPIN:=WSPIN*F;
   WHO:=WACT+WED+WRI;
   WCS:=WCS*F;
   WMT:=WMT*F;
   WCM:=WCM*F;
   WIL:=WIL*F;
   WHM:=WHM*F;
   WSP:=WSP*F;
   WPF:=WPF*F;
   WRU:=WRU*F;
   WAP:=WAP*F;
   WFL:=WFL*F;
   WPY:=WPY*F;
   WCC:=WCC*F;
   MinComp;
   DifIndex;
   WSALIC:=WDIX+WCO+WZN+WAN+WHL+WTH;
   WFEMIC:=WAC+WNS+WKS+WWO+WDI+WHY+WOL+WBI+WHO+WSPIN+WCS+WMT+WCM;
   WFEMIC:=WFEMIC+WIL+WHM+WSP+WPF+WRU+WAP+WFL+WPY+WCC;
   WTOTAL:=WFEMIC+WSALIC;
end;{proc CatToWt100}


procedure TfmNormMain.CalcCIPW;
begin
   Item:=0;
   Names;
   dmNrm.cdsNormChem.First;
   repeat
      SetMinZero;      {set minerals to 0}
      DES:=dmNrm.cdsNormChemSAMPLENUM.AsString;
      Item:=Item+1;
      sbMain.SimpleText := 'CIPW-norm :   Item '+IntToStr(Item)+'   Sample '
                      +dmNrm.cdsNormChemSampleNum.AsString+'  -  '+DES;
      Record_To_Memory;
      CatEquiv;
      Agpaitic;
      CatEquivPerc;
      SundryRatios;
      if OXTOT>10.0 then
      begin
         CalciumInNonsilicates := 0.0;
         Apatite;
         Halite;
         Thenardite;
         Pyrite;
         Chromite;
         Ilmenite;
         Fluorite;
         Calcite;
         Zircon;
         Orthoclase;  {Orthoclase and Potassium metasilicate}
         Albite;
         AnCor;       {Anorthite and Corundum}
         Acmite;      {Acmite and Sodium metasilicate}
         SphRut;      {Sphene and Rutile}
         MtHm;        {Magnetite and Hematite}
         F2MgRatio;
         WollHyp;     {Wollastonite and Hypersthene}
         if HY>0.0 then Diopside;
         if SI < 0.0 then DesilHy;    {Test for Si deficiency}
         if SI < 0.0 then DesilSp;    {Test for Si deficiency}
         if SI < 0.0 then DesilAb;    {Test for Si deficiency}
         if SI < 0.0 then DesilOr;    {Test for Si deficiency}
         if SI < 0.0 then DesilWo;    {Test for Si deficiency}
         if SI < 0.0 then DesilDi;    {Test for Si deficiency}
         if SI < 0.0 then DesilLc;    {Test for Si deficiency}
         MakeQz;
         CatToWt;
         MinComp;
         DifIndex;
         Totals;
         if Toggle100=1 then
         begin
           CatToWt100;
         end;
         OrAbAn;
         QzNeKp;
         QzAbOr;
         OtherRatios;
         if OutRoute='F' then
         begin
           dmNrm.cdsNormsMin.Append;
           dmNrm.cdsNormsMinNORMTYPE.AsString := 'CIPW';
           //dmNrm.cdsNormsMinGROUPNAME.AsString := dmNrm.cdsNormChemGROUPNAME.AsString;
           dmNrm.cdsNormsMinSAMPLENUM.AsString := dmNrm.cdsNormChemSAMPLENUM.AsString;
           Memory_To_Record;
           dmNrm.cdsNormsMin.Post;
         end;
         if cbPrint.Checked then
         begin
           CIPWPrint;
         end;
      end;{if OXTOT}
     dmNrm.cdsNormChem.Next;
   until dmNrm.cdsNormChem.EOF;{until}
end;{proc CalcCIPW}


procedure TfmNormMain.CalcMeso;
begin
   Item:=0;
   Names;
   dmNrm.cdsNormChem.First;
   repeat
      SetMinZero;      {set minerals to 0}
      DES:=dmNrm.cdsNormChemSAMPLENUM.AsString;
      Item:=Item+1;
      sbMain.SimpleText := 'Meso-norm :   Item '+IntToStr(Item)+'   Sample '+dmNrm.cdsNormChemSampleNum.AsString;
      Record_To_Memory;
      CatEquiv;
      Agpaitic;
      CatEquivPerc;
      SundryRatios;
      if OXTOT>10.0 then begin
         CalciumInNonsilicates := 0.0;
         Calcite;
         Apatite;
         Fluorite;
         Halite;
         Thenardite;
         Pyrite;
         Chromite;
         Zircon;
         SphRut;      {Sphene and Rutile}
         Orthoclase;  {Orthoclase and Potassium metasilicate}
         Albite;
         AnCor;       {Anorthite and Corundum}
         Riebeckite;  {Riebeckite and Sodium metasilicate}
         MtHm;        {Magnetite and Hematite}
         F2MgRatio;
         Biotite;
         if InOption='2' then
         begin
            ActHyWo;
         end
         else begin
            WollHyp;
            Diopside;
         end;{if}
         if SI < 0.0 then DesilAct;   {Test for Si deficiency}
         if SI < 0.0 then DesilHy;    {Test for Si deficiency}
         if SI < 0.0 then DesilOl;    {Test for Si deficiency}
         if SI < 0.0 then DesilAb;    {Test for Si deficiency}
         MakeQz;
         Hornblende;
         CatToWt;
         MinComp;
         DifIndex;
         Totals;
         if Toggle100=1 then
         begin
           CatToWt100;
         end;
         OrAbAn;
         QzNeKp;
         QzAbOr;
         OtherRatios;
         if OutRoute='F' then
         begin
           dmNrm.cdsNormsMin.Append;
           if (InOption = '1') then
             dmNrm.cdsNormsMinNORMTYPE.AsString := 'Meso-px'
           else
             dmNrm.cdsNormsMinNORMTYPE.AsString := 'Meso-hb';
           //dmNrm.cdsNormsMinGROUPNAME.AsString := dmNrm.cdsNormChemGROUPNAME.AsString;
           dmNrm.cdsNormsMinSAMPLENUM.AsString := dmNrm.cdsNormChemSAMPLENUM.AsString;
           Memory_To_Record;
           dmNrm.cdsNormsMin.Post;
         end;
         if cbPrint.Checked then
         begin
           MesoPrint;
         end;
      end;{if OXTOT}
     dmNrm.cdsNormChem.Next;
   until dmNrm.cdsNormChem.EOF;{until}
end;{proc CalcMeso}

procedure TfmNormMain.bbCalculateClick(Sender: TObject);
begin
   Toggle100:=0;
   if cbRecalculate100.Checked then Toggle100 := 1
                               else Toggle100 := 0;
   DES:=' ';
   A14:='          Species     Input     Calc     Comp     Mineral    Wt       Cat     ';
   A15:='                       Wt        Cat     Error              Prcnt    Prcnt    ';
   A16:='                      Prcnt     Prcnt    Ctpcnt      ';
   done:=False;
   OutRoute:='F';
   Nbeg:=1;
   Nfin:=dmNrm.cdsNormChem.RecordCount;
   InOption := 'Q';
   if rbCIPW.Checked then InOption := '0';
   if rbMesoPx.Checked then InOption := '1';
   if rbMesoHb.Checked then InOption := '2';
   case InOption of
     '0' : begin
              CalcCIPW;
     end;
     '1','2' : begin
              CalcMeso;
     end;
   end;{case}
   dmNrm.cdsNormsMin.First;
   dmNrm.cdsNormChem.First;
   sbMain.SimpleText := '';
end;


procedure TfmNormMain.Exit1Click(Sender: TObject);
begin
  SetIniFile;
  Close;
end;

procedure TfmNormMain.bbEmptyNormsMinClick(Sender: TObject);
begin
  dmNrm.cdsNormsMin.MasterSource := nil;
  dmNrm.cdsNormsMin.MasterFields := '';
  if (dmNrm.cdsNormsMin.RecordCount > 0) then
  begin
    dmNrm.cdsNormsMin.Last;
    repeat
      dmNrm.cdsNormsMin.Delete;
      dmNrm.cdsNormsMin.Next;
    until dmNrm.cdsNormsMin.BOF;
  end;
end;

procedure TfmNormMain.Export1Click(Sender: TObject);
var
  j : integer;
begin
  ExportNormativeMinerals;
end;

procedure TfmNormMain.ExportNormativeMinerals;
var
  fr: TFlexCelReport;
  frTemplateStr, frFileNameStr : string;
begin
  frTemplateStr := FlexTemplatePath+'nrm_normresults.xlsx';
  SaveDialogSprdSheet.InitialDir := ExportPath;
  SaveDialogSprdSheet.FileName := 'NORM_results_minerals';
  frFileNameStr := SaveDialogSprdSheet.FileName;
  if SaveDialogSprdSheet.Execute then
  begin
    frFileNameStr := SaveDialogSprdSheet.FileName;
    ExportPath := ExtractFilePath(SaveDialogSprdSheet.FileName);
    fr := TFlexCelReport.Create(true);
    try
      fr.AddTable('cdsNormsMin',dmNrm.cdsNormsMin);
      fr.Run(frTemplateStr,frFileNameStr);
    finally
      fr.Free;
    end;
  end;
end;

procedure TfmNormMain.Import1Click(Sender: TObject);
begin
  try
    ImportForm := TfmSheetImport.Create(Self);
    ImportForm.OpenDialogSprdSheet.FileName := 'NORMCHEM';
    ImportForm.ShowModal;
  finally
    ImportForm.Free;
  end;
end;

procedure TfmNormMain.ImportTemplate1Click(Sender: TObject);
begin
  try
    ImportFormTemplate := TfmSheetImportTemplate.Create(Self);
    ImportFormTemplate.OpenDialogSprdSheet.FileName := 'NORMFAC';
    ImportFormTemplate.ShowModal;
  finally
    ImportFormTemplate.Free;
  end;
end;

procedure TfmNormMain.bbEmptyNormChemClick(Sender: TObject);
begin
  if (dmNrm.cdsNormChem.RecordCount > 0) then
  begin
    dmNrm.cdsNormChem.Last;
    repeat
      dmNrm.cdsNormChem.Delete;
      dmNrm.cdsNormChem.Next;
    until dmNrm.cdsNormChem.BOF;
  end;
  dmNrm.cdsNormsCat.MasterSource := nil;
  dmNrm.cdsNormsCat.MasterFields := '';
  if (dmNrm.cdsNormsCat.RecordCount > 0) then
  begin
    dmNrm.cdsNormsCat.Last;
    repeat
      dmNrm.cdsNormsCat.Delete;
      dmNrm.cdsNormsCat.Next;
    until dmNrm.cdsNormsCat.BOF;
  end;
  dmNrm.cdsNormsCat.MasterSource := dmNrm.dsNormChem;
  dmNrm.cdsNormsCat.MasterFields := 'SAMPLENUM';
end;

procedure TfmNormMain.About1Click(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TfmNormMain.Printersetup1Click(Sender: TObject);
begin
  PrinterSetupDialog1.Execute;
end;

procedure TfmNormMain.GetIniFile;
var
  AppIni   : TIniFile;
  tmpStr   : string;
  iCode    : integer;
  PublicPath : string;
begin
  //PublicPath := TPath.GetPublicPath;
  //CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  //Used to use CSIDL_COMMON_APPDATA but some users do not have access to this
  //and don't know how to change their system settings and permissions to all
  //software to write to this path.
  //Now changed to use CSIDL_COMMON_DOCUMENTS which automatically permits
  //all users to have both read and write permission
  PublicPath := TPath.GetHomePath;
  CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  IniFilename := CommonFilePath + 'Norm.ini';
  IniFilePath := CommonFilePath;
  ProgramFilePath := IniFilePath + 'Norm\';
  AppIni := TIniFile.Create(IniFilename);
  try
    MineralTablePath := AppIni.ReadString('File Paths','Mineral Results Table',CommonFilePath+'Norm\Data\');
    DataPath := AppIni.ReadString('File Paths','Data spreadsheets',CommonFilePath+'Norm\Data\');
    cdsPath := AppIni.ReadString('File Paths','Internal files',CommonFilePath+'Norm\Data\');
    FlexTemplatePath := AppIni.ReadString('File Paths','Spreadsheet template path',CommonFilePath+'Norm\Templates\');
    ExportPath := AppIni.ReadString('File Paths','Spreadsheet export path',CommonFilePath+'Norm\Exports\');
    GlobalChosenStyle := AppIni.ReadString('Styles','Chosen style','Windows');
    if (GlobalChosenStyle = '') then GlobalChosenStyle := 'Windows';
    dmNrm.ChosenStyle := GlobalChosenStyle;
    PositionColStr := AppIni.ReadString('ColumnDefinitions','PositionColStr','B');
    RequiredColStr := AppIni.ReadString('ColumnDefinitions','RequiredColStr','C');
    EnteredColStr := AppIni.ReadString('ColumnDefinitions','EnteredColStr','C');
    ColumnColStr := AppIni.ReadString('ColumnDefinitions','ColumnColStr','D');
    FactorColStr := AppIni.ReadString('ColumnDefinitions','FactorColStr','A');
    tmpStr := AppIni.ReadString('Defaults','DefaultMinimum','1.0e-3');
    Val(tmpStr,DefaultMinimum,iCode);
    if (iCode > 0) then DefaultMinimum := 1.0e-3;
  finally
    AppIni.Free;
  end;
  MineralTablePath := UpperCase(MineralTablePath);
  DataPath := UpperCase(DataPath);
  cdsPath := UpperCase(cdsPath);
  FlexTemplatePath := UpperCase(FlexTemplatePath);
  ExportPath := UpperCase(ExportPath);
end;

procedure TfmNormMain.SetIniFile;
var
  AppIni   : TIniFile;
  tmpStr   : string;
  iCode    : integer;
  PublicPath : string;
begin
  //PublicPath := TPath.GetPublicPath;
  //CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  //Used to use CSIDL_COMMON_APPDATA but some users do not have access to this
  //and don't know how to change their system settings and permissions to all
  //software to write to this path.
  //Now changed to use CSIDL_COMMON_DOCUMENTS which automatically permits
  //all users to have both read and write permission
  PublicPath := TPath.GetHomePath;
  CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  IniFilename := CommonFilePath + 'Norm.ini';
  IniFilePath := CommonFilePath;
  AppIni := TIniFile.Create(IniFilename);
  try
    AppIni.WriteString('File Paths','Mineral Results Table',MineralTablePath);
    AppIni.WriteString('File Paths','Data spreadsheets',DataPath);
    AppIni.WriteString('File Paths','Internal files',cdsPath);
    AppIni.WriteString('File Paths','Spreadsheet template path',FlexTemplatePath);
    AppIni.WriteString('File Paths','Spreadsheet export path',ExportPath);
    AppIni.WriteString('Styles','Chosen style',GlobalChosenStyle);
    AppIni.WriteString('ColumnDefinitions','PositionColStr',PositionColStr);
    AppIni.WriteString('ColumnDefinitions','RequiredColStr',RequiredColStr);
    AppIni.WriteString('ColumnDefinitions','EnteredColStr',EnteredColStr);
    AppIni.WriteString('ColumnDefinitions','ColumnColStr',ColumnColStr);
    AppIni.WriteString('ColumnDefinitions','FactorColStr',FactorColStr);
    AppIni.WriteString('Defaults','DefaultMinimum',FormatFloat('##0.0000e-00',DefaultMinimum));
  finally
    AppIni.Free;
  end;
  MineralTablePath := UpperCase(MineralTablePath);
  DataPath := UpperCase(DataPath);
  cdsPath := UpperCase(cdsPath);
  FlexTemplatePath := UpperCase(FlexTemplatePath);
  ExportPath := UpperCase(ExportPath);
end;

procedure TfmNormMain.StyleClick(Sender: TObject);
var
  StyleName : String;
  i : integer;
begin
  //get style name
  StyleName := TMenuItem(Sender).Caption;
  StyleName := StringReplace(StyleName, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  GlobalChosenStyle := StyleName;
  dmNrm.ChosenStyle := GlobalChosenStyle;
  //set active style
  Application.ProcessMessages;
  TStyleManager.SetStyle(GlobalChosenStyle);
  dmNrm.ChosenStyle := GlobalChosenStyle;
  Application.ProcessMessages;
  //check the currently selected menu item
  (Sender as TMenuItem).Checked := true;
  //uncheck all other style menu items
  for i := 0 to Styles1.Count-1 do
  begin
    if not Styles1.Items[i].Equals(Sender) then
      Styles1.Items[i].Checked := false;
  end;
  for i := 0 to Styles1.Count-1 do
  begin
    if Styles1.Items[i].Checked then GlobalChosenStyle := StringReplace(Styles1.Items[i].Caption, '&', '',
    [rfReplaceAll,rfIgnoreCase]);
  end;
  TStyleManager.SetStyle(GlobalChosenStyle);
  try
    dmNrm.ChosenStyle := GlobalChosenStyle;
  finally
    dmNrm.ChosenStyle := GlobalChosenStyle;
  end;
end;

procedure TfmNormMain.Test1Click(Sender: TObject);
begin
  ShowMessage(cdsPath);
  ShowMessage(MineralTablePath);
  ShowMessage(DataPath);
  ShowMessage(FlexTemplatePath);
end;

procedure TfmNormMain.FormShow(Sender: TObject);
var
  CanProceed : boolean;
begin
  pc1.ActivePage := tsControl;
  CanProceed := true;
  ShowOnly50Rows := true;
  FromRowValueString := '2';
  ToRowValueString := '3';
  GetIniFile;
  try
    with dmNrm do
    begin
      cdsNormChem.FileName := cdsPath+'\'+'NormChem.xml';
      if not FileExists(cdsNormChem.FileName) then
      begin
        MessageDlg('Required internal file not found - '+cdsPath+'\'+'NormChem.xml',mtWarning,[mbOK],0);
        CanProceed := false;
      end else
      begin
        {
        MessageDlg('Found - '+cdsPath+'\'+'NormChem.xml',mtWarning,[mbOK],0);
        }
      end;
      cdsNormsFac.FileName := cdsPath+'\'+'NormsFac.xml';
      if not FileExists(cdsNormsFac.FileName) then
      begin
        MessageDlg('Required internal file not found - '+cdsPath+'\'+'NormsFac.xml',mtWarning,[mbOK],0);
        CanProceed := false;
      end else
      begin
        {
        MessageDlg('Found - '+cdsPath+'\'+'NormsFac.xml',mtWarning,[mbOK],0);
        }
      end;
      if CanProceed then
      begin
        cdsNormChem.LoadFromFile(cdsPath+'\'+'NormChem.xml');
        cdsNormsFac.LoadFromFile(cdsPath+'\'+'NormsFac.xml');
        cdsNormChem.Open;
        cdsNormsCat.Open;
        cdsNormsFac.Open;
        cdsNormsMin.Open;
        cdsNormsMinLinked.CloneCursor(cdsNormsMin,false);
        cdsNormsMinLinked.AddIndex('NORMTYPE','NORMTYPE',[]);
        cdsNormsMinLinked.IndexName := 'NORMTYPE';
        //cdsNormsMinLinked.AddIndex('GROUPNAME','GROUPNAME',[]);
        //cdsNormsMinLinked.IndexName := 'GROUPNAME';
        cdsNormsMinLinked.AddIndex('SAMPLENUM','SAMPLENUM',[]);
        cdsNormsMinLinked.IndexName := 'SAMPLENUM';
        //cdsNormsMinLinked.IndexFieldNames := 'GROUPNAME;SAMPLENUM';
        cdsNormsMinLinked.IndexFieldNames := 'SAMPLENUM';
        cdsNormsMinLinked.MasterSource := dsNormChem;
        //cdsNormsMinLinked.MasterFields := 'GROUPNAME;SAMPLENUM';
        cdsNormsMinLinked.MasterFields := 'SAMPLENUM';
        cdsNormsMinLinked.Open;
      end else
      begin
        Close;
      end;
    end;
    dsNormsMin.DataSet := dmNrm.cdsNormsMinLinked;
  except
  end;
end;

procedure TfmNormMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SetIniFile;
  with dmNrm do
  begin
    cdsNormsFac.SaveToFile(cdsPath+'\'+'NormsFac.xml',dfXML);
    cdsNormChem.SaveToFile(cdsPath+'\'+'NormChem.xml',dfXML);
  end;
end;

procedure TfmNormMain.FormCreate(Sender: TObject);
var
  Style: String;
  Item: TMenuItem;
begin
  //Add child menu items based on available styles.
  GetIniFile;
  TStyleManager.TrySetStyle(GlobalChosenStyle);
  for Style in TStyleManager.StyleNames do
  begin
    Item := TMenuItem.Create(Styles1);
    Item.Caption := Style;
    Item.OnClick := StyleClick;
    if TStyleManager.ActiveStyle.Name = Style then
      Item.Checked := true;
    Styles1.Add(Item);
  end;
end;

procedure TfmNormMain.Mineraltablepath1Click(Sender: TObject);
begin
  MineralTablePath := InputBox('Path definition','Mineral table export path is ',MineralTablePath);
  MineralTablePath := UpperCase(MineralTablePath);
end;

end.
