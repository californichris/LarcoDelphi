unit InventariosConfiguracion;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math;

type
  TfrmInventariosConf = class(TForm)
    ddlOpcion: TComboBox;
    Button1: TButton;
    Button2: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInventariosConf: TfrmInventariosConf;

implementation

uses Main;

{$R *.dfm}

procedure TfrmInventariosConf.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmInventariosConf.Button2Click(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmInventariosConf.FormCreate(Sender: TObject);
var Qry : TADOQuery;
Conn : TADOConnection;
begin

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT TOP 1 CONF_Inventarios FROM tblInventariosConf';
  Qry.Open;

  ddlOpcion.ItemIndex := ddlOpcion.Items.IndexOf(VarToStr(Qry['CONF_Inventarios']))
end;

procedure TfrmInventariosConf.Button1Click(Sender: TObject);
var Conn : TADOConnection;
SQLStr : String;
begin

  SQLStr := 'UPDATE tblInventariosConf SET CONF_Inventarios = ' + QuotedStr(ddlOpcion.Text);

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Conn.Execute(SQLStr);

  ShowMessage('La configuracion de cambio exitosamente.');
end;

end.
