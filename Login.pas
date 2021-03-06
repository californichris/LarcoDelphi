unit Login;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls;

type
  TfrmLogin = class(TForm)
    cmdOk: TButton;
    cmdCancel: TButton;
    Panel1: TPanel;
    Label2: TLabel;
    txtPassword: TEdit;
    txtUser: TEdit;
    Label1: TLabel;
    lblValidate: TLabel;
    procedure cmdCancelClick(Sender: TObject);
    procedure cmdOkClick(Sender: TObject);
    function IsValidUser(User: String;Password: String):boolean;
    procedure txtUserKeyPress(Sender: TObject; var Key: Char);
    procedure txtPasswordKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }

  end;

var
  frmLogin: TfrmLogin;
  gsConnString,StartDDir: String;

implementation

uses Main, Scrap;

{$R *.dfm}

procedure TfrmLogin.cmdCancelClick(Sender: TObject);
begin
Close;
if lblValidate.Caption <> 'true' then
  frmMain.Close;
end;

procedure TfrmLogin.cmdOkClick(Sender: TObject);
begin
   if lblValidate.Caption = 'true' then begin
     if  IsValidUser(txtUser.Text,txtPassword.Text ) then
         begin
              ModalResult := mrOK;
         end
     else
         Begin
              ShowMessage('No tienes permisos para scrapear.');
              txtPassword.Text := '';
              txtPassword.SetFocus;
         end;
   end
   else begin
     if  IsValidUser(txtUser.Text,txtPassword.Text ) then
         begin
              frmMain.Enabled := True;
              frmMain.sUserLogin := txtUser.Text;
              frmMain.sUserPassword := txtPassword.Text;
              frmMain.StatusBar.Panels[2].Text := txtUser.Text;
              frmLogin.Hide;
         end
     else
         Begin
              ShowMessage('Usuario o Password Invalido');
              txtPassword.Text := '';
              txtPassword.SetFocus;
         end;
   end;


end;

function TfrmLogin.IsValidUser(User: String;Password: String):boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := False;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT USE_ID FROM tblUsers WHERE USE_Login = ' + QuotedStr(User) +
                ' AND USE_Password = ' + QuotedStr(Password);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      if Qry.RecordCount > 0 then begin
          if lblValidate.Caption <> 'true' then begin
             frmMain.sUserID := Qry['USE_ID'];
             frmMain.BindMenu(Qry['USE_ID']);
          end;
          Result := True;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure TfrmLogin.txtUserKeyPress(Sender: TObject; var Key: Char);
begin
     if (key = chr(vk_return)) or (key = chr(vk_tab)) then
        txtPassword.SetFocus;
end;

procedure TfrmLogin.txtPasswordKeyPress(Sender: TObject; var Key: Char);
begin
     if (key = chr(vk_return)) or (key = chr(vk_tab)) then
        cmdOkClick(nil);
end;

procedure TfrmLogin.FormShow(Sender: TObject);
begin
  frmMain.Enabled := False;
end;

end.
