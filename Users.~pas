unit Users;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,ComObj,Larco_Functions;


type
  TfrmUsers = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    gbButtons: TGroupBox;
    Label8: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    lblId: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    cmbEmpleados: TComboBox;
    txtLogin: TEdit;
    txtPassword: TEdit;
    GroupBox2: TGroupBox;
    GridView2: TGridView;
    btnDelete: TButton;
    btnAdd: TButton;
    GridView1: TGridView;
    Label3: TLabel;
    lblUsuario: TLabel;
    procedure FormCreate(Sender: TObject);
    Procedure BindUsers();
    Procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BorrarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure BindEmpleados();
    procedure cmbEmpleadosChange(Sender: TObject);
    procedure BindGrupos();
    procedure BindPermisos();
    function ValidateData(userExists: Boolean):Boolean;
    procedure btnAddClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUsers: TfrmUsers;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmUsers.FormCreate(Sender: TObject);
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT * FROM tblUsers ORDER BY USE_Login';
    Qry.Open;

    BindEmpleados;
    if Qry.RecordCount > 0 then
        BindUsers();

  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmUsers.BindUsers();
begin
        lblId.Caption := VarToStr(Qry['USE_ID']);
        cmbEmpleados.Text := VarToStr(Qry['USE_Name']);
        txtLogin.Text := VarToStr(Qry['USE_Login']);
        txtPassword.Text := VarToStr(Qry['USE_Password']);
        lblUsuario.Caption := cmbEmpleados.Text;
        BindGrupos();
        BindPermisos();
end;

procedure TfrmUsers.ClearData();
begin
        cmbEmpleados.Text := '';
        txtLogin.Text := '';
        txtPassword.Text := '';
end;

procedure TfrmUsers.EnableControls(Value:Boolean);
begin
    //txtLogin.ReadOnly := Value;
    txtPassword.ReadOnly := Value;
    cmbEmpleados.Enabled := not Value;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;

procedure TfrmUsers.Button2Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Prior;
BindUsers;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmUsers.Button3Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Next;
BindUsers;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmUsers.Button4Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Last;
BindUsers;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;


procedure TfrmUsers.Button1Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.First;
BindUsers;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmUsers.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmUsers.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtLogin.SetFocus;
giOpcion := 1;

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmUsers.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
txtLogin.SetFocus;
end;

procedure TfrmUsers.BuscarClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtLogin.ReadOnly := False;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;
cmbEmpleados.SetFocus;
giOpcion := 4;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
end;

procedure TfrmUsers.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmUsers.btnCancelarClick(Sender: TObject);
begin
ClearData();
EnableControls(True);

Nuevo.Enabled := True;
if Qry.RecordCount > 0 Then
begin
      Editar.Enabled := True;
      Borrar.Enabled := True;
      Buscar.Enabled := True;
end;
BindUsers();
giOpcion := 0;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmUsers.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if ValidateData(True) = False then
          begin
                  Exit;
          end;

        Qry.Insert;
        Qry['use_login'] := txtLogin.Text;
        Qry['USE_Password'] := txtPassword.Text;
        Qry['USE_Name'] := cmbEmpleados.Text;
        Qry['USE_Role'] := 'admin';
        Qry.Post;
  end
  else if giOpcion = 2 then
  begin
        if ValidateData(False) = False then
          begin
                  Exit;
          end;

        Qry.Edit;
        Qry['use_login'] := txtLogin.Text;
        Qry['USE_Password'] := txtPassword.Text;
        Qry['USE_Name'] := cmbEmpleados.Text;
        Qry['USE_Role'] := 'admin';
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar el usuario : ' +
                      cmbEmpleados.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar el usuario : ' +
                            cmbEmpleados.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                      Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if not Qry.Locate('USE_Login',txtLogin.Text ,[loPartialKey] ) then
          begin
              MessageDlg('No se encontro ningun Usuario con estos datos.', mtInformation,[mbOk], 0);
              txtLogin.SetFocus;
              Exit;
          end;

          //Exit;
  end;

EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;
Nuevo.Enabled := True;
if Qry.RecordCount > 0 Then
begin
      Editar.Enabled := True;
      Borrar.Enabled := True;
      Buscar.Enabled := True;

end;
BindUsers;

giOpcion := 0;
  EnableFormButtons(gbButtons, sPermits);
end;

function  TfrmUsers.ValidateData(userExists: Boolean):Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
  Result := True;

  if txtLogin.Text = '' then begin
            ShowMessage('El Login no puede estar vacio.');
            Result := False;
            Exit;
  end;

  if txtPassword.Text = '' then begin
      ShowMessage('El password no puede estar vacio.');
      Result := False;
      Exit;
  end;

  if userExists = True then begin
      Qry2 := TADOQuery.Create(nil);
      Qry2.Connection :=Conn;

      SQLStr := 'SELECT * FROM tblUsers WHERE USE_Login = ' + QuotedStr(txtLogin.Text);

      Qry2.SQL.Clear;
      Qry2.SQL.Text := SQLStr;
      Qry2.Open;

      if Qry2.RecordCount > 0 then begin
              ShowMessage('El Usuario: ' + txtLogin.Text + ' ya existe.');
              Result := False;
              Exit;
      end;
  end

end;


procedure TfrmUsers.BindEmpleados();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblEmpleados Order By Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbEmpleados.Items.Clear;
    While not Qry2.Eof do
    Begin
        cmbEmpleados.Items.Add(Qry2['Nombre']);
        Qry2.Next;
    End;

    cmbEmpleados.Text := '';
    Qry2.Close;
end;


procedure TfrmUsers.cmbEmpleadosChange(Sender: TObject);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Id FROM tblempleados WHERE Nombre = ' + QuotedStr(cmbEmpleados.Text);

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then
        txtLogin.Text := Qry2['Id'];

    Qry2.Close;

end;

procedure TfrmUsers.BindGrupos();
var SQLStr : String;
Qry2 : TADOQuery;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblGroups ' +
                  'WHERE Group_ID NOT IN ( SELECT Group_ID FROM tblUser_Groups WHERE USE_ID = ' +
                  lblId.Caption + ') ' +
                  'ORDER BY Group_ID';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        GridView1.ClearRows;
        While not Qry2.Eof do
        begin
            GridView1.AddRow(1);
            GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry2['Group_ID']);
            GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry2['Group_Name']);
            Qry2.Next;
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry2.Close;
end;

procedure TfrmUsers.BindPermisos();
var SQLStr : String;
Qry2 : TADOQuery;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT U.USE_ID,U.Group_ID,G.Group_Name FROM tblUser_Groups U ' +
                  'INNER JOIN tblGroups G ON  U.Group_ID = G.Group_ID ' +
                  'WHERE USE_ID = ' + lblId.Caption + ' ORDER BY U.Group_ID';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        GridView2.ClearRows;
        While not Qry2.Eof do
        begin
            GridView2.AddRow(1);
            GridView2.Cells[0,GridView2.RowCount -1] := VarToStr(Qry2['USE_ID']);
            GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry2['Group_ID']);
            GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry2['Group_Name']);
            Qry2.Next;
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry2.Close;
end;


procedure TfrmUsers.btnAddClick(Sender: TObject);
var SQLStr : String;
begin
    try
        SQLStr := 'INSERT INTO tblUser_Groups(USE_ID,Group_ID) VALUES(' +
                  lblId.Caption + ',' + GridView1.Cells[0,GridView1.SelectedRow] + ')';

        conn.Execute(SQLStr);

        BindGrupos();
        BindPermisos();
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

end;

procedure TfrmUsers.btnDeleteClick(Sender: TObject);
var SQLStr : String;
begin
    try
        SQLStr := 'DELETE FROM tblUser_Groups WHERE USE_ID = ' +
                  GridView2.Cells[0,GridView2.SelectedRow] + ' AND Group_ID = ' +
                  GridView2.Cells[1,GridView2.SelectedRow];

        conn.Execute(SQLStr);

        BindGrupos();
        BindPermisos();
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

end;

end.
