unit Clientes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,Larco_Functions;

type
  TfrmClientes = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtClave: TEdit;
    txtNombre: TEdit;
    txtCiudad: TEdit;
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
    lblId: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    txtContacto: TEdit;
    txtCalle: TEdit;
    txtEstado: TEdit;
    Label7: TLabel;
    txtNumero: TEdit;
    Label8: TLabel;
    txtColonia: TEdit;
    Label9: TLabel;
    txtCP: TEdit;
    Label11: TLabel;
    txtTelefono: TEdit;
    Label12: TLabel;
    txtCelular: TEdit;
    Label13: TLabel;
    txtFax: TEdit;
    Label10: TLabel;
    txtRFC: TEdit;
    Label14: TLabel;
    function ValidateData():Boolean;
    Procedure EnableControls(Value:Boolean);
    Procedure BindCliente();
    Procedure ClearData();
    procedure SendTab(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure txtClaveKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmClientes: TfrmClientes;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmClientes.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;
end;


procedure TfrmClientes.Button1Click(Sender: TObject);
begin
Qry.First;
BindCliente;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmClientes.Button2Click(Sender: TObject);
begin
Qry.Prior;
BindCliente;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmClientes.Button3Click(Sender: TObject);
begin
Qry.Next;
BindCliente;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmClientes.Button4Click(Sender: TObject);
begin
Qry.Last;
BindCliente;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmClientes.FormCreate(Sender: TObject);
begin
    sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
    EnableFormButtons(gbButtons, sPermits);


    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT * FROM tblClientes ORDER BY Id';
    Qry.Open;

    if Qry.RecordCount > 0 then
        BindCliente();
end;

procedure TfrmClientes.BindCliente();
begin
        lblId.Caption  := VarToStr(Qry['Id']);
        txtClave.Text := VarToStr(Qry['Clave']);
        txtNombre.Text := VarToStr(Qry['Nombre']);
        txtContacto.Text := VarToStr(Qry['Contacto']);
        txtRFC.Text := VarToStr(Qry['RFC']);
        txtCalle.Text := VarToStr(Qry['Calle']);
        txtNumero.Text := VarToStr(Qry['Numero']);
        txtColonia.Text := VarToStr(Qry['Colonia']);
        txtCP.Text := VarToStr(Qry['CP']);
        txtCiudad.Text := VarToStr(Qry['Ciudad']);
        txtEstado.Text := VarToStr(Qry['Estado']);
        txtTelefono.Text := VarToStr(Qry['Telefono']);
        txtCelular.Text := VarToStr(Qry['Celular']);
        txtFax.Text := VarToStr(Qry['Fax']);
end;

procedure TfrmClientes.ClearData();
begin
        txtClave.Text := '';
        txtNombre.Text := '';
        txtContacto.Text := '';
        txtRFC.Text := '';
        txtCalle.Text := '';
        txtNumero.Text := '';
        txtColonia.Text := '';
        txtCP.Text := '';
        txtCiudad.Text := '';
        txtEstado.Text := '';
        txtTelefono.Text := '';
        txtCelular.Text := '';
        txtFax.Text := '';
end;

procedure TfrmClientes.EnableControls(Value:Boolean);
begin
        txtClave.ReadOnly := Value;
        txtNombre.ReadOnly := Value;
        txtContacto.ReadOnly := Value;
        txtRFC.ReadOnly := Value;
        txtCalle.ReadOnly := Value;
        txtNumero.ReadOnly := Value;
        txtColonia.ReadOnly := Value;
        txtCP.ReadOnly := Value;
        txtCiudad.ReadOnly := Value;
        txtEstado.ReadOnly := Value;
        txtTelefono.ReadOnly := Value;
        txtCelular.ReadOnly := Value;
        txtFax.ReadOnly := Value;
end;

procedure TfrmClientes.btnCancelarClick(Sender: TObject);
begin
EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;

BindCliente;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmClientes.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['Clave'] := txtClave.Text;
        Qry['Nombre'] := txtNombre.Text;
        Qry['Contacto'] := txtContacto.Text;
        Qry['RFC'] := txtRFC.Text;
        Qry['Calle'] := txtCalle.Text;
        Qry['Numero'] := txtNumero.Text;
        Qry['Colonia'] := txtColonia.Text;
        Qry['CP'] := txtCP.Text;
        Qry['Ciudad'] := txtCiudad.Text;
        Qry['Estado'] := txtEstado.Text;
        Qry['Telefono'] := txtTelefono.Text;
        Qry['Celular'] := txtCelular.Text;
        Qry['Fax'] := txtFax.Text;
        Qry.Post;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['Clave'] := txtClave.Text;
        Qry['Nombre'] := txtNombre.Text;
        Qry['Contacto'] := txtContacto.Text;
        Qry['RFC'] := txtRFC.Text;
        Qry['Calle'] := txtCalle.Text;
        Qry['Numero'] := txtNumero.Text;
        Qry['Colonia'] := txtColonia.Text;
        Qry['CP'] := txtCP.Text;
        Qry['Ciudad'] := txtCiudad.Text;
        Qry['Estado'] := txtEstado.Text;
        Qry['Telefono'] := txtTelefono.Text;
        Qry['Celular'] := txtCelular.Text;
        Qry['Fax'] := txtFax.Text;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar el cliente : ' +
                      txtNombre.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar el cliente : ' +
                            txtNombre.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtClave.Text <> '' then
        begin
              if not Qry.Locate('Clave',txtClave.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con clave ' + txtClave.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end
        else if txtNombre.Text <> '' then
        begin
              if not Qry.Locate('Nombre',txtNombre.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con nombre ' + txtNombre.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end
        else if txtContacto.Text <> '' then
        begin
              if not Qry.Locate('Contacto',txtContacto.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con Contacto ' + txtContacto.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end
        else if txtCalle.Text <> '' then
        begin
              if not Qry.Locate('Direccion',txtCalle.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con Direccion ' + txtCalle.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end
        else if txtCiudad.Text <> '' then
        begin
              if not Qry.Locate('Ciudad',txtCiudad.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con Ciudad ' + txtCiudad.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end
        else if txtEstado.Text <> '' then
        begin
              if not Qry.Locate('Estado',txtEstado.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Cliente con Estado ' + txtEstado.Text + '.', mtInformation,[mbOk], 0);
                    txtClave.SetFocus;
                    Exit;
                end;
        end;
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
BindCliente;
EnableFormButtons(gbButtons, sPermits);
end;

function TfrmClientes.ValidateData():Boolean;
begin
        ValidateData := True;
        if txtClave.Text = '' Then
          begin
            MessageDlg('La clave no puede estar vacia.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if txtNombre.Text = '' Then
          begin
            MessageDlg('El nombre del cliente no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtClave.Text) Then
          begin
            MessageDlg('La clave debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

end;


procedure TfrmClientes.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtClave.SetFocus;
giOpcion := 1;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmClientes.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
txtClave.SetFocus;
end;

procedure TfrmClientes.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmClientes.BuscarClick(Sender: TObject);
begin
ClearData();
txtClave.ReadOnly := False;
EnableControls(False);
txtClave.SetFocus;
giOpcion := 4;

btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
end;

procedure TfrmClientes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmClientes.txtClaveKeyPress(Sender: TObject; var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
       else
                Key := #0;

end;

end.