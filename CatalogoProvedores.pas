unit CatalogoProvedores;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,Larco_Functions;

type
  TfrmProvedores = class(TForm)
    gbButtons: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    lblId: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label10: TLabel;
    Label14: TLabel;
    txtNombre: TEdit;
    txtCiudad: TEdit;
    Primero: TButton;
    Anterior: TButton;
    Siguiente: TButton;
    Ultimo: TButton;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    txtContacto: TEdit;
    txtCalle: TEdit;
    txtEstado: TEdit;
    txtNumero: TEdit;
    txtColonia: TEdit;
    txtCP: TEdit;
    txtTelefono: TEdit;
    txtCelular: TEdit;
    txtFax: TEdit;
    txtRFC: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    Procedure BindData();
    Procedure ClearData();
    procedure SendTab(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PrimeroClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProvedores: TfrmProvedores;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmProvedores.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmProvedores.FormCreate(Sender: TObject);
begin
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblProvedores ORDER BY PROV_ID';
  Qry.Open;

  Nuevo.Enabled := True;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;

  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
        BindData();
  end;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmProvedores.EnableControls(Value:Boolean);
begin
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

procedure TfrmProvedores.ClearData();
begin
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

procedure TfrmProvedores.BindData();
begin
        lblId.Caption  := VarToStr(Qry['PROV_ID']);
        txtNombre.Text := VarToStr(Qry['PROV_Nombre']);
        txtContacto.Text := VarToStr(Qry['PROV_Contacto']);
        txtRFC.Text := VarToStr(Qry['PROV_RFC']);
        txtCalle.Text := VarToStr(Qry['PROV_Calle']);
        txtNumero.Text := VarToStr(Qry['PROV_Numero']);
        txtColonia.Text := VarToStr(Qry['PROV_Colonia']);
        txtCP.Text := VarToStr(Qry['PROV_CP']);
        txtCiudad.Text := VarToStr(Qry['PROV_Ciudad']);
        txtEstado.Text := VarToStr(Qry['PROV_Estado']);
        txtTelefono.Text := VarToStr(Qry['PROV_Telefono']);
        txtCelular.Text := VarToStr(Qry['PROV_Celular']);
        txtFax.Text := VarToStr(Qry['PROV_Fax']);
end;


procedure TfrmProvedores.PrimeroClick(Sender: TObject);
begin
  if Qry.RecordCount = 0 then
          Exit;

  if (Sender as TButton).Caption = '| <' then
    Qry.First
  else if (Sender as TButton).Caption = '<' then
    Qry.Prior
  else if (Sender as TButton).Caption = '>' then
    Qry.Next
  else if (Sender as TButton).Caption = '> |' then
    Qry.Last;


  BindData;
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmProvedores.btnCancelarClick(Sender: TObject);
begin
EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;

BindData();
EnableFormButtons(gbButtons, sPermits);
end;

function TfrmProvedores.ValidateData():Boolean;
begin
        ValidateData := True;
        if txtNombre.Text = '' Then
          begin
            MessageDlg('El nombre del Proveedor no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;
end;

procedure TfrmProvedores.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;
end;


procedure TfrmProvedores.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtNombre.SetFocus;
giOpcion := 1;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmProvedores.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
txtNombre.SetFocus;
end;

procedure TfrmProvedores.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmProvedores.BuscarClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtNombre.SetFocus;
giOpcion := 4;

btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
end;

procedure TfrmProvedores.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['PROV_Nombre'] := txtNombre.Text;
        Qry['PROV_Contacto'] := txtContacto.Text;
        Qry['PROV_RFC'] := txtRFC.Text;
        Qry['PROV_Calle'] := txtCalle.Text;
        Qry['PROV_Numero'] := txtNumero.Text;
        Qry['PROV_Colonia'] := txtColonia.Text;
        Qry['PROV_CP'] := txtCP.Text;
        Qry['PROV_Ciudad'] := txtCiudad.Text;
        Qry['PROV_Estado'] := txtEstado.Text;
        Qry['PROV_Telefono'] := txtTelefono.Text;
        Qry['PROV_Celular'] := txtCelular.Text;
        Qry['PROV_Fax'] := txtFax.Text;
        Qry.Post;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['PROV_Nombre'] := txtNombre.Text;
        Qry['PROV_Contacto'] := txtContacto.Text;
        Qry['PROV_RFC'] := txtRFC.Text;
        Qry['PROV_Calle'] := txtCalle.Text;
        Qry['PROV_Numero'] := txtNumero.Text;
        Qry['PROV_Colonia'] := txtColonia.Text;
        Qry['PROV_CP'] := txtCP.Text;
        Qry['PROV_Ciudad'] := txtCiudad.Text;
        Qry['PROV_Estado'] := txtEstado.Text;
        Qry['PROV_Telefono'] := txtTelefono.Text;
        Qry['PROV_Celular'] := txtCelular.Text;
        Qry['PROV_Fax'] := txtFax.Text;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar el Proveedor : ' +
                      txtNombre.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar el Proveedor : ' +
                            txtNombre.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtNombre.Text <> '' then
        begin
              if not Qry.Locate('PROV_Nombre',txtNombre.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Proveedor con nombre ' + txtNombre.Text + '.', mtInformation,[mbOk], 0);
                    txtNombre.SetFocus;
                    Exit;
                end;
        end
        else if txtContacto.Text <> '' then
        begin
              if not Qry.Locate('PROV_Contacto',txtContacto.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Proveedor con Contacto ' + txtContacto.Text + '.', mtInformation,[mbOk], 0);
                    txtContacto.SetFocus;
                    Exit;
                end;
        end
        else if txtCalle.Text <> '' then
        begin
              if not Qry.Locate('PROV_Direccion',txtCalle.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Proveedor con Direccion ' + txtCalle.Text + '.', mtInformation,[mbOk], 0);
                    txtCalle.SetFocus;
                    Exit;
                end;
        end
        else if txtCiudad.Text <> '' then
        begin
              if not Qry.Locate('PROV_Ciudad',txtCiudad.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Proveedor con Ciudad ' + txtCiudad.Text + '.', mtInformation,[mbOk], 0);
                    txtCiudad.SetFocus;
                    Exit;
                end;
        end
        else if txtEstado.Text <> '' then
        begin
              if not Qry.Locate('PROV_Estado',txtEstado.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Proveedor con Estado ' + txtEstado.Text + '.', mtInformation,[mbOk], 0);
                    txtEstado.SetFocus;
                    Exit;
                end;
        end;
  end;

  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  Nuevo.Enabled := True;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
  end;
  BindData();
  EnableFormButtons(gbButtons, sPermits);
end;

end.
