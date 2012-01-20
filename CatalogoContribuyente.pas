unit CatalogoContribuyente;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,Larco_Functions;

type
  TfrmContribuyente = class(TForm)
    gbButtons: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    lblId: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label14: TLabel;
    txtRazon: TEdit;
    txtEntidad: TEdit;
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
    txtCalle: TEdit;
    txtNumero: TEdit;
    txtColonia: TEdit;
    txtCP: TEdit;
    txtRFC: TEdit;
    Label4: TLabel;
    txtRegistro: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PrimeroClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BindData();
    procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure btnCancelarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmContribuyente: TfrmContribuyente;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;  
implementation

uses Main;

{$R *.dfm}

procedure TfrmContribuyente.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmContribuyente.PrimeroClick(Sender: TObject);
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

procedure TfrmContribuyente.FormCreate(Sender: TObject);
begin
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblContribuyente ORDER BY CON_ID';
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

procedure TfrmContribuyente.BindData();
begin
  lblId.Caption  := VarToStr(Qry['CON_ID']);
  txtRazon.Text := VarToStr(Qry['CON_RazonSocial']);
  txtRFC.Text := VarToStr(Qry['CON_RFC']);
  txtRegistro.Text := VarToStr(Qry['CON_Registro']);
  txtCalle.Text := VarToStr(Qry['CON_Calle']);
  txtNumero.Text := VarToStr(Qry['CON_Numero']);
  txtColonia.Text := VarToStr(Qry['CON_Colonia']);
  txtCP.Text := VarToStr(Qry['CON_CP']);
  txtEntidad.Text := VarToStr(Qry['CON_Entidad']);
end;

procedure TfrmContribuyente.ClearData();
begin
  txtRazon.Text := '';
  txtRFC.Text := '';
  txtRegistro.Text := '';
  txtCalle.Text := '';
  txtNumero.Text := '';
  txtColonia.Text := '';
  txtCP.Text := '';
  txtEntidad.Text := '';
end;

procedure TfrmContribuyente.EnableControls(Value:Boolean);
begin
  txtRazon.ReadOnly := Value;
  txtRFC.ReadOnly := Value;
  txtRegistro.ReadOnly := Value;
  txtCalle.ReadOnly := Value;
  txtNumero.ReadOnly := Value;
  txtColonia.ReadOnly := Value;
  txtCP.ReadOnly := Value;
  txtEntidad.ReadOnly := Value;
end;

procedure TfrmContribuyente.btnCancelarClick(Sender: TObject);
begin
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;

  BindData;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmContribuyente.NuevoClick(Sender: TObject);
begin
  giOpcion := 1;
  ClearData();
  EnableControls(False);
  txtRazon.SetFocus;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmContribuyente.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  txtRazon.SetFocus;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmContribuyente.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmContribuyente.BuscarClick(Sender: TObject);
begin
  ClearData();
  EnableControls(False);
  txtRazon.SetFocus;
  giOpcion := 4;

  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
end;

procedure TfrmContribuyente.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['CON_RazonSocial'] := txtRazon.Text;
        Qry['CON_RFC'] := txtRFC.Text;
        Qry['CON_Registro'] := txtRegistro.Text;
        Qry['CON_Calle'] := txtCalle.Text;
        Qry['CON_Numero'] := txtNumero.Text;
        Qry['CON_Colonia'] := txtNumero.Text;
        Qry['CON_CP'] := txtCP.Text;
        Qry['CON_Entidad'] := txtEntidad.Text;
        Qry.Post;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['CON_RazonSocial'] := txtRazon.Text;
        Qry['CON_RFC'] := txtRFC.Text;
        Qry['CON_Registro'] := txtRegistro.Text;
        Qry['CON_Calle'] := txtCalle.Text;
        Qry['CON_Numero'] := txtNumero.Text;
        Qry['CON_Colonia'] := txtNumero.Text;
        Qry['CON_CP'] := txtCP.Text;
        Qry['CON_Entidad'] := txtEntidad.Text;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar al Contribuyente : ' +
                      txtRazon.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar al Contribuyente : ' +
                            txtRazon.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtRazon.Text <> '' then
        begin
              if not Qry.Locate('CON_RazonSocial',txtRazon.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con Razon Social ' + txtRazon.Text + '.', mtInformation,[mbOk], 0);
                    txtRazon.SetFocus;
                    Exit;
                end;
        end
        else if txtRFC.Text <> '' then
        begin
              if not Qry.Locate('CON_RFC',txtRFC.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con RFC ' + txtRFC.Text + '.', mtInformation,[mbOk], 0);
                    txtRFC.SetFocus;
                    Exit;
                end;
        end
        else if txtRegistro.Text <> '' then
        begin
              if not Qry.Locate('CON_Registro',txtRegistro.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con Numero de Registro ' + txtRegistro.Text + '.', mtInformation,[mbOk], 0);
                    txtRegistro.SetFocus;
                    Exit;
                end;
        end
        else if txtCalle.Text <> '' then
        begin
              if not Qry.Locate('CON_Calle',txtCalle.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con Calle ' + txtCalle.Text + '.', mtInformation,[mbOk], 0);
                    txtCalle.SetFocus;
                    Exit;
                end;
        end
        else if txtNumero.Text <> '' then
        begin
              if not Qry.Locate('CON_Numero',txtNumero.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con Numero ' + txtNumero.Text + '.', mtInformation,[mbOk], 0);
                    txtNumero.SetFocus;
                    Exit;
                end;
        end
        else if txtCP.Text <> '' then
        begin
              if not Qry.Locate('CON_CP',txtCP.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Contribuyente con Codigo Postal ' + txtCP.Text + '.', mtInformation,[mbOk], 0);
                    txtCP.SetFocus;
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

function TfrmContribuyente.ValidateData():Boolean;
begin
        ValidateData := True;
        if txtRazon.Text = '' Then
          begin
            MessageDlg('Por favor especifique la Denominacion o Razon Social.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if txtRFC.Text = '' Then
          begin
            MessageDlg('Por favor especifique el RFC.', mtInformation,[mbOk], 0);
            result :=  False;
          end;
end;

end.
