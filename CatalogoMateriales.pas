unit CatalogoMateriales;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,LTCUtils,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,Larco_Functions;

type
  TfrmMateriales = class(TForm)
    gbButtons: TGroupBox;
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
    GroupBox2: TGroupBox;
    Label2: TLabel;
    txtFraccion: TEdit;
    Label10: TLabel;
    txtDescripcion: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    lblId: TLabel;
    ddlUnidad: TComboBox;
    ddlTipos: TComboBox;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    txtCantidad: TEdit;
    Label3: TLabel;
    txtCosto: TEdit;
    Label6: TLabel;
    txtUltimoCosto: TEdit;
    Label7: TLabel;
    txtCostoPromedio: TEdit;
    txtMaximo: TEdit;
    txtMinimo: TEdit;
    Maximo: TLabel;
    Label9: TLabel;
    Label8: TLabel;
    txtStock: TEdit;
    Actualizar: TButton;
    Label11: TLabel;
    txtID: TEdit;
    Label12: TLabel;
    txtUbicacion: TEdit;
    ddlDesc: TComboBox;
    Label13: TLabel;
    txtProvId: TEdit;
    chkKilos: TCheckBox;
    Label14: TLabel;
    txtDensidad: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindData();
    procedure BindUnidades();
    procedure BindTipos();
    procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure PrimeroClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure ActualizarClick(Sender: TObject);
    procedure txtDescripcionKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ddlDescSelect(Sender: TObject);
    procedure ddlDescCloseUp(Sender: TObject);
    procedure ddlDescKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chkKilosClick(Sender: TObject);
    function BoolToStrInt(Value:Boolean):String;
    procedure txtDensidadKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMateriales: TfrmMateriales;
  gsUnidadesIds, gsTiposIds: TStringList;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmMateriales.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmMateriales.FormCreate(Sender: TObject);
begin
  gsUnidadesIds := TStringList.Create;
  gsTiposIds := TStringList.Create;

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblMateriales ORDER BY MAT_ID';
  Qry.Open;

  BindUnidades();
  BindTipos();

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

procedure TfrmMateriales.BindData();
begin
  lblId.Caption  := VarToStr(Qry['MAT_ID']);
  txtFraccion.Text := VarToStr(Qry['MAT_Fraccion']);
  txtDescripcion.Text := VarToStr(Qry['MAT_Descripcion']);
  txtID.Text := VarToStr(Qry['MAT_Numero']);
  ddlUnidad.ItemIndex := getItemIndex(VarToStr(Qry['MAT_Unidad']), gsUnidadesIds);
  ddlTipos.ItemIndex := getItemIndex(VarToStr(Qry['MAT_Tipo']), gsTiposIds);
  txtCantidad.Text := VarToStr(Qry['MAT_Cantidad']);
  txtCosto.Text := VarToStr(Qry['MAT_Costo']);
  txtUltimoCosto.Text := VarToStr(Qry['MAT_UltimoCosto']);
  txtCostoPromedio.Text := VarToStr(Qry['MAT_CostoPromedio']);
  txtMaximo.Text := VarToStr(Qry['MAT_Maximo']);
  txtMinimo.Text := VarToStr(Qry['MAT_Minimo']);
  txtStock.Text := VarToStr(Qry['MAT_Stock']);
  txtUbicacion.Text :=  VarToStr(Qry['MAT_Ubicacion']);
  txtProvId.Text :=  VarToStr(Qry['MAT_ProvNumero']);
  chkKilos.Checked := StrToBool(VarToStr(Qry['MAT_Kilos']));
  txtDensidad.Text :=  VarToStr(Qry['MAT_Densidad']);
end;

procedure TfrmMateriales.ClearData();
begin
  txtFraccion.Text := '';
  txtDescripcion.Text := '';
  txtID.Text := '';
  ddlTipos.ItemIndex := 0;
  ddlUnidad.ItemIndex := 0;
  txtCantidad.Text := '';
  txtCosto.Text := '';
  txtUltimoCosto.Text := '';
  txtCostoPromedio.Text := '';
  txtMaximo.Text := '';
  txtMinimo.Text := '';
  txtStock.Text := '';
  txtUbicacion.Text := '';
  txtProvId.Text := '';
  chkKilos.Checked := False;
  txtDensidad.Text := '';
end;

procedure TfrmMateriales.EnableControls(Value:Boolean);
begin
  txtFraccion.ReadOnly := Value;
  txtDescripcion.ReadOnly := Value;
  txtID.ReadOnly := Value;
  ddlUnidad.Enabled := not Value;
  ddlTipos.Enabled := not Value;
  txtCantidad.ReadOnly := Value;
  txtCosto.ReadOnly := Value;
  txtUltimoCosto.ReadOnly := Value;
  txtCostoPromedio.ReadOnly := Value;
  txtMaximo.ReadOnly := Value;
  txtMinimo.ReadOnly := Value;
  txtStock.ReadOnly := Value;
  txtUbicacion.ReadOnly := Value;
  txtProvId.ReadOnly := Value;
  chkKilos.Enabled := not Value;
end;


procedure TfrmMateriales.PrimeroClick(Sender: TObject);
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

procedure TfrmMateriales.btnCancelarClick(Sender: TObject);
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

procedure TfrmMateriales.NuevoClick(Sender: TObject);
var Qry2 : TADOQuery;
SQLStr,secuencia : String;
fSecuencia : Double;
begin
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT MAX(left(MAT_Numero,4)) AS Secuencia FROM tblMateriales';

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  secuencia := '';
  if(Qry2.RecordCount > 0) then begin
        secuencia := VarToStr(Qry2['Secuencia']);
        fSecuencia := StrToFloat(Secuencia) + 1;
        secuencia := FormatFloat('0000', fSecuencia) + '-';
  end;

  giOpcion := 1;
  ClearData();
  EnableControls(False);
  txtId.SetFocus;
  txtId.Text := secuencia;
  txtId.SelStart := 5;//Length(secuencia);
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  ddlUnidad.ItemIndex := ddlUnidad.Items.IndexOf('pza');
end;

procedure TfrmMateriales.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  txtFraccion.SetFocus;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmMateriales.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmMateriales.BuscarClick(Sender: TObject);
begin
  ClearData();
  EnableControls(False);
  txtFraccion.SetFocus;
  giOpcion := 4;

  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
end;

procedure TfrmMateriales.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        lblId.Caption := '';
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['MAT_Fraccion'] := txtFraccion.Text;
        Qry['MAT_Descripcion'] := txtDescripcion.Text;
        Qry['MAT_Numero'] := txtID.Text;
        Qry['MAT_Unidad'] := gsUnidadesIds[ddlUnidad.ItemIndex];
        Qry['MAT_Tipo'] := gsTiposIds[ddlTipos.ItemIndex];
        Qry['MAT_Cantidad'] := getDecimalValue(txtCantidad.Text);
        Qry['MAT_Costo'] := getDecimalValue(txtCosto.Text);
        Qry['MAT_UltimoCosto'] := getDecimalValue(txtUltimoCosto.Text);
        Qry['MAT_CostoPromedio'] := getDecimalValue(txtCostoPromedio.Text);
        Qry['MAT_Minimo'] := getDecimalValue(txtMinimo.Text);
        Qry['MAT_Maximo'] := getDecimalValue(txtMaximo.Text);
        Qry['MAT_Stock'] := getDecimalValue(txtStock.Text);
        Qry['MAT_Usuario'] := frmMain.sUserID;
        Qry['MAT_Fecha'] := Date();
        Qry['MAT_Ubicacion'] := txtUbicacion.Text;
        Qry['MAT_ProvNumero'] := txtProvId.Text;
        Qry['MAT_Kilos'] := BoolToStrInt(chkKilos.Checked);
        if txtDensidad.Text = '' then txtDensidad.Text := '0.0';
        Qry['MAT_Densidad'] := txtDensidad.Text;
        Qry.Post;

        Qry.SQL.Clear;
        Qry.SQL.Text := 'SELECT * FROM tblMateriales ORDER BY MAT_ID';
        Qry.Open;
        Qry.Last;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['MAT_Fraccion'] := txtFraccion.Text;
        Qry['MAT_Descripcion'] := txtDescripcion.Text;
        Qry['MAT_Numero'] := txtID.Text;        
        Qry['MAT_Unidad'] := gsUnidadesIds[ddlUnidad.ItemIndex];
        Qry['MAT_Tipo'] := gsTiposIds[ddlTipos.ItemIndex];
        Qry['MAT_Cantidad'] := getDecimalValue(txtCantidad.Text);
        Qry['MAT_Costo'] := getDecimalValue(txtCosto.Text);
        Qry['MAT_UltimoCosto'] := getDecimalValue(txtUltimoCosto.Text);
        Qry['MAT_CostoPromedio'] := getDecimalValue(txtCostoPromedio.Text);
        Qry['MAT_Minimo'] := getDecimalValue(txtMinimo.Text);
        Qry['MAT_Maximo'] := getDecimalValue(txtMaximo.Text);
        Qry['MAT_Stock'] := getDecimalValue(txtStock.Text);
        Qry['MAT_Usuario'] := frmMain.sUserID;
        Qry['MAT_Fecha'] := Date();
        Qry['MAT_Ubicacion'] := txtUbicacion.Text;
        Qry['MAT_ProvNumero'] := txtProvId.Text;
        Qry['MAT_Kilos'] := BoolToStrInt(chkKilos.Checked);
        if txtDensidad.Text = '' then txtDensidad.Text := '0.0';        
        Qry['MAT_Densidad'] := txtDensidad.Text;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar el Material : ' +
                      txtDescripcion.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar el Material : ' +
                            txtDescripcion.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtFraccion.Text <> '' then
        begin
              if not Qry.Locate('MAT_Fraccion',txtFraccion.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Material con Fraccion Arancelaria ' + txtFraccion.Text + '.', mtInformation,[mbOk], 0);
                    txtFraccion.SetFocus;
                    Exit;
                end;
        end
        else if txtDescripcion.Text <> '' then
        begin
              if not Qry.Locate('MAT_Descripcion',txtDescripcion.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Material con Descripcion ' + txtDescripcion.Text + '.', mtInformation,[mbOk], 0);
                    txtDescripcion.SetFocus;
                    Exit;
                end;
        end
        else if ddlUnidad.Text <> '' then
        begin
              if not Qry.Locate('MAT_Unidad',gsUnidadesIds[ddlUnidad.ItemIndex] ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Material con Tipo de Unidad ' + ddlUnidad.Text + '.', mtInformation,[mbOk], 0);
                    ddlUnidad.SetFocus;
                    Exit;
                end;
        end
        else if ddlTipos.Text <> '' then
        begin
              if not Qry.Locate('MAT_Tipo',gsTiposIds[ddlTipos.ItemIndex] ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Material con Tipo de Material ' + ddlTipos.Text + '.', mtInformation,[mbOk], 0);
                    ddlTipos.SetFocus;
                    Exit;
                end;
        end;
  end;

  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Nuevo.Enabled := True;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
  end;
  BindData();
  giOpcion := 0;
  EnableFormButtons(gbButtons, sPermits);
end;

function TfrmMateriales.ValidateData():Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
    result :=  True;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    if(lblId.Caption = '') then begin
            SQLStr := 'SELECT MAT_ID FROM tblMateriales WHERE MAT_Numero = ' + QuotedStr(txtID.Text);
    end
    else begin
            SQLStr := 'SELECT MAT_ID FROM tblMateriales WHERE MAT_Numero = ' + QuotedStr(txtID.Text) +
                      ' AND MAT_ID <> ' + lblID.Caption;
    end;
    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if(Qry2.RecordCount > 0) then begin
            MessageDlg('Este ID de material ya existe.' , mtInformation,[mbOk], 0);
            result :=  False;
    end;

    if(lblId.Caption = '') then begin
            SQLStr := 'SELECT MAT_ID FROM tblMateriales WHERE Left(MAT_Numero,4) = ' + QuotedStr(LeftStr(txtID.Text,4));
    end
    else begin
            SQLStr := 'SELECT MAT_ID FROM tblMateriales WHERE Left(MAT_Numero,4) = ' + QuotedStr(LeftStr(txtID.Text,4)) +
                      ' AND MAT_ID <> ' + lblID.Caption;
    end;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if(Qry2.RecordCount > 0) then begin
            MessageDlg('Este consecutivo ya existe.' , mtInformation,[mbOk], 0);
            result :=  False;
    end;

end;

procedure TfrmMateriales.BindUnidades();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT * FROM tblUnidadesMedida ORDER BY UNI_Medida';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlUnidad.Items.Clear;
    gsUnidadesIds.Clear;
    While not Qry2.Eof do
    Begin
        ddlUnidad.Items.Add(Qry2['UNI_Medida']);
        gsUnidadesIds.Add(VarToStr(Qry2['UNI_ID']));
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmMateriales.BindTipos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT * FROM tblTiposMaterial ORDER BY TIP_Descripcion';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlTipos.Items.Clear;
    gsTiposIds.Clear;
    While not Qry2.Eof do
    Begin
        ddlTipos.Items.Add(Qry2['TIP_Descripcion']);
        gsTiposIds.Add(VarToStr(Qry2['TIP_ID']));
        Qry2.Next;
    End;

    Qry2.Close;
end;


procedure TfrmMateriales.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar((Sender as TEdit).Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;

end;

procedure TfrmMateriales.ActualizarClick(Sender: TObject);
begin
  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblMateriales ORDER BY MAT_ID';
  Qry.Open;

  BindUnidades();
  BindTipos();

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
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmMateriales.txtDescripcionKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
var selectedText, value: String;
p, found: integer;
begin

  with txtDescripcion do
    case key of 13,37..39:; //No enter, or arrows except down
    else
    begin
      if key = 40 then begin//abajo
        if ddlDesc.DroppedDown then begin
            ddlDesc.SetFocus;
            ddlDesc.ItemIndex := 0;
            ddlDesc.Text := ddlDesc.Items[0];
            txtDescripcion.Text := ddlDesc.Text;
            Application.ProcessMessages;
        end;
      end
      else  begin
          ddlDesc.Items.Clear;      
          p := selStart;
          selectedText := copy(text, 0, p);

          Qry.First;
          While not Qry.Eof do
          Begin
            value := VarToStr(Qry['MAT_Descripcion']);
            found := InStr(0,UT(value),UT(selectedText));
            if (found <> 0) then begin
              ddlDesc.Items.Add(value);
            end;

            Qry.Next;
          end;
          Qry.Locate('MAT_ID',lblId.Caption ,[loPartialKey] )
      end;
    end;//case
  end;//with

  if ddlDesc.Items.Count > 0 then
    ddlDesc.DroppedDown := True
  else
    ddlDesc.DroppedDown := False;
    
end;

procedure TfrmMateriales.ddlDescSelect(Sender: TObject);
begin
  txtDescripcion.Text := ddlDesc.Text;
end;

procedure TfrmMateriales.ddlDescCloseUp(Sender: TObject);
begin
        ddlDesc.Items.Clear;  
        txtDescripcion.SetFocus;
        txtDescripcion.SelStart := Length(txtDescripcion.Text);
end;

procedure TfrmMateriales.ddlDescKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
        if not (key in [37..40]) then begin
          txtDescripcion.SetFocus;
          txtDescripcion.SelStart := Length(txtDescripcion.Text);
        end;
end;

procedure TfrmMateriales.chkKilosClick(Sender: TObject);
begin
  txtDensidad.Enabled := chkKilos.Checked;
  txtDensidad.ReadOnly := not chkKilos.Checked;
end;

function TfrmMateriales.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;

procedure TfrmMateriales.txtDensidadKeyPress(Sender: TObject;
  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar(txtDensidad.Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;
end;

end.
