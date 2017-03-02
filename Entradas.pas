unit Entradas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math;

type
  TfrmEntradas = class(TForm)
    gbButtons: TGroupBox;
    GroupBox2: TGroupBox;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    Imprimir: TButton;
    Panel1: TPanel;
    Primero: TButton;
    Anterior: TButton;
    Ultimo: TButton;
    Siguiente: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    txtPedimento: TEdit;
    txtClavePedimento: TEdit;
    txtFactura: TEdit;
    txtOrdenCompra: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    deFecha: TDateEditor;
    Label3: TLabel;
    Label4: TLabel;
    ddlPais: TComboBox;
    Label9: TLabel;
    OrdenCompra: TLabel;
    lblProvedor: TLabel;
    IVA: TLabel;
    ddlProvedor: TComboBox;
    txtIVA: TEdit;
    gvEntradas: TGridView;
    ddlMaterial: TComboBox;
    Label7: TLabel;
    Label8: TLabel;
    txtCantidad: TEdit;
    Label10: TLabel;
    txtCosto: TEdit;
    btnAdd: TButton;
    btnDelete: TButton;
    btnClear: TButton;
    Label13: TLabel;
    txtSubtotal: TEdit;
    Label14: TLabel;
    txtTIVA: TEdit;
    Label15: TLabel;
    txtTotal: TEdit;
    lblId: TLabel;
    lblAnio: TLabel;
    txtID: TEdit;
    lblMaterialDesc: TLabel;
    ddlNacional: TComboBox;
    ddlTipoImp: TComboBox;
    Label5: TLabel;
    Label6: TLabel;
    chkDlls: TCheckBox;
    txtTipo: TEdit;
    Label11: TLabel;
    gbMateriales: TGroupBox;
    tvMateriales: TTreeView;
    btnOrden: TButton;
    Label12: TLabel;
    txtProvId: TEdit;
    lblDiametro: TLabel;
    txtDiametro: TEdit;
    lblLongitud: TLabel;
    txtLongitud: TEdit;
    lblPies: TLabel;
    lblPulgadas: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure EnableButtons();
    Procedure BindData();
    Procedure ClearData();
    procedure SendTab(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PrimeroClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure BindPaises();
    procedure BindProvedores();
    procedure BindMateriales();
    function ValidateData():Boolean;
    procedure txtTasaGeneralKeyPress(Sender: TObject; var Key: Char);
    procedure txtIVAKeyPress(Sender: TObject; var Key: Char);
    procedure btnClearClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure calcularTotal();
    procedure btnAddClick(Sender: TObject);
    procedure BindDetalle(EntradaID: String);
    procedure ActualizarDetalle(EntradaID: String);
    procedure gvEntradasAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: String; var Accept: Boolean);
    procedure ddlNacionalChange(Sender: TObject);
    procedure txtIDExit(Sender: TObject);
    procedure txtIDKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    function BoolToStrInt(Value:Boolean):String;
    procedure btnOrdenClick(Sender: TObject);
    procedure tvMaterialesDblClick(Sender: TObject);
    procedure txtProvIdKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure displayKilos(i: integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEntradas: TfrmEntradas;
  giOpcion : Integer;
  gsPaisesIds, gsProvIds, gsMaterialesIds, gsMaterialesNumeros, gsProvNumeroIds, gsKilos: TStringList;
  gsDensidad: TStringList;
  Conn : TADOConnection;
  Qry : TADOQuery;
  gbMaterialValido, gbKilos : boolean;
  sDensidad, sPermits : String;
implementation

uses Main, Login;

{$R *.dfm}

procedure TfrmEntradas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmEntradas.FormCreate(Sender: TObject);
var control : TControl;
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);
  lblMaterialDesc.Caption := '';

  gsPaisesIds := TStringList.Create;
  gsProvIds := TStringList.Create;
  gsMaterialesIds := TStringList.Create;
  gsMaterialesNumeros := TStringList.Create;
  gsProvNumeroIds := TStringList.Create;
  gsKilos := TStringList.Create;
  gsDensidad := TStringList.Create;
  gbMaterialValido := false;
  gbKilos := false;
  sDensidad := '';
  
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblEntradas WHERE YEAR(ENT_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' ORDER BY ENT_ID';
  Qry.Open;

  BindPaises();
  BindProvedores();
  BindMateriales();
  ClearData();

  EnableButtons();
  BindData();
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEntradas.EnableControls(Value:Boolean);
begin
  txtPedimento.ReadOnly := Value;
  txtClavePedimento.ReadOnly := Value;
  deFecha.Enabled := not Value;
  ddlPais.Enabled := not Value;
  ddlNacional.Enabled := not Value;
  ddlTipoImp.Enabled := not Value;
  txtFactura.ReadOnly := Value;
  txtOrdenCompra.ReadOnly := Value;
  ddlProvedor.Enabled := not Value;
  txtIVA.ReadOnly := Value;
  txtTipo.ReadOnly := Value;
  chkDlls.Enabled := not Value;

  ddlMaterial.Enabled := not Value;
  txtID.ReadOnly := Value;
  txtCantidad.ReadOnly := Value;
  txtCosto.ReadOnly := Value;
  btnAdd.Enabled := not Value;
  btnDelete.Enabled := not Value;
  btnClear.Enabled :=  not Value;
  gvEntradas.Enabled := not Value;
  btnOrden.Enabled := not Value;
  txtProvId.ReadOnly := Value;
end;

procedure TfrmEntradas.EnableButtons();
begin
  Nuevo.Enabled := True;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
        Imprimir.Enabled := True;
  end;
  EnableFormButtons(gbButtons, sPermits);  
end;


procedure TfrmEntradas.ClearData();
begin
  txtPedimento.Text := '';
  txtClavePedimento.Text := '';
  deFecha.Text := DateToStr(Now);
  ddlPais.ItemIndex := 0;
  ddlNacional.ItemIndex := 0;
  ddlTipoImp.ItemIndex := 0;
  txtFactura.Text := '';
  txtOrdenCompra.Text := '';
  ddlProvedor.ItemIndex := 0;
  txtIVA.Text := '';

  ddlMaterial.ItemIndex := 0;
  txtId.Text := '';
  txtCantidad.Text := '';
  txtCosto.Text := '';
  txtTipo.Text := '';
  chkDlls.Checked := False;
  lblMaterialDesc.Caption := '';
  gbMaterialValido := False;
  btnClearClick(nil);
  gbMateriales.Visible := False;
  txtProvId.Text := '';
  txtDiametro.Text := '';
  txtLongitud.Text := '';

  txtLongitud.Visible := False;
  txtDiametro.Visible := False;
  lblLongitud.Visible := False;
  lblDiametro.Visible := False;

  txtLongitud.Enabled := False;
  txtDiametro.Enabled := False;

  lblPulgadas.Visible := False;
  lblPies.Visible := False;

  if gbKilos = True then begin
    btnAdd.Left := btnAdd.Left - 300;
    btnDelete.Left := btnDelete.Left - 300;
    btnClear.Left := btnClear.Left - 300;
  end;
  gbKilos := false;
end;

procedure TfrmEntradas.BindData();
begin
  if Qry.RecordCount = 0 then
          Exit;

  lblId.Caption  := VarToStr(Qry['ENT_ID']);
  txtPedimento.Text := VarToStr(Qry['ENT_Pedimento']);
  txtClavePedimento.Text := VarToStr(Qry['ENT_ClavePedimento']);
  deFecha.Text := VarToStr(Qry['ENT_Fecha']);
  ddlPais.ItemIndex := getItemIndex(VarToStr(Qry['ENT_PaisOrigen']), gsPaisesIds);
  ddlNacional.ItemIndex := ddlNacional.Items.IndexOf(VarToStr(Qry['ENT_Nacional']));
  ddlTipoImp.ItemIndex := ddlTipoImp.Items.IndexOf(VarToStr(Qry['ENT_TipoImp']));
  txtFactura.Text := VarToStr(Qry['ENT_Factura']);
  txtOrdenCompra.Text := VarToStr(Qry['ENT_OrdenCompra']);
  ddlProvedor.ItemIndex := getItemIndex(VarToStr(Qry['PROV_ID']), gsProvIds);
  txtIVA.Text := VarToStr(Qry['ENT_IVA']);
  txtTipo.Text := FormatFloat('###0.00',StrToFloat(VarToStr(Qry['ENT_TipoCambio'])));

  chkDlls.Checked := StrToBool(VarToStr(Qry['ENT_Dolares']));

  BindDetalle(lblId.Caption);
  calcularTotal();
end;

procedure TfrmEntradas.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;
end;

procedure TfrmEntradas.PrimeroClick(Sender: TObject);
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


  BindData();
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;
  Imprimir.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);  
end;

procedure TfrmEntradas.NuevoClick(Sender: TObject);
begin
  ClearData();
  EnableControls(False);
  txtPedimento.SetFocus;
  giOpcion := 1;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
  txtTipo.Text := '11.0';
end;

procedure TfrmEntradas.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
  txtPedimento.SetFocus;
end;

procedure TfrmEntradas.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
end;

procedure TfrmEntradas.btnCancelarClick(Sender: TObject);
begin
  ClearData();
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  EnableButtons();
  BindData();
end;

procedure TfrmEntradas.btnAceptarClick(Sender: TObject);
var user : String;
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        if txtTipo.Text = '' then txtTipo.Text := '11.00';

        Qry.Insert;
        Qry['ENT_Pedimento'] := txtPedimento.Text;
        Qry['ENT_ClavePedimento'] := txtClavePedimento.Text;
        Qry['ENT_Fecha'] := deFecha.Text;
        Qry['ENT_PaisOrigen'] := gsPaisesIds[ddlPais.ItemIndex];
        Qry['ENT_Nacional'] := ddlNacional.Text;
        Qry['ENT_TipoImp'] := ddlTipoImp.Text;
        Qry['ENT_Factura'] := txtFactura.Text;
        Qry['ENT_OrdenCompra'] := txtOrdenCompra.Text;
        Qry['PROV_ID'] := gsProvIds[ddlProvedor.ItemIndex];
        Qry['ENT_IVA'] := txtIVA.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry['ENT_Dolares'] := BoolToStrInt(chkDlls.Checked);
        Qry['ENT_TipoCambio'] := txtTipo.Text;
        Qry.Post;

        ActualizarDetalle(Qry['ENT_ID']);
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        if txtTipo.Text = '' then txtTipo.Text := '11.00';


        Qry.Edit;
        Qry['ENT_Pedimento'] := txtPedimento.Text;
        Qry['ENT_ClavePedimento'] := txtClavePedimento.Text;
        Qry['ENT_Fecha'] := deFecha.Text;
        Qry['ENT_PaisOrigen'] := gsPaisesIds[ddlPais.ItemIndex];
        Qry['ENT_Nacional'] := ddlNacional.Text;
        Qry['ENT_TipoImp'] := ddlTipoImp.Text;
        Qry['ENT_Factura'] := txtFactura.Text;
        Qry['ENT_OrdenCompra'] := txtOrdenCompra.Text;
        Qry['PROV_ID'] := gsProvIds[ddlProvedor.ItemIndex];
        Qry['ENT_IVA'] := txtIVA.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry['ENT_Dolares'] := BoolToStrInt(chkDlls.Checked);
        Qry['ENT_TipoCambio'] := txtTipo.Text;
        Qry.Post;

        ActualizarDetalle(Qry['ENT_ID']);
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar esta Entrada?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro esta Entrada?',
                        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                      Application.CreateForm(TfrmLogin, frmLogin);
                      frmLogin.lblValidate.Caption := 'true';
                      if frmLogin.ShowModal <> mrOK then begin
                            ShowMessage('No tienes permiso para scrapear esta orden.');
                      end
                      else begin
                          user := frmLogin.txtUser.Text;

                          Qry.Edit;
                          Qry['USE_ID'] := user;
                          Qry.Post;

                          Qry.Delete;
                          gvEntradas.ClearRows;
                          ActualizarDetalle(lblId.Caption);
                      end;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtPedimento.Text <> '' then
        begin
              if not Qry.Locate('ENT_Pedimento',txtPedimento.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Entrada con Pedimento ' + txtPedimento.Text + '.', mtInformation,[mbOk], 0);
                    txtPedimento.SetFocus;
                    Exit;
                end;
        end
        else if txtClavePedimento.Text <> '' then
        begin
              if not Qry.Locate('ENT_ClavePedimento',txtClavePedimento.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Entrada con Clave de Pedimento ' + txtClavePedimento.Text + '.', mtInformation,[mbOk], 0);
                    txtClavePedimento.SetFocus;
                    Exit;
                end;
        end
        else if deFecha.Text <> '' then
        begin
              if not Qry.Locate('ENT_Fecha',deFecha.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Entrada con Fecha ' + deFecha.Text + '.', mtInformation,[mbOk], 0);
                    deFecha.SetFocus;
                    Exit;
                end;
        end
        else if ddlPais.Text <> '' then
        begin
              if not Qry.Locate('ENT_PaisOrigen',gsPaisesIds[ddlPais.ItemIndex] ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Entrada con Pais ' + gsPaisesIds[ddlPais.ItemIndex] + '.', mtInformation,[mbOk], 0);
                    ddlPais.SetFocus;
                    Exit;
                end;
        end
        else if ddlProvedor.Text <> '' then
        begin
              if not Qry.Locate('PROV_ID',gsProvIds[ddlProvedor.ItemIndex] ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguns Entrada con Proveedor ' + gsProvIds[ddlProvedor.ItemIndex] + '.', mtInformation,[mbOk], 0);
                    ddlProvedor.SetFocus;
                    Exit;
                end;
        end;
  end;

  ClearData();
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  EnableButtons();
  BindData();
end;

procedure TfrmEntradas.BindPaises();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT * FROM tblPaises ORDER BY PAIS_Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlPais.Items.Clear;
    gsPaisesIds.Clear;
    While not Qry2.Eof do
    Begin
        ddlPais.Items.Add(Qry2['PAIS_Nombre']);
        gsPaisesIds.Add(VarToStr(Qry2['PAIS_ID']));
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmEntradas.BindProvedores();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT PROV_ID, PROV_Nombre FROM tblProvedores ORDER BY PROV_Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlProvedor.Items.Clear;
    gsProvIds.Clear;
    While not Qry2.Eof do
    Begin
        ddlProvedor.Items.Add(Qry2['PROV_Nombre']);
        gsProvIds.Add(VarToStr(Qry2['PROV_ID']));
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmEntradas.BindMateriales();
var Qry2 : TADOQuery;
SQLStr, tipo : String;
treeNode : TTreeNode;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT M.MAT_ID, M.MAT_Descripcion, M.MAT_Numero, T.TIP_Descripcion, ' +
              'M.MAT_ProvNumero, M.MAT_Kilos, M.MAT_Densidad ' +
              'FROM tblMateriales M ' +
              'INNER JOIN tblTiposMaterial T ON M.MAT_Tipo = T.TIP_ID ' +
              'ORDER BY TIP_Descripcion';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlMaterial.Items.Clear;
    gsMaterialesIds.Clear;
    gsProvNumeroIds.Clear;
    gsMaterialesNumeros.Clear;
    gsKilos.Clear;
    gsDensidad.Clear;
    tipo := '';
    tvMateriales.Items.Clear;
    While not Qry2.Eof do
    Begin
        if tipo <> VarToStr(Qry2['TIP_Descripcion']) then
        begin
                treeNode := tvMateriales.Items.Add(nil, VarToStr(Qry2['TIP_Descripcion']));
        end;
        tipo := VarToStr(Qry2['TIP_Descripcion']);
        tvMateriales.Items.AddChild(treeNode, VarToStr(Qry2['MAT_Descripcion']));

        ddlMaterial.Items.Add(VarToStr(Qry2['MAT_Descripcion']));
        gsMaterialesIds.Add(VarToStr(Qry2['MAT_ID']));
        gsMaterialesNumeros.Add(VarToStr(Qry2['MAT_Numero']));
        gsProvNumeroIds.Add(VarToStr(Qry2['MAT_ProvNumero']));
        gsKilos.Add(BoolToStrInt(StrToBool(VarToStr(Qry2['MAT_Kilos']))));
        gsDensidad.Add(VarToStr(Qry2['MAT_Densidad']));
        Qry2.Next;
    End;

    Qry2.Close;
end;

function TfrmEntradas.ValidateData():Boolean;
begin
  ValidateData := True;
  if txtPedimento.Text = '' Then
    begin
      MessageDlg('Por favor capture el pedimimento.', mtInformation,[mbOk], 0);
      result :=  False;
    end;

  if gvEntradas.RowCount <= 0 then
    begin
      MessageDlg('Por favor agrege el detalle de esta Entrada.', mtInformation,[mbOk], 0);
      result :=  False;
    end;

end;


procedure TfrmEntradas.txtTasaGeneralKeyPress(Sender: TObject;
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

procedure TfrmEntradas.txtIVAKeyPress(Sender: TObject; var Key: Char);
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

procedure TfrmEntradas.btnClearClick(Sender: TObject);
begin
  gvEntradas.ClearRows;
  txtSubtotal.Text := '';
  txtTIVA.Text := '';
  txtTotal.Text := '';
end;

procedure TfrmEntradas.btnDeleteClick(Sender: TObject);
begin
  gvEntradas.DeleteRow(gvEntradas.SelectedRow);
  if gvEntradas.RowCount <= 0 then
  begin
    txtSubtotal.Text := '';
    txtTIVA.Text := '';
    txtTotal.Text := '';
    Exit;
  end;

  calcularTotal();
end;

procedure TfrmEntradas.calcularTotal();
var sImporte: String;
dImporte, dIVA: Double;
i: Integer;
begin
  if gvEntradas.RowCount <= 0 then
     Exit;
       
  dImporte := 0.00;
  for i:= 0 to gvEntradas.RowCount - 1 do
  begin
          sImporte := StringReplace(gvEntradas.Cells[6,i],',','',[rfReplaceAll, rfIgnoreCase]);
          dImporte := dImporte + StrToFloat(sImporte);
  end;

  txtSubTotal.Text := FormatFloat('###,##0.00', dImporte);

  if txtIVA.Text = '' then
     txtIVA.Text := '10';

  dIVA := StrToFloat(txtIVA.Text);
  txtTIVA.Text := FormatFloat('###,##0.00', RoundTo(dImporte * (dIVA / 100),-2 ));

  txtTotal.Text := FormatFloat('###,##0.00', dImporte + ( RoundTo(dImporte * (dIVA / 100),-2 ) ));
end;


procedure TfrmEntradas.btnAddClick(Sender: TObject);
var i:Integer;
cantidad, longitud, diametro, densidad: double;
begin
 if gbMaterialValido = false then
  begin
    MessageDlg('El material ingresado no es valido.', mtInformation,[mbOk], 0);
    txtId.SetFocus;
    Exit;
  end;

  if gbKilos = False then begin
    if txtCantidad.Text = '' then
    begin
      MessageDlg('Por favor ingresa la cantidad del material.', mtInformation,[mbOk], 0);
      txtCantidad.SetFocus;
      Exit;
    end;
  end;

  if txtCosto.Text = '' then
  begin
    MessageDlg('Por favor ingresa la costo del material.', mtInformation,[mbOk], 0);
    txtCosto.SetFocus;
    Exit;
  end;

  for i:= 0 to gvEntradas.RowCount - 1 do
  begin
    if gvEntradas.Cells[3, i] = ddlMaterial.Text then
    begin
      MessageDlg('El material ' + ddlMaterial.Text + ' ya existe en esta Entrada.', mtInformation,[mbOk], 0);
      txtId.SetFocus;
      Exit;
    end;
  end;


  if gbKilos = True then begin
        if txtDiametro.Text = '' then begin
          MessageDlg('Por favor ingresa el diametro del Material.', mtInformation,[mbOk], 0);
          txtCosto.SetFocus;
          Exit;
        end;

        if txtLongitud.Text = '' then begin
          MessageDlg('Por favor ingresa la longitud del Material.', mtInformation,[mbOk], 0);
          txtCosto.SetFocus;
          Exit;
        end;
  end;

  gvEntradas.AddRow(1);
  gvEntradas.Cells[0,gvEntradas.RowCount -1] := '';
  gvEntradas.Cells[1,gvEntradas.RowCount -1] := txtId.Text;
  gvEntradas.Cells[2,gvEntradas.RowCount -1] := txtProvId.Text;

  gvEntradas.Cells[3,gvEntradas.RowCount -1] := ddlMaterial.Text;
  if gbKilos = False then  begin
    gvEntradas.Cells[4,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(txtCantidad.Text));
    cantidad := StrToFloat(txtCantidad.Text);
  end
  else begin
    longitud := StrToFloat(txtLongitud.Text);
    diametro := StrToFloat(txtDiametro.Text);
    densidad := StrToFloat(sDensidad);

    cantidad := 3.1416 * (diametro /2) * (diametro /2) * longitud * 12 * densidad;
    cantidad := cantidad / 2.22;
    gvEntradas.Cells[4,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', cantidad);
  end;
  gvEntradas.Cells[5,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(txtCosto.Text));
  gvEntradas.Cells[6,gvEntradas.RowCount -1] :=
          FormatFloat('#,###,##0.00', StrToFloat(txtCosto.Text) * cantidad );
  gvEntradas.Cells[7,gvEntradas.RowCount -1] := gsMaterialesIds[ddlMaterial.ItemIndex];

  ddlMaterial.ItemIndex := 0;
  txtId.Text := '';
  txtCantidad.Text := '';
  txtCosto.Text := '';
  lblMaterialDesc.Caption := '';
  txtProvId.Text := '';
  txtLongitud.Text := '';
  txtDiametro.Text := '';
  txtProvId.SetFocus;
  gbMaterialValido := False;
  calcularTotal();

  if gbKilos = True then begin
    txtLongitud.Visible := False;
    txtDiametro.Visible := False;
    lblLongitud.Visible := False;
    lblDiametro.Visible := False;

    txtLongitud.Enabled := False;
    txtDiametro.Enabled := False;
    lblPulgadas.Visible := False;
    lblPies.Visible := False;

    btnAdd.Left := btnAdd.Left - 300;
    btnDelete.Left := btnDelete.Left - 300;
    btnClear.Left := btnClear.Left - 300;
  end;
  gbKilos := False;
end;

procedure TfrmEntradas.BindDetalle(EntradaID: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
  SQLStr := 'SELECT E.ED_ID, E.MAT_ID, E.ED_Cantidad, E.ED_Costo, M.MAT_Descripcion, ' +
            'E.ED_Restante, M.MAT_Numero, M.MAT_ProvNumero ' +
            'FROM tblEntradasDetalle E ' +
            'INNER JOIN tblMateriales M ON E.MAT_ID = M.MAT_ID ' +
            'WHERE E.ENT_ID = ' + EntradaID;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvEntradas.ClearRows;
  While not Qry2.Eof do
  Begin
      gvEntradas.AddRow(1);
      gvEntradas.Cells[0,gvEntradas.RowCount -1] := VarToStr(Qry2['ED_ID']);
      gvEntradas.Cells[1,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_Numero']);
      gvEntradas.Cells[2,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_ProvNumero']);      
      gvEntradas.Cells[3,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_Descripcion']);
      gvEntradas.Cells[4,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(VarToStr(Qry2['ED_Cantidad'])) );
      gvEntradas.Cells[5,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(VarToStr(Qry2['ED_Costo'])) );
      gvEntradas.Cells[6,gvEntradas.RowCount -1] :=
         FormatFloat('#,###,##0.00', StrToFloat(VarToStr(Qry2['ED_Cantidad'])) * StrToFloat(VarToStr(Qry2['ED_Costo'])) );
      gvEntradas.Cells[7,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_ID']);
      gvEntradas.Cells[8,gvEntradas.RowCount -1] := VarToStr(Qry2['ED_Restante']);
      gvEntradas.Cells[9,gvEntradas.RowCount -1] := VarToStr(Qry2['ED_Cantidad']);
      Qry2.Next;
  End;

  Qry2.Close;
end;

procedure TfrmEntradas.ActualizarDetalle(EntradaID: String);
var i : Integer;
SQLStr : String;
sValor,sActual: String;
dValor,dActual,dDiff: Double;
begin
  if (giOpcion = 1) or (giOpcion = 3) then begin
      SQLStr := 'DELETE FROM tblEntradasDetalle WHERE ENT_ID = ' + EntradaID;
      conn.Execute(SQLStr);

      for i:= 0 to gvEntradas.RowCount - 1 do
      begin
            SQLStr := 'INSERT INTO tblEntradasDetalle(ENT_ID, MAT_ID, ED_Cantidad, ED_Restante, ED_Costo) ' +
                      'VALUES(' + EntradaID + ',' + gvEntradas.Cells[7,i] +
                      ',' + StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]) + ',' +
                      StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]) + ',' +
                      StringReplace(gvEntradas.Cells[5,i],',','',[rfReplaceAll, rfIgnoreCase]) + ')';

            conn.Execute(SQLStr);
      end;
  end
  else begin
      for i:= 0 to gvEntradas.RowCount - 1 do
      begin
            if gvEntradas.Cells[0,i] = '' then begin
                SQLStr := 'INSERT INTO tblEntradasDetalle(ENT_ID, MAT_ID, ED_Cantidad, ED_Restante, ED_Costo) ' +
                          'VALUES(' + EntradaID + ',' + gvEntradas.Cells[7,i] +
                          ',' + StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]) + ',' +
                          StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]) + ',' +
                          StringReplace(gvEntradas.Cells[5,i],',','',[rfReplaceAll, rfIgnoreCase]) + ')';
            end
            else begin
                sValor := StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]); //ED_Cantidad
                sActual := gvEntradas.Cells[8,i]; //ED_Restante

                dValor := StrToFloat(sValor);
                dActual := StrToFloat(sActual);

                dDiff := dValor - dActual;

                SQLStr := 'UPDATE tblEntradasDetalle SET ED_Cantidad = ' +
                          StringReplace(gvEntradas.Cells[4,i],',','',[rfReplaceAll, rfIgnoreCase]) +
                          ', ED_Costo = ' +
                          StringReplace(gvEntradas.Cells[5,i],',','',[rfReplaceAll, rfIgnoreCase]) +
                          ', ED_Restante = ED_Restante + ' + FloatToStr(dDiff) +
                          ' WHERE ED_ID = ' + gvEntradas.Cells[0,i];
            end;

            conn.Execute(SQLStr);
      end;
  end;

end;

procedure TfrmEntradas.gvEntradasAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: String; var Accept: Boolean);
var sValor,sRestante,sActual: String;

begin
  sValor := StringReplace(Value,',','',[rfReplaceAll, rfIgnoreCase]);
  sRestante := gvEntradas.Cells[7,ARow];
  sActual := gvEntradas.Cells[8,ARow];
{  if(sRestante <> sActual) then
  begin
          ShowMessage('No se puede cambiar este registro ya que el material ya salio del almacen.');
          Accept := False;
          Exit;
  end;
}

  if ((ACol = 2) and (not IsNumeric(sValor)) )then
  begin
          ShowMessage('La cantidad debe de ser numerica.');
          Accept := False;
          Exit;
  end;


  if ((ACol = 3) and (not IsNumeric(sValor)) )then
  begin
          ShowMessage('El Costo debe de ser numerico.');
          Accept := False;
          Exit;
  end;

  if ACol = 2 then begin
    gvEntradas.Cells[5,ARow] :=
            FormatFloat('#,###,##0.00', StrToFloat(gvEntradas.Cells[4,ARow]) * StrToFloat(sValor) );
  end;

  if ACol = 3 then begin
    gvEntradas.Cells[5,ARow] :=
            FormatFloat('#,###,##0.00', StrToFloat(gvEntradas.Cells[3,ARow]) * StrToFloat(sValor) );
  end;

  calcularTotal();


end;

procedure TfrmEntradas.ddlNacionalChange(Sender: TObject);
begin

  ddlTipoImp.ItemIndex := 0;
  ddlTipoImp.Enabled := True;
  if ddlNacional.Text = 'Nacional' then
  begin
        ddlTipoImp.ItemIndex := -1;
        ddlTipoImp.Enabled := False;
  end;
end;

procedure TfrmEntradas.txtIDExit(Sender: TObject);
var i: Integer;
begin
  if ( (txtId.ReadOnly = False) and (txtId.Text <> '') ) then begin
      i := getItemIndex(txtID.Text, gsMaterialesNumeros);
      if i = -1 then begin
        lblMaterialDesc.Caption := 'No se encontro ningun material con numero de identificacion : ' +
                                    txtID.Text;
        gbMaterialValido := false;
      end
      else begin
        ddlMaterial.ItemIndex := i;
        lblMaterialDesc.Caption := 'Descripcion del Material : ' + ddlMaterial.Text;
        txtProvId.Text := gsProvNumeroIds[i];
        //txtCantidad.SetFocus;
        gbMaterialValido := true;
        sDensidad := gsDensidad[i];
        displayKilos(i);
      end;
   end;
end;

procedure TfrmEntradas.txtIDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        if txtCantidad.Enabled = True then begin
          txtCantidad.SetFocus;
        end
        else begin
          txtCosto.SetFocus;
        end;

   end;
end;

procedure TfrmEntradas.txtProvIdKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var i: Integer;
begin
   If Key = vk_return then
   begin
    i := getItemIndex(txtProvId.Text, gsProvNumeroIds);
    if i = -1 then begin
      lblMaterialDesc.Caption := 'No se encontro ningun material con numero de identificacion : ' +
                                  txtID.Text;
      gbMaterialValido := false;
    end
    else begin
      ddlMaterial.ItemIndex := i;
      lblMaterialDesc.Caption := 'Descripcion del Material : ' + ddlMaterial.Text;
      txtId.Text := gsMaterialesNumeros[i];
      //txtCantidad.SetFocus;
      gbMaterialValido := true;
      sDensidad := gsDensidad[i];
      displayKilos(i);
    end;
   end;
end;

function TfrmEntradas.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;

procedure TfrmEntradas.btnOrdenClick(Sender: TObject);
begin
  gbMateriales.Visible := not gbMateriales.Visible;
  gbMateriales.Top := txtID.Top + 20;
  gbMateriales.Left := txtID.Left;
end;

procedure TfrmEntradas.tvMaterialesDblClick(Sender: TObject);
var selectedNode: TTreeNode;
i: integer;
begin
  gbMaterialValido := false;
  selectedNode := tvMateriales.Selected;
  i := ddlMaterial.Items.IndexOf(selectedNode.Text);
  if i = -1 then
    Exit;

  lblMaterialDesc.Caption := 'Descripcion del Material : ' + selectedNode.Text;
  txtID.Text := gsMaterialesNumeros[i];
  txtProvId.Text := gsProvNumeroIds[i];
  ddlMaterial.ItemIndex := i;
  sDensidad := gsDensidad[i];
  displayKilos(i);

  gbMateriales.Visible := not gbMateriales.Visible;
  gbMaterialValido := true;
end;

procedure TfrmEntradas.displayKilos(i: integer);
begin
    if '1' = gsKilos[i] then begin
      txtLongitud.Visible := True;
      txtDiametro.Visible := True;
      lblLongitud.Visible := True;
      lblDiametro.Visible := True;

      txtLongitud.Enabled := True;
      txtDiametro.Enabled := True;
      lblPulgadas.Visible := True;
      lblPies.Visible := True;
      if gbKilos = False then begin
        btnAdd.Left := btnAdd.Left + 300;
        btnDelete.Left := btnDelete.Left + 300;
        btnClear.Left := btnClear.Left + 300;
      end;
      gbKilos := True;
      txtCantidad.Enabled := False;
      txtCosto.SetFocus;
    end
    else begin
      txtLongitud.Visible := False;
      txtDiametro.Visible := False;
      lblLongitud.Visible := False;
      lblDiametro.Visible := False;

      txtLongitud.Enabled := False;
      txtDiametro.Enabled := False;
      lblPulgadas.Visible := False;
      lblPies.Visible := False;
      if gbKilos = True then begin
        btnAdd.Left := btnAdd.Left - 300;
        btnDelete.Left := btnDelete.Left - 300;
        btnClear.Left := btnClear.Left - 300;
      end;
      txtCantidad.Enabled := True;
      txtCantidad.SetFocus;
      gbKilos := False;
    end;
end;

end.
