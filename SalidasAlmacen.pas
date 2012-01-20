unit SalidasAlmacen;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math, Mask;

type
  TfrmSalidasAlmacen = class(TForm)
    GroupBox2: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    gvEntradas: TGridView;
    ddlMaterial: TComboBox;
    txtCantidad: TEdit;
    btnDelete: TButton;
    btnClear: TButton;
    gbButtons: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    lblId: TLabel;
    lblAnio: TLabel;
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
    deFecha: TDateEditor;
    ddlEmpleado: TComboBox;
    txtOrden: TMaskEdit;
    btnAdd: TButton;
    txtEmpleado: TEdit;
    Label1: TLabel;
    lblEmpleado: TLabel;
    txtID: TEdit;
    lblMaterialDesc: TLabel;
    btnOrden: TButton;
    gbMateriales: TGroupBox;
    tvMateriales: TTreeView;
    Label12: TLabel;
    txtProvId: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure EnableButtons();
    Procedure BindData();
    Procedure ClearData();
    procedure BindEmpleados();
    procedure BindMateriales();
    procedure BindDetalle(EntradaID: String);    
    procedure SendTab(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btnClearClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure PrimeroClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure gvEntradasAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: String; var Accept: Boolean);
    procedure ActualizarDetalle(SalidaID: String);
    procedure txtEmpleadoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtEmpleadoExit(Sender: TObject);
    procedure txtIDExit(Sender: TObject);
    procedure txtIDKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnOrdenClick(Sender: TObject);
    procedure tvMaterialesDblClick(Sender: TObject);
    procedure txtProvIdKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSalidasAlmacen: TfrmSalidasAlmacen;
  giOpcion : Integer;
  gsEmpleadosIds, gsMaterialesIds, gsMaterialesNumeros,gsProvNumeroIds: TStringList;
  Conn : TADOConnection;
  Qry : TADOQuery;
  gbMaterialValido: boolean;
  gsYear : String;
  sPermits : String;
implementation

uses Main, Login;

{$R *.dfm}

procedure TfrmSalidasAlmacen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmSalidasAlmacen.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);
  gsYear := RightStr(lblAnio.Caption,2) + '-';
  lblMaterialDesc.Caption := '';
  
  gsEmpleadosIds := TStringList.Create;
  gsMaterialesIds := TStringList.Create;
  gsMaterialesNumeros := TStringList.Create;
  gsProvNumeroIds := TStringList.Create;
  gbMaterialValido := false;

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblSalidas WHERE YEAR(SAL_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' ORDER BY SAL_ID';
  Qry.Open;

  BindEmpleados();
  BindMateriales();
  ClearData();

  EnableButtons();
  BindData();
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmSalidasAlmacen.EnableControls(Value:Boolean);
begin
  txtOrden.ReadOnly := Value;
  deFecha.Enabled := not Value;
  ddlEmpleado.Enabled := not Value;
  txtEmpleado.ReadOnly := Value;

  ddlMaterial.Enabled := not Value;
  txtCantidad.ReadOnly := Value;
  txtID.ReadOnly := Value;
  btnAdd.Enabled := not Value;
  btnDelete.Enabled := not Value;
  btnClear.Enabled :=  not Value;
  gvEntradas.Enabled := not Value;
  btnOrden.Enabled := not Value;
  txtProvId.ReadOnly := Value;
end;

procedure TfrmSalidasAlmacen.EnableButtons();
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


procedure TfrmSalidasAlmacen.ClearData();
begin
  txtOrden.Text := '';
  deFecha.Text := DateToStr(Now);
  ddlEmpleado.ItemIndex := 0;
  txtEmpleado.Text := '';

  ddlMaterial.ItemIndex := 0;
  txtID.Text := '';
  txtCantidad.Text := '';
  btnClearClick(nil);
  gbMateriales.Visible := False;
  txtProvId.Text := '';
end;

procedure TfrmSalidasAlmacen.BindData();
begin
  if Qry.RecordCount = 0 then
          Exit;

  lblId.Caption  := VarToStr(Qry['SAL_ID']);
  txtOrden.Text := RightStr(VarToStr(Qry['SAL_Orden']), 10);
  deFecha.Text := VarToStr(Qry['SAL_Fecha']);
  ddlEmpleado.ItemIndex := getItemIndex(VarToStr(Qry['SAL_Solicitado']), gsEmpleadosIds);
  lblEmpleado.Caption := ddlEmpleado.Text;
  txtEmpleado.Text := gsEmpleadosIds[ddlEmpleado.ItemIndex];

  BindDetalle(lblId.Caption);
end;

procedure TfrmSalidasAlmacen.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;
end;


procedure TfrmSalidasAlmacen.btnClearClick(Sender: TObject);
begin
  gvEntradas.ClearRows;
end;

procedure TfrmSalidasAlmacen.BindEmpleados();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT ID, Nombre FROM tblEmpleados ORDER BY Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    ddlEmpleado.Items.Clear;
    gsEmpleadosIds.Clear;
    While not Qry2.Eof do
    Begin
        ddlEmpleado.Items.Add(Qry2['Nombre']);
        gsEmpleadosIds.Add(VarToStr(Qry2['ID']));
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmSalidasAlmacen.BindMateriales();
var Qry2 : TADOQuery;
SQLStr, tipo : String;
treeNode : TTreeNode;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT M.MAT_ID, M.MAT_Descripcion, M.MAT_Numero, T.TIP_Descripcion, M.MAT_ProvNumero  ' +
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
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmSalidasAlmacen.BindDetalle(EntradaID: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
  SQLStr := 'SELECT S.SD_ID, S.MAT_ID, S.SD_Cantidad, M.MAT_Descripcion ' +
            'FROM tblSalidasDetalle S ' +
            'INNER JOIN tblMateriales M ON S.MAT_ID = M.MAT_ID ' +
            'WHERE S.SAL_ID = ' + EntradaID;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvEntradas.ClearRows;
  While not Qry2.Eof do
  Begin
      gvEntradas.AddRow(1);
      gvEntradas.Cells[0,gvEntradas.RowCount -1] := VarToStr(Qry2['SD_ID']);
      gvEntradas.Cells[1,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_Descripcion']);
      gvEntradas.Cells[2,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(VarToStr(Qry2['SD_Cantidad'])) );
      gvEntradas.Cells[3,gvEntradas.RowCount -1] := VarToStr(Qry2['MAT_ID']);
      Qry2.Next;
  End;

  Qry2.Close;
end;


procedure TfrmSalidasAlmacen.btnAddClick(Sender: TObject);
var i:Integer;
Qry2 : TADOQuery;
SQLStr : String;
cantidad : Double;
begin
  if gbMaterialValido = false then
  begin
    MessageDlg('El material ingresado no es valido.', mtInformation,[mbOk], 0);
    txtId.SetFocus;
    Exit;
  end;

  if txtCantidad.Text = '' then
  begin
    MessageDlg('Por favor ingresa la cantidad del material.', mtInformation,[mbOk], 0);
    txtCantidad.SetFocus;
    Exit;
  end;

  for i:= 0 to gvEntradas.RowCount - 1 do
  begin
    if gvEntradas.Cells[1, i] = ddlMaterial.Text then
    begin
      MessageDlg('El material ' + ddlMaterial.Text + ' ya existe en esta Salida.', mtInformation,[mbOk], 0);
      txtId.SetFocus;
      Exit;
    end;
  end;

  SQLStr := 'SELECT MAT_Cantidad FROM tblMateriales WHERE MAT_ID = ' +
            gsMaterialesIds[ddlMaterial.ItemIndex];

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  cantidad := StrToFloat(VarToStr(Qry2['MAT_Cantidad']));
  if StrToFloat(txtCantidad.Text) <= 0.0 then
  begin
      MessageDlg('La cantidad no puede ser igual o menor a cero.', mtInformation,[mbOk], 0);
      Exit;
  end;


  if (cantidad < StrToFloat(txtCantidad.Text)) then
  begin
      MessageDlg('La cantidad excede la existencia en almacen.', mtInformation,[mbOk], 0);
      Exit;
  end;

  gvEntradas.AddRow(1);
  gvEntradas.Cells[0,gvEntradas.RowCount -1] := '';
  gvEntradas.Cells[1,gvEntradas.RowCount -1] := ddlMaterial.Text;
  gvEntradas.Cells[2,gvEntradas.RowCount -1] := FormatFloat('#,###,##0.00', StrToFloat(txtCantidad.Text));
  gvEntradas.Cells[3,gvEntradas.RowCount -1] := gsMaterialesIds[ddlMaterial.ItemIndex];

  ddlMaterial.ItemIndex := 0;
  txtId.Text := '';
  txtCantidad.Text := '';
  lblMaterialDesc.Caption := '';
  txtProvId.Text := '';
  txtProvId.SetFocus;
  gbMaterialValido := False;
end;

procedure TfrmSalidasAlmacen.PrimeroClick(Sender: TObject);
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

procedure TfrmSalidasAlmacen.NuevoClick(Sender: TObject);
begin
  ClearData();
  EnableControls(False);
  txtOrden.SetFocus;
  giOpcion := 1;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
end;

procedure TfrmSalidasAlmacen.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
  txtOrden.SetFocus;
end;

procedure TfrmSalidasAlmacen.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
end;

procedure TfrmSalidasAlmacen.btnCancelarClick(Sender: TObject);
begin
  ClearData();
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  EnableButtons();
  BindData();
end;

procedure TfrmSalidasAlmacen.btnDeleteClick(Sender: TObject);
begin
  gvEntradas.DeleteRow(gvEntradas.SelectedRow);
end;

procedure TfrmSalidasAlmacen.btnAceptarClick(Sender: TObject);
var user: String;
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['SAL_Orden'] := gsYear + txtOrden.Text;
        Qry['SAL_Fecha'] := deFecha.Text;
        Qry['SAL_Solicitado'] := txtEmpleado.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry.Post;

        ActualizarDetalle(Qry['SAL_ID']);
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['SAL_Orden'] := gsYear + txtOrden.Text;
        Qry['SAL_Fecha'] := deFecha.Text;
        Qry['SAL_Solicitado'] := txtEmpleado.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry.Post;

        ActualizarDetalle(Qry['SAL_ID']);
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar esta Salida?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro esta Salida?',
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
        if txtOrden.Text <> '' then
        begin
              if not Qry.Locate('SAL_Orden',txtOrden.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Salida para esta Orden de Trabajo : ' + txtOrden.Text + '.', mtInformation,[mbOk], 0);
                    txtOrden.SetFocus;
                    Exit;
                end;
        end
        else if deFecha.Text <> '' then
        begin
              if not Qry.Locate('SAL_Fecha',deFecha.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Salida con Fecha : ' + deFecha.Text + '.', mtInformation,[mbOk], 0);
                    deFecha.SetFocus;
                    Exit;
                end;
        end
        else if ddlEmpleado.Text <> '' then
        begin
              if not Qry.Locate('SAL_Solicitado',gsEmpleadosIds[ddlEmpleado.ItemIndex] ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Salida solicitada por : ' + ddlEmpleado.Text + '.', mtInformation,[mbOk], 0);
                    ddlEmpleado.SetFocus;
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

function TfrmSalidasAlmacen.ValidateData():Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
  result := True;
  if txtOrden.Text = '   -   -  ' Then
    begin
      MessageDlg('Por favor ingrese el numero de Orden de Trabajo.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;

  if lblEmpleado.Caption = '' Then
    begin
      MessageDlg('Por favor ingrese el numero de empleado en "Solicitado por".', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;


  if lblEmpleado.Caption = 'Empleado Invalido.' Then
    begin
      MessageDlg('El numero de empleado ingresado en "Solicitado por" no es valido.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;


  SQLStr := 'SELECT ITE_Nombre FROM tblOrdenes WHERE ITE_Nombre = ' +
            QuotedStr(RightStr(lblAnio.Caption,2) + '-' + txtOrden.Text);

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount <= 0 then
    begin
      MessageDlg('El numero de Orden de Trabajo es incorrecto.', mtInformation,[mbOk], 0);
      result :=  False;
    end
  else begin
      SQLStr := 'SELECT * FROM tblItemTasks WHERE ITE_Nombre = ' +
                QuotedStr(RightStr(lblAnio.Caption,2) + '-' + txtOrden.Text) +
                ' AND TAS_ID = 19 AND ITS_Status = 2';

      Qry2.SQL.Clear;
      Qry2.SQL.Text := SQLStr;
      Qry2.Open;

      if Qry2.RecordCount > 0 then
        begin
          MessageDlg('La Orden de Trabajo ya salio de Ventas Final.', mtInformation,[mbOk], 0);
          result :=  False;
        end;

  end


end;


procedure TfrmSalidasAlmacen.txtCantidadKeyPress(Sender: TObject;
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

procedure TfrmSalidasAlmacen.gvEntradasAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: String; var Accept: Boolean);
var Qry2 : TADOQuery;
SQLStr : String;
sValor: String;
cantidad : Double;
begin
  sValor := StringReplace(Value,',','',[rfReplaceAll, rfIgnoreCase]);

  if ((ACol = 2) and (not IsNumeric(sValor)) )then
  begin
          ShowMessage('La cantidad debe de ser numerica.');
          Accept := False;
          Exit;
  end;

  SQLStr := 'SELECT MAT_Cantidad FROM tblMateriales WHERE MAT_ID = ' +
            gsMaterialesIds[ddlMaterial.ItemIndex];

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  cantidad := StrToFloat(VarToStr(Qry2['MAT_Cantidad']));
  if (cantidad < StrToFloat(txtCantidad.Text)) then
  begin
      MessageDlg('La cantidad excede la existencia en almacen.', mtInformation,[mbOk], 0);
      Accept := False;
      Exit;
  end;



end;

procedure TfrmSalidasAlmacen.ActualizarDetalle(SalidaID: String);
var i : Integer;
SQLStr : String;                                                                                         
begin
  if (giOpcion = 1) or (giOpcion = 3) then begin
      SQLStr := 'DELETE FROM tblSalidasDetalle WHERE SAL_ID = ' + SalidaID;
      conn.Execute(SQLStr);

      for i:= 0 to gvEntradas.RowCount - 1 do
      begin
            SQLStr := 'INSERT INTO tblSalidasDetalle(SAL_ID, MAT_ID, SD_Cantidad) VALUES(' +
                      SalidaID + ',' + gvEntradas.Cells[3,i] + ',' +
                      StringReplace(gvEntradas.Cells[2,i],',','',[rfReplaceAll, rfIgnoreCase]) + ')';

            conn.Execute(SQLStr);
      end;
  end
  else begin
      for i:= 0 to gvEntradas.RowCount - 1 do
      begin
            if gvEntradas.Cells[0,i] = '' then begin
                SQLStr := 'INSERT INTO tblSalidasDetalle(SAL_ID, MAT_ID, SD_Cantidad) VALUES(' +
                          SalidaID + ',' + gvEntradas.Cells[3,i] + ',' +
                          StringReplace(gvEntradas.Cells[2,i],',','',[rfReplaceAll, rfIgnoreCase]) + ')';
            end
            else begin
                SQLStr := 'UPDATE tblSalidasDetalle SET SD_Cantidad = ' +
                          StringReplace(gvEntradas.Cells[2,i],',','',[rfReplaceAll, rfIgnoreCase]) +
                          ', IS_SL = 0 WHERE SD_ID = ' + gvEntradas.Cells[0,i];
            end;

            conn.Execute(SQLStr);
      end;
  end;

end;


procedure TfrmSalidasAlmacen.txtEmpleadoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var i:Integer;
begin
  if key = vk_Return then
  begin
      i := getItemIndex(txtEmpleado.Text, gsEmpleadosIds);
      if i = -1 then begin
        lblEmpleado.Caption := 'Empleado Invalido.';
      end
      else begin
        ddlEmpleado.ItemIndex := i;
        lblEmpleado.Caption := ddlEmpleado.Text;
      end;

  end

end;

procedure TfrmSalidasAlmacen.txtEmpleadoExit(Sender: TObject);
var i:Integer;
begin
      i := getItemIndex(txtEmpleado.Text, gsEmpleadosIds);
      if i = -1 then begin
        lblEmpleado.Caption := 'Empleado Invalido.';
      end
      else begin
        ddlEmpleado.ItemIndex := i;
        lblEmpleado.Caption := ddlEmpleado.Text;
      end;
end;

procedure TfrmSalidasAlmacen.txtIDExit(Sender: TObject);
var i: Integer;
begin
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
  end;
end;

procedure TfrmSalidasAlmacen.txtProvIdKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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
      txtCantidad.SetFocus;
      gbMaterialValido := true;
    end;
   end;
end;

procedure TfrmSalidasAlmacen.txtIDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        txtCantidad.SetFocus;
   end;
end;

procedure TfrmSalidasAlmacen.btnOrdenClick(Sender: TObject);
begin
  gbMateriales.Visible := not gbMateriales.Visible;
  gbMateriales.Top := txtID.Top + 20;
  gbMateriales.Left := txtID.Left;  
end;

procedure TfrmSalidasAlmacen.tvMaterialesDblClick(Sender: TObject);
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

  gbMateriales.Visible := not gbMateriales.Visible;
  txtCantidad.SetFocus;
  gbMaterialValido := true;
end;

end.
