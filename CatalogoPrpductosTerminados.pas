unit CatalogoPrpductosTerminados;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,Larco_Functions;

type
  TfrmProductosTerminados = class(TForm)
    gbButtons: TGroupBox;
    Label2: TLabel;
    lblId: TLabel;
    Label10: TLabel;
    Label4: TLabel;
    txtFraccion: TEdit;
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
    txtDescripcion: TEdit;
    txtUnidad: TEdit;
    Label1: TLabel;
    GroupBox2: TGroupBox;
    gvMateriales: TGridView;
    Label3: TLabel;
    cmbMateriales: TComboBox;
    Label5: TLabel;
    txtCantidad: TEdit;
    Agregar: TButton;
    BorrarMaterial: TButton;
    Limpiar: TButton;
    ddlUnidad: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindData(Detail:Boolean);
    procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure PrimeroClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure btnCancelarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure BindMateriales();
    procedure BindUnidades();
    procedure BindDetails(productID: String);
    procedure AgregarClick(Sender: TObject);
    procedure LimpiarClick(Sender: TObject);
    procedure BorrarMaterialClick(Sender: TObject);
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure ActualizarDetalle(productID: String);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProductosTerminados: TfrmProductosTerminados;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  gsMaterialesIds,gsUnidadesIds : TStringList;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmProductosTerminados.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmProductosTerminados.FormCreate(Sender: TObject);
begin
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  gsMaterialesIds := TStringList.Create;
  gsUnidadesIds := TStringList.Create;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblProductosTerminados ORDER BY PT_ID';
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
        BindData(true);
  end;
  BindUnidades();
  BindMateriales();

  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmProductosTerminados.BindData(Detail:Boolean);
begin
  if Qry.RecordCount <= 0 Then begin
        ClearData();
        Exit;
  end;
  lblId.Caption  := VarToStr(Qry['PT_ID']);
  txtFraccion.Text := VarToStr(Qry['PT_Fraccion']);
  txtDescripcion.Text := VarToStr(Qry['PT_Descripcion']);
  txtUnidad.Text := VarToStr(Qry['PT_Unidad']);
  ddlUnidad.ItemIndex := getItemIndex(VarToStr(Qry['PT_Unidad']), gsUnidadesIds);

  if Detail =  true then BindDetails(lblId.Caption);


end;

procedure TfrmProductosTerminados.ClearData();
begin
  txtFraccion.Text := '';
  txtDescripcion.Text := '';
  txtUnidad.Text := '';
  ddlUnidad.ItemIndex := 0;

  gvMateriales.ClearRows;
  txtCantidad.Text := '';
end;

procedure TfrmProductosTerminados.EnableControls(Value:Boolean);
begin
  txtFraccion.ReadOnly := Value;
  txtDescripcion.ReadOnly := Value;
  txtUnidad.ReadOnly := Value;
  ddlUnidad.Enabled := not Value;

  txtCantidad.ReadOnly := Value;
  cmbMateriales.Enabled := Not Value;
  Agregar.Enabled := Not Value;
  Limpiar.Enabled := Not Value;
  BorrarMaterial.Enabled := Not Value;
end;


procedure TfrmProductosTerminados.PrimeroClick(Sender: TObject);
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


  BindData(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

function TfrmProductosTerminados.ValidateData():Boolean;
begin
        Result := False;
        if gvMateriales.RowCount <= 0 then
        begin
                ShowMessage('Por favor agrege los materiales que conforman este producto terminado.');
                Exit;
        end;


        Result :=  True;
end;

procedure TfrmProductosTerminados.btnCancelarClick(Sender: TObject);
begin
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;

  BindData(True);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmProductosTerminados.NuevoClick(Sender: TObject);
begin
  giOpcion := 1;
  ClearData();
  EnableControls(False);
  txtFraccion.SetFocus;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmProductosTerminados.EditarClick(Sender: TObject);
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

procedure TfrmProductosTerminados.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmProductosTerminados.BuscarClick(Sender: TObject);
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

procedure TfrmProductosTerminados.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['PT_Fraccion'] := txtFraccion.Text;
        Qry['PT_Descripcion'] := txtDescripcion.Text;
        //Qry['PT_Unidad'] := txtUnidad.Text;
        Qry['PT_Unidad'] := gsUnidadesIds[ddlUnidad.ItemIndex];
        Qry.Post;

        BindData(false);
        ActualizarDetalle(lblId.Caption);
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['PT_Fraccion'] := txtFraccion.Text;
        Qry['PT_Descripcion'] := txtDescripcion.Text;
        //Qry['PT_Unidad'] := txtUnidad.Text;
        Qry['PT_Unidad'] := gsUnidadesIds[ddlUnidad.ItemIndex];
        Qry.Post;

        BindData(false);
        ActualizarDetalle(lblId.Caption);
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar el Producto Terminado : ' +
                      txtDescripcion.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar el Producto Terminado : ' +
                            txtDescripcion.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        Qry.Delete;
                        gvMateriales.ClearRows;
                        ActualizarDetalle(lblId.Caption);
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtFraccion.Text <> '' then
        begin
              if not Qry.Locate('PT_Fraccion',txtFraccion.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Producto Terminado con Fraccion Arancelaria ' + txtFraccion.Text + '.', mtInformation,[mbOk], 0);
                    txtFraccion.SetFocus;
                    Exit;
                end;
        end
        else if txtDescripcion.Text <> '' then
        begin
              if not Qry.Locate('PT_Descripcion',txtDescripcion.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Producto Terminado con Descripcion ' + txtDescripcion.Text + '.', mtInformation,[mbOk], 0);
                    txtDescripcion.SetFocus;
                    Exit;
                end;
        end
        else if txtUnidad.Text <> '' then
        begin
              if not Qry.Locate('PT_Unidad',txtUnidad.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Producto Terminado con Tipo de Unidad ' + txtUnidad.Text + '.', mtInformation,[mbOk], 0);
                    txtUnidad.SetFocus;
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
  BindData(True);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmProductosTerminados.BindMateriales();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT MAT_ID,MAT_Descripcion FROM tblMateriales Order By MAT_ID';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbMateriales.Items.Clear;
    gsMaterialesIds.Clear;
    While not Qry2.Eof do
    Begin
        cmbMateriales.Items.Add(VarToStr(Qry2['MAT_Descripcion']));
        gsMaterialesIds.Add(VarToStr(Qry2['MAT_ID']));
        Qry2.Next;
    End;

    cmbMateriales.Text := cmbMateriales.Items[0];

    Qry2.Close;

end;

procedure TfrmProductosTerminados.BindUnidades();
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

procedure TfrmProductosTerminados.AgregarClick(Sender: TObject);
begin
  if cmbMateriales.Text = '' then begin
        ShowMessage('Especifique el Material que se va a agregar.');
        Exit;
  end;

  if txtCantidad.Text = '' then begin
        ShowMessage('Especifique la cantidad.');
        Exit;
  end;

  gvMateriales.AddRow(1);
  gvMateriales.Cells[0,gvMateriales.RowCount -1] := '';
  gvMateriales.Cells[1,gvMateriales.RowCount -1] := gsMaterialesIds[cmbMateriales.Items.IndexOf(cmbMateriales.Text)];
  gvMateriales.Cells[2,gvMateriales.RowCount -1] := cmbMateriales.Text;
  gvMateriales.Cells[3,gvMateriales.RowCount -1] := txtCantidad.Text;

  txtCantidad.Text := '';
  cmbMateriales.SetFocus;
end;

procedure TfrmProductosTerminados.LimpiarClick(Sender: TObject);
begin
  gvMateriales.ClearRows;
end;

procedure TfrmProductosTerminados.BorrarMaterialClick(Sender: TObject);
begin
  gvMateriales.DeleteRow(gvMateriales.SelectedRow);
end;

procedure TfrmProductosTerminados.txtCantidadKeyPress(Sender: TObject;
  var Key: Char);
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

procedure TfrmProductosTerminados.ActualizarDetalle(productID: String);
var i : Integer;
SQLStr : String;
begin
  SQLStr := 'DELETE FROM tblproductosTerminadosDetalle WHERE PT_ID = ' + productID;
  conn.Execute(SQLStr);

  for i:= 0 to gvMateriales.RowCount - 1 do
  begin
        SQLStr := 'INSERT INTO tblproductosTerminadosDetalle(PT_ID, MAT_ID, PTD_Cantidad) ' +
                  'VALUES(' + productID + ',' + gvMateriales.Cells[1,i] +
                  ',' + gvMateriales.Cells[3,i] + ')';

        conn.Execute(SQLStr);
  end;

end;

procedure TfrmProductosTerminados.BindDetails(productID: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
  SQLStr := 'SELECT P.PTD_ID, P.PT_ID, P.MAT_ID, P.PTD_Cantidad, M.MAT_Descripcion ' +
            'FROM tblproductosTerminadosDetalle P ' +
            'INNER JOIN tblMateriales M ON P.MAT_ID = M.MAT_ID ' +
            'WHERE PT_ID = ' + productID;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvMateriales.ClearRows;
    While not Qry2.Eof do
    Begin
        gvMateriales.AddRow(1);
        gvMateriales.Cells[0,gvMateriales.RowCount -1] := VarToStr(Qry2['PTD_ID']);
        gvMateriales.Cells[1,gvMateriales.RowCount -1] := VarToStr(Qry2['MAT_ID']);
        gvMateriales.Cells[2,gvMateriales.RowCount -1] := VarToStr(Qry2['MAT_Descripcion']);
        gvMateriales.Cells[3,gvMateriales.RowCount -1] := VarToStr(Qry2['PTD_Cantidad']);
        Qry2.Next;
    End;

    Qry2.Close;
end;


end.
