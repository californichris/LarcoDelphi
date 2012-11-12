unit Ventas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors,ImpresionOrden,Larco_Functions, ExtCtrls, Menus,Clipbrd;

type
  TfrmVentas = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label9: TLabel;
    Label14: TLabel;
    lblId: TLabel;
    txtProceso: TEdit;
    txtRequerida: TEdit;
    cmbProductos: TComboBox;
    txtNumero: TEdit;
    txtTerminal: TEdit;
    txtRecibido: TEdit;
    cmbEmpleados: TComboBox;
    chkAprobacion: TCheckBox;
    txtUnitario: TEdit;
    txtObservaciones: TMemo;
    txtOtras: TMemo;
    txtOrden: TMaskEdit;
    txtOrdenada: TEdit;
    txtTotal: TEdit;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    TabSheet2: TTabSheet;
    gvVentas: TGridView;
    btnExport: TButton;
    SaveDialog1: TSaveDialog;
    Label15: TLabel;
    deInterna: TDateEditor;
    deEntrega: TDateEditor;
    Label16: TLabel;
    txtCompra: TEdit;
    chkDlls: TCheckBox;
    Button5: TButton;
    lblStock: TLabel;
    lblAnio: TLabel;
    chkStock: TCheckBox;
    chkPlano: TCheckBox;
    cmbPlanos: TComboBox;
    Panel1: TPanel;
    Primero: TButton;
    Anterior: TButton;
    Ultimo: TButton;
    Siguiente: TButton;
    GroupBox1: TGroupBox;
    Label22: TLabel;
    gvNumParte: TGridView;
    Label23: TLabel;
    gvNumPlano: TGridView;
    Label19: TLabel;
    gvProdNumero: TGridView;
    Label20: TLabel;
    gvProdPlano: TGridView;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    gvStatus: TGridView;
    chkStockParcial: TCheckBox;
    txtStockParcial: TEdit;
    chkMezclar: TCheckBox;
    gvMezclado: TGridView;
    txtOrdenMezclar: TMaskEdit;
    txtCantidadMezclar: TEdit;
    DeleteOrden: TButton;
    AddOrden: TButton;
    lblStockParcial: TLabel;
    lblRequerida: TLabel;
    PopupMenu1: TPopupMenu;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    PopupMenu2: TPopupMenu;
    Mezclar1: TMenuItem;
    Copiar1: TMenuItem;
    Label17: TLabel;
    txtRequisicion: TEdit;
    function FormIsRunning(FormName: String):Boolean;
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindProductos();
    procedure BindPlanos();
    procedure BindEmpleados();
    Procedure BindOrden();
    procedure FormCreate(Sender: TObject);
    procedure SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
    Procedure ClearData();
    procedure BorrarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function  ValidateData():Boolean;
    function  ValidateCliente(Clave:String):Boolean;
    function BoolToStrInt(Value:Boolean):String;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure txtRequeridaKeyPress(Sender: TObject; var Key: Char);
    procedure txtUnitarioKeyPress(Sender: TObject; var Key: Char);
    procedure txtUnitarioChange(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure cmbProductosDropDown(Sender: TObject);
    procedure txtRequeridaExit(Sender: TObject);
    procedure cmbProductosChange(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure txtOrdenChange(Sender: TObject);
    procedure deInternaChange(Sender: TObject);
    procedure deInternaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure deEntregaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cmbEmpleadosDropDown(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chkPlanoClick(Sender: TObject);
    procedure cmbPlanosDropDown(Sender: TObject);
    procedure cmbPlanosChange(Sender: TObject);
    procedure PrimeroClick(Sender: TObject);
    procedure EnableButtons();
    procedure txtNumeroExit(Sender: TObject);
    procedure txtOrdenKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chkStockParcialClick(Sender: TObject);
    procedure chkMezclarClick(Sender: TObject);
    procedure DeleteOrdenClick(Sender: TObject);
    procedure txtCantidadMezclarKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure AddOrdenClick(Sender: TObject);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure Mezclar1Click(Sender: TObject);
    procedure InsertOrdenesMezclar();
    procedure BindOrdenesMezclar();
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmVentas: TfrmVentas;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  gsFirstTask,gsYear,gsOYear : String;
  sPermits : String;
implementation

{$R *.dfm}
uses Main, Editor;

procedure TfrmVentas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmVentas.BindPlanos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT PN_Id, PN_Numero FROM tblPlano ORDER BY PN_Numero';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbPlanos.Items.Clear;
    While not Qry2.Eof do
    Begin
        cmbPlanos.AddItem(VarToStr(Qry2['PN_Numero']), createValue(VarToStr(Qry2['PN_Id'])));
        Qry2.Next;
    End;

    cmbPlanos.Text := '';
    Qry2.Close;
end;

procedure TfrmVentas.BindProductos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblProductos ORDER BY Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbProductos.Items.Clear;
    While not Qry2.Eof do
    Begin
        cmbProductos.Items.Add(Qry2['Nombre']);
        Qry2.Next;
    End;

    cmbProductos.Text := '';
    Qry2.Close;
end;

procedure TfrmVentas.BindEmpleados();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblEmpleados WHERE Departamento IN (' + QuotedStr('Ventas') +
              ',' + QuotedStr('Administracion') + ')';

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


procedure TfrmVentas.FormCreate(Sender: TObject);
begin
    gsFirstTask := 'Ventas';

    lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);
    gsOYear := RightStr(lblAnio.Caption,2);
    gsYear := gsOYear + '-';

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    //Qry.SQL.Text := 'SELECT * FROM tblOrdenes WHERE Left(ITE_Nombre,2) = ' +
    //                QuotedStr(gsOYear) + ' ORDER BY ITE_ID';
    Qry.SQL.Text := 'Traer_Ordenes ' + QuotedStr(gsOYear);
    Qry.Open;

    BindProductos();
    BindEmpleados();
    BindPlanos();

    if Qry.RecordCount > 0 then
        BindOrden()
    else
    begin
        Editar.Enabled := False;
        Borrar.Enabled := False;
        Buscar.Enabled := False;
    end;

    sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
    EnableFormButtons(gbButtons, sPermits);

    giOpcion := 0;
end;

procedure TfrmVentas.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end
   else if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
    begin
            btnCancelarClick(nil);
    end;
end;

procedure TfrmVentas.ClearData();
begin
    txtOrden.Text := '';
    txtProceso.Text := '';
    txtRequerida.Text := '';
    lblRequerida.Caption := '';
    txtNumero.Text := '';
    txtTerminal.Text := ''; // El label fue cambiado por Revision el 1 oct 2012 a peticion
                            // de daria para que apareciera en el export del ROC
    txtRequisicion.Text := '';                             
    deEntrega.Text := DateToStr(Now);
    deInterna.Text := DateToStr(Now);
    txtRecibido.Text := DateToStr(Now);
    txtUnitario.Text := '';
    txtObservaciones.Text := '';
    txtOtras.Text := '';
    txtOrdenada.Text := '';
    txtTotal.Text := '';
    cmbEmpleados.Text := '';
    cmbProductos.Text := '';
    chkAprobacion.Checked := False;
    chkDlls.Checked := False;
    txtCompra.Text := '';
    chkStock.Checked := False;
    chkPlano.Checked := False;
    cmbPlanos.Text := '';

    chkStockParcial.Checked := false;
    txtStockParcial.Text := '';
    lblStockParcial.Caption := '';

    chkMezclar.Checked := False;
    txtOrdenMezclar.Text := '';
    txtCantidadMezclar.Text := '';
    gvMezclado.ClearRows;

    gvNumParte.ClearRows;
    gvNumPlano.ClearRows;
    gvProdNumero.ClearRows;
    gvProdPlano.ClearRows;
    gvStatus.ClearRows;
end;

procedure TfrmVentas.NuevoClick(Sender: TObject);
begin
giOpcion := 1;
ClearData();
EnableControls(False);
EnableButtons();

deEntrega.Text := DateToStr(Now);
deInterna.Text := DateToStr(Now);
txtRecibido.Text := DateToStr(Now);
txtUnitario.Text := '0';
txtTotal.Text := '0';

txtOrden.SetFocus;
end;

procedure TfrmVentas.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);
EnableButtons();
txtOrden.SetFocus;
end;

procedure TfrmVentas.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
EnableButtons();
end;

procedure TfrmVentas.BuscarClick(Sender: TObject);
begin
giOpcion := 4;
ClearData();
EnableButtons();

txtOrden.ReadOnly := False;
txtOrden.SetFocus;
end;

procedure TfrmVentas.EnableControls(Value:Boolean);
begin
    txtOrden.ReadOnly := Value;
    txtProceso.ReadOnly := Value;
    txtRequerida.ReadOnly := Value;
    txtNumero.ReadOnly := Value;
    txtTerminal.ReadOnly := Value;
    txtRequisicion.ReadOnly := Value;
    txtUnitario.ReadOnly := Value;
    txtObservaciones.ReadOnly := Value;
    txtOtras.ReadOnly := Value;
    txtOrdenada.ReadOnly := Value;
    txtCompra.ReadOnly := Value;

    deEntrega.Enabled := not Value;
    deInterna.Enabled := not Value;
    cmbProductos.Enabled := not Value;
    cmbEmpleados.Enabled := not Value;
    chkAprobacion.Enabled := not Value;
    chkDlls.Enabled := not Value;
    chkStock.Enabled := not Value;
    chkPlano.Enabled := not Value;
    if giOpcion = 0 then begin
      cmbPlanos.Enabled := not Value;
    end else begin
      cmbPlanos.Enabled := chkPlano.Checked;
    end;

    chkStockParcial.Enabled := not Value;
    if giOpcion = 0 then begin
      txtStockParcial.Enabled := not Value;
    end else begin
      txtStockParcial.Enabled := chkStockParcial.Checked;
    end;

    chkMezclar.Enabled := not Value;
    if giOpcion = 0 then begin
      txtOrdenMezclar.Enabled := not Value;
      txtCantidadMezclar.Enabled := not Value;
      AddOrden.Enabled := not Value;
      DeleteOrden.Enabled := not Value;
      gvMezclado.Enabled := not Value;
    end else begin
      txtOrdenMezclar.Enabled := chkMezclar.Checked;
      txtCantidadMezclar.Enabled := chkMezclar.Checked;
      AddOrden.Enabled := chkMezclar.Checked;
      DeleteOrden.Enabled := chkMezclar.Checked;
      gvMezclado.Enabled := chkMezclar.Checked;
    end;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;



procedure TfrmVentas.btnCancelarClick(Sender: TObject);
begin
ClearData();
giOpcion := 0;
EnableControls(True);

Nuevo.Enabled := True;
Button5.Enabled := True;
if Qry.RecordCount > 0 Then
begin
      Editar.Enabled := True;
      Borrar.Enabled := True;
      Buscar.Enabled := True;
end;
BindOrden();

EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.BindOrden();
var SQLStr, planoId : String;
Qry2 : TADOQuery;
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['ITE_Id']);
    txtOrden.Text := RightStr( VarToStr(Qry['ITE_Nombre']), Length(VarToStr(Qry['ITE_Nombre']))-3 );
    txtProceso.Text := VarToStr(Qry['TipoProceso']);
    txtRequerida.Text := VarToStr(Qry['Requerida']);
    lblRequerida.Caption := txtRequerida.Text;
    txtOrdenada.Text := VarToStr(Qry['Ordenada']);
    txtNumero.Text := VarToStr(Qry['Numero']);
    txtTerminal.Text := VarToStr(Qry['Terminal']);
    txtRequisicion.Text := VarTOStr(Qry['Requisicion']);
    deEntrega.Text := VarToStr(Qry['Entrega']);
    deInterna.Text := VarToStr(Qry['Interna']);
    txtRecibido.Text := VarToStr(Qry['Recibido']);
    txtUnitario.Text := VarToStr(Qry['Unitario']);
    txtObservaciones.Text := VarToStr(Qry['Observaciones']);
    txtOtras.Text := VarToStr(Qry['Otras']);
    txtTotal.Text := VarToStr(Qry['Total']);
    cmbEmpleados.Text := VarToStr(Qry['Nombre']);
    cmbProductos.Text := VarToStr(Qry['Producto']);
    chkAprobacion.Checked := StrToBool(VarToStr(Qry['Aprobacion']));
    chkDlls.Checked := StrToBool(VarToStr(Qry['Dolares']));
    txtCompra.Text := VarToStr(Qry['OrdenCompra']);
    chkStock.Checked := StrToBool(VarToStr(Qry['Stock']));

    chkplano.Checked := False;
    cmbPlanos.Text := '';
    planoId := VarToStr(Qry['PN_Id']);
    if planoId <> '' then begin
      chkplano.Checked := True;
      setValue(planoId, cmbPlanos);
    end;

    chkStockParcial.Checked := StrToBool(VarToStr(Qry['StockParcial']));
    txtStockParcial.Text := VarToStr(Qry['StockParcialCantidad']);
    lblStockParcial.Caption := txtStockParcial.Text;

    chkMezclar.Checked := StrToBool(VarToStr(Qry['Mezclado']));

    if chkMezclar.Checked then begin
      BindOrdenesMezclar();
    end;

    application.ProcessMessages;

    cmbPlanosChange(nil);
    txtNumeroExit(nil);

    if chkStock.Checked then
        Exit;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT TOP 1 T.Nombre,CASE WHEN I.ITS_Status = 0 THEN ''Listo'' ' +
              'WHEN I.ITS_Status = 1 THEN ''Activo'' ' +
              'WHEN I.ITS_Status = 2 THEN ''Terminado'' ' +
              'WHEN I.ITS_Status = 9 THEN ''Scrap'' END AS Status ' +
              'FROM tblitemtasks I ' +
              'INNER JOIN tbltareas T on I.TAS_ID = T.ID ' +
              'WHERE ITE_ID = ' + lblId.Caption + ' AND ITS_Status IS NOT NULL ' +
              'ORDER BY TAS_Order DESC';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvStatus.ClearRows;
    if Qry2.RecordCount > 0 then
    begin
        gvStatus.AddRow();
        gvStatus.Cells[0, gvStatus.RowCount -1] := VarToStr(Qry2['Nombre']);
        gvStatus.Cells[1, gvStatus.RowCount -1] := VarToStr(Qry2['Status']);
    end;

    Qry2.Close;
    Qry2.Free;

end;

procedure TfrmVentas.btnAceptarClick(Sender: TObject);
var SQLStr,sOrden : String;
Qry2 : TADOQuery;
sNew, planoId, stockParcial : String;
stock,cambio : boolean;
begin

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        sNew := gsYear + txtOrden.Text;

        planoId := 'NULL';
        if chkPlano.Checked then begin
          planoId := getSelectedValue(cmbPlanos);
        end;

        stockParcial := 'NULL';
        if chkStockParcial.Checked then begin
          stockParcial := txtStockParcial.Text;
        end;

        stock := false;
        if (chkStock.Checked) or (chkMezclar.Checked) then begin
          stock := True
        end;

        SQLStr := 'Insert_Orden ' + QuotedStr(gsYear + txtOrden.Text) + ',' + QuotedStr(txtProceso.Text) +
                  ',' + txtRequerida.Text + ',' + txtOrdenada.Text + ',' + QuotedStr(cmbProductos.Text) +
                  ',' + QuotedStr(txtNumero.Text) + ',' + QuotedStr(txtTerminal.Text) +
                  ',' + QuotedStr(deEntrega.Text) + ',' + QuotedStr(txtRecibido.Text) +
                  ',' + QuotedStr(deInterna.Text) +
                  ',' + QuotedStr(cmbEmpleados.Text) + ',' + BoolToStrInt(chkAprobacion.Checked) +
                  ',' + QuotedStr(txtObservaciones.Text) + ',' + QuotedStr(txtOtras.Text) +
                  ',' + txtUnitario.Text + ',' + txtTotal.Text + ',' + QuotedStr(gsFirstTask) +
                  ',' + QuotedStr(frmMain.sUserLogin) + ',' + QuotedStr(GetLocalIP) +
                  ',' + QuotedStr(txtCompra.Text) + ',' + BoolToStrInt(chkDlls.Checked) +
                  ',' + BoolToStrInt(chkStock.Checked) + ',' + planoId + ',' + BoolToStrInt(chkStockParcial.Checked) +
                  ',' + stockParcial + ',' + BoolToStrInt(chkMezclar.Checked) + ',' + BoolToStrInt(stock) +
                  ',' + QuotedStr(txtRequisicion.Text);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        if VarToStr(Qry2['ERROR']) = '-1' Then
          begin
                ShowMessage(VarToStr(Qry2['MSG']));
                Exit;
          end
        else
          begin
            if chkMezclar.Checked then begin
              InsertOrdenesMezclar();
            end;

            Qry.SQL.Clear;
            Qry.SQL.Text := 'Traer_Ordenes ' + QuotedStr(gsOYear);
            Qry.Open;

            Qry.Locate('ITE_Nombre', sNew, [loPartialKey] );

            if MessageDlg('Quieres Imprimir la orden de trabajo?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
                Button5Click(nil);
            end;

          end;
        Qry2.Close;

        if lblStock.Caption = '1' then
        begin
            sOrden := gsOYear + '-' + txtOrden.Text;
            SQLStr := 'UPDATE tblStockOrdenes SET Programado = 1 WHERE ITE_Nombre = ' + QuotedStr(sOrden);

            Conn.Execute(SQLStr);
        end;

  end
  else if giOpcion = 2 then
  begin
        cambio := false;
        if not ValidateData() then
          Exit;

        //stock := StrToBool(VarToStr(Qry['Stock']));
        //if stock <> chkStock.Checked then
        //        cambio := true;

        sNew := gsYear + txtOrden.Text;

        planoId := 'NULL';
        if chkPlano.Checked then begin
          planoId := getSelectedValue(cmbPlanos);
        end;

        stockParcial := 'NULL';
        if chkStockParcial.Checked then begin
          stockParcial := txtStockParcial.Text;
        end;

        SQLStr := 'Update_Orden ' + lblId.Caption  + ',' + QuotedStr(gsYear + txtOrden.Text) + ',' + QuotedStr(txtProceso.Text) +
                  ',' + txtRequerida.Text + ',' + txtOrdenada.Text + ',' + QuotedStr(cmbProductos.Text) +
                  ',' + QuotedStr(txtNumero.Text) + ',' + QuotedStr(txtTerminal.Text) +
                  ',' + QuotedStr(deEntrega.Text) + ',' + QuotedStr(txtRecibido.Text) +
                  ',' + QuotedStr(deInterna.Text) +
                  ',' + QuotedStr(cmbEmpleados.Text) + ',' + BoolToStrInt(chkAprobacion.Checked) +
                  ',' + QuotedStr(txtObservaciones.Text) + ',' + QuotedStr(txtOtras.Text) +
                  ',' + txtUnitario.Text + ',' + txtTotal.Text + ',' + QuotedStr(gsFirstTask) +
                  ',' + QuotedStr(frmMain.sUserLogin) + ',' + QuotedStr(GetLocalIP) +
                  ',' + QuotedStr(txtCompra.Text) + ',' + BoolToStrInt(chkDlls.Checked) +
                  ',' + BoolToStrInt(cambio) + ',' + BoolToStrInt(chkStock.Checked)+ ',' + planoId +
                  ',' + BoolToStrInt(chkStockParcial.Checked) + ',' + stockParcial +
                  ',' + BoolToStrInt(chkMezclar.Checked) + ',' + QuotedStr(txtRequisicion.Text);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        if VarToStr(Qry2['ERROR']) = '-1' Then
          begin
                ShowMessage(VarToStr(Qry2['MSG']));
                Exit;
          end
        else
          begin
            if chkMezclar.Checked then begin
              InsertOrdenesMezclar();
            end;

            Qry.SQL.Clear;
            Qry.SQL.Text := 'Traer_Ordenes ' + QuotedStr(gsOYear);
            Qry.Open;

            Qry.Locate('ITE_Nombre', sNew, [loPartialKey] )
          end;

        Qry2.Close;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar la orden : ' +
                      txtOrden.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar la orden : ' +
                            txtOrden.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin

              SQLStr := 'Delete_Orden ' + lblId.Caption + ',' + BoolToStrInt(chkStock.Checked);
              Qry2.SQL.Clear;
              Qry2.SQL.Text := SQLStr;
              Qry2.Open;

              if VarToStr(Qry2['ERROR']) = '-1' Then
                begin
                      ShowMessage(VarToStr(Qry2['MSG']));
                      Exit;
                end
              else
                begin
                  Qry.SQL.Clear;
                  //Qry.SQL.Text := 'SELECT * FROM tblOrdenes ORDER BY ITE_ID';
                  Qry.SQL.Text := 'Traer_Ordenes ' + QuotedStr(gsOYear);
                  Qry.Open;

                  Qry.First;
                end;

              end;
              Qry2.Close;
        end;
  end
  else if giOpcion = 4 then
  begin
        if not Qry.Locate('ITE_Nombre',gsYear + txtOrden.Text ,[loPartialKey] ) then
          begin
              MessageDlg('No se encontro ninguna Orden con estos datos.', mtInformation,[mbOk], 0);
              txtOrden.SetFocus;
              Exit;
          end;

          //Exit;
  end;


  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  Nuevo.Enabled := True;
  Button5.Enabled := True;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;

  end;
  BindOrden;
  EnableFormButtons(gbButtons, sPermits);
  if giOpcion = 1 then begin
        NuevoClick(nil);
  end
  else begin
    giOpcion := 0;
    EnableControls(True);
  end;
end;

function TfrmVentas.ValidateData():Boolean;
var i, opciones, stock, diff,cantidad :Integer;
bfound : boolean;
stockVal : String;
begin
        result := True;

        if cmbEmpleados.Text = '' then
          begin
            MessageDlg('Por favor captura el empleado.' , mtInformation,[mbOk], 0);
            result :=  False;
          end;


        if cmbProductos.Text = '' then
          begin
            MessageDlg('Por favor captura la descripcion.' , mtInformation,[mbOk], 0);
            result :=  False;
          end;

        bfound := False;
        for i:= 0 to cmbProductos.Items.Count do
        begin
                if cmbProductos.Text = cmbproductos.Items[i] then
                begin
                     bfound := True;
                     break;
                end;
        end;

        if bfound = false then
          begin
            MessageDlg('Descripcion Incorrecta : ' + cmbProductos.Text , mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(Copy(txtOrden.Text,0,3)) then
          begin
            MessageDlg('Cliente incorrecto: ' + Copy(txtOrden.Text,0,3), mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not ValidateCliente(Copy(txtOrden.Text,0,3)) then
          begin
            MessageDlg('Cliente incorrecto: ' + Copy(txtOrden.Text,0,3), mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(Copy(txtOrden.Text,5,3)) then
          begin
            MessageDlg('El numero de orden debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not isnumeric(RightStr(txtOrden.Text,2)) then
          begin
            MessageDlg('El numero de orden debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

{        if not IsDate(txtRecibido.Text) Then
          begin
            MessageDlg('Por favor escriba una fecha de recibido valida.', mtInformation,[mbOk], 0);
            result :=  False;
          end;
 }
        if not IsDate(deEntrega.Text) Then
          begin
            MessageDlg('Por favor escriba una fecha de Entrega valida.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsDate(deInterna.Text) Then
          begin
            MessageDlg('Por favor escriba una fecha Interna valida.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtUnitario.Text) Then
          begin
            MessageDlg('El Valor Unitario debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtRequerida.Text) Then
          begin
            MessageDlg('La cantidad requerida debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtOrdenada.Text) Then
          begin
            MessageDlg('La cantidad ordenada debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        cmbPlanos.Text := Trim(cmbPlanos.Text);
        if (chkPlano.Checked) and (cmbPlanos.Text = '') then
        begin
            MessageDlg('El Numero de Plano es requerido.', mtInformation,[mbOk], 0);
            result :=  False;
            Exit;
        end;

        if (chkPlano.Checked) and (cmbPlanos.Items.IndexOf(cmbPlanos.Text) = -1) then
        begin
            MessageDlg('Numero de Plano incorrecto seleccionelo de la lista.', mtInformation,[mbOk], 0);
            result :=  False;
            Exit;
        end;

        opciones := 0;
        if chkStock.Checked then Inc(opciones);
        if chkStockParcial.Checked then Inc(opciones);
        if chkMezclar.Checked then Inc(opciones);

        if opciones > 1 then
        begin
            MessageDlg('Solo puedes seleccionar una de las siguientes opciones Stock, StockParcial, Mezclar.', mtInformation,[mbOk], 0);
            result :=  False;
            Exit;
        end;

        txtStockParcial.Text := Trim(txtStockParcial.Text);
        if chkStockParcial.Checked then
        begin
            if cmbPlanos.Text = '' then
            begin
              MessageDlg('Numero de Plano es requerido para ordenes con Stock Parcial.', mtInformation,[mbOk], 0);
              result :=  False;
              Exit;
            end;

            if txtStockParcial.Text = '' then
            begin
              MessageDlg('La Cantidad de Stock Parcial es requerida.', mtInformation,[mbOk], 0);
              result :=  False;
              Exit;
            end;

            if not IsNumeric(txtStockParcial.Text) then
            begin
                MessageDlg('La Cantidad de Stock Parcial debe de ser un valor numerico.', mtInformation,[mbOk], 0);
                result :=  False;
                Exit;
            end;

            if StrToInt(txtStockParcial.Text) <= 0 then
            begin
                MessageDlg('La Cantidad de Stock Parcial debe de ser mayor que 0.', mtInformation,[mbOk], 0);
                result :=  False;
                Exit;
            end;

            if StrToInt(txtStockParcial.Text) >= StrToInt(txtRequerida.Text) then
            begin
                MessageDlg('La Cantidad de Stock Parcial debe de ser menor que la cantidad Cliente.', mtInformation,[mbOk], 0);
                result :=  False;
                Exit;
            end;

            if lblStockParcial.Caption = '' then lblStockParcial.Caption := '0';

            if giOpcion = 1 then begin // Only validate in new not in updates
                stock := 0;
                stockVal := gvNumPlano.Cell[2, 0].AsString;
                if stockVal <> '' then
                  stock := StrToInt(stockVal);

                if StrToInt(txtStockParcial.Text) > stock then
                begin
                    MessageDlg('La Cantidad de Stock Parcial debe de ser menor o igual que la cantidad en Stock(' + IntToStr(stock) + ').', mtInformation,[mbOk], 0);
                    result :=  False;
                    Exit;
                end;
            end;

            if (giOpcion = 2) and (StrToInt(lblStockParcial.Caption) <> StrToInt(txtStockParcial.Text)) then begin
                stock := 0;
                stockVal := gvNumPlano.Cell[2, 0].AsString;
                if stockVal <> '' then
                  stock := StrToInt(stockVal);

                if  StrToInt(txtStockParcial.Text) > StrToInt(lblStockParcial.Caption) then begin
                  diff := StrToInt(txtStockParcial.Text) - StrToInt(lblStockParcial.Caption);
                  if diff > stock then
                  begin
                      MessageDlg('El aumento en la Cantidad de Stock Parcial (cambio de ' + lblStockParcial.Caption +
                      ' a ' + txtStockParcial.Text + ') debe de ser menor o igual que la cantidad en Stock (' + IntToStr(stock) + ').', mtInformation,[mbOk], 0);
                      result :=  False;
                      Exit;
                  end;
                end;
            end;

        end;

        if chkStock.Checked then
        begin
            if cmbPlanos.Text = '' then
            begin
              MessageDlg('Numero de Plano es requerido para ordenes con Stock.', mtInformation,[mbOk], 0);
              result :=  False;
              Exit;
            end;

            stock := 0;
            stockVal := gvNumPlano.Cell[2, 0].AsString;
            if stockVal <> '' then
              stock := StrToInt(stockVal);

            if lblRequerida.Caption = '' then lblRequerida.Caption := '0';
            if giOpcion = 1 then begin // Only validate in new not in updates
              if StrToInt(txtRequerida.Text) > stock then
              begin
                  MessageDlg('La Cantidad de Cliente debe de ser menor o igual que la cantidad en Stock(' + IntToStr(stock) + ').', mtInformation,[mbOk], 0);
                  result :=  False;
                  Exit;
              end;
            end;

            if (giOpcion = 2) and (StrToInt(lblRequerida.Caption) <> StrToInt(txtRequerida.Text)) then begin
                if  StrToInt(txtRequerida.Text) > StrToInt(lblRequerida.Caption) then begin
                  diff := StrToInt(txtRequerida.Text) - StrToInt(lblRequerida.Caption);
                  if diff > stock then
                  begin
                      MessageDlg('El aumento en la Cantidad de Cliente (cambio de ' + lblRequerida.Caption +
                      ' a ' + txtRequerida.Text + ') debe de ser menor o igual que la cantidad en Stock (' + IntToStr(stock) + ').', mtInformation,[mbOk], 0);
                      result :=  False;
                      Exit;
                  end;
                end;
            end
            else begin
              if (giOpcion = 2) and (StrToInt(lblRequerida.Caption) = StrToInt(txtRequerida.Text)) then begin
                if StrToInt(txtRequerida.Text) > stock then
                begin
                    MessageDlg('La Cantidad de Cliente debe de ser menor o igual que la cantidad en Stock(' + IntToStr(stock) + ').', mtInformation,[mbOk], 0);
                    result :=  False;
                    Exit;
                end;
              end;
            end;
        end;

        if (chkMezclar.Checked) and (gvMezclado.RowCount = 0) then begin
          MessageDlg('Es necesario agregar al menos una Orden con la que se va a Mezclar.', mtInformation,[mbOk], 0);
          result :=  False;
          Exit;
        end;


        if (chkMezclar.Checked) then begin
          cantidad := 0;
          for i:= 0 to gvMezclado.RowCount - 1 do
          begin
             cantidad := cantidad + StrToInt(gvMezclado.Cells[1,i]);
          end;

          if (cantidad <> StrToInt(txtRequerida.Text)) then begin
            MessageDlg('La Cantidad Cliente debe de ser igual que la sumatoria de las ordenes a mezclar.', mtInformation,[mbOk], 0);
            result :=  False;
            Exit;
          end;
        end;
end;

function  TfrmVentas.ValidateCliente(Clave:String):Boolean;
var Qry2 : TADOQuery;
begin
    Result := False;
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := 'SELECT Clave FROM tblClientes WHERE Clave = ' + QuotedStr(Clave);
    Qry2.Open;

    If Qry2.RecordCount > 0 Then
        Result := True;

end;

function TfrmVentas.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;


procedure TfrmVentas.Button1Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.First;
BindOrden;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.PrimeroClick(Sender: TObject);
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

  ClearData();
  BindOrden();
  EnableButtons();
end;

procedure TfrmVentas.Button2Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Prior;

ClearData();
BindOrden;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.Button3Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Next;
BindOrden;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.Button4Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Last;
BindOrden;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.txtRequeridaKeyPress(Sender: TObject; var Key: Char);
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
procedure TfrmVentas.txtUnitarioKeyPress(Sender: TObject; var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar(txtUnitario.Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;

end;

procedure TfrmVentas.txtUnitarioChange(Sender: TObject);
begin
        if (txtUnitario.Text = '') or (txtRequerida.Text = '') Then
        begin
                txtTotal.Text := '';
                Exit;
        end
        else
                txtTotal.Text := FloatToStr( StrToFloat(txtUnitario.Text) * StrToFloat(txtRequerida.Text) );
end;

procedure TfrmVentas.BindGrid();
var iCurrent : Integer;
begin
  iCurrent := Qry.RecNo;

  gvVentas.ClearRows;
  Qry.First;
  while not Qry.Eof do
  begin
      gvVentas.AddRow(1);
      gvVentas.Cells[0,gvVentas.RowCount -1] := RightStr( VarToStr(Qry['ITE_Nombre']), Length(VarToStr(Qry['ITE_Nombre']))-3 );
      gvVentas.Cells[1,gvVentas.RowCount -1] := VarToStr(Qry['TipoProceso']);
      gvVentas.Cells[2,gvVentas.RowCount -1] := VarToStr(Qry['Requerida']);
      gvVentas.Cells[3,gvVentas.RowCount -1] := VarToStr(Qry['Ordenada']);
      gvVentas.Cells[4,gvVentas.RowCount -1] := VarToStr(Qry['Producto']);
      gvVentas.Cells[5,gvVentas.RowCount -1] := VarToStr(Qry['Numero']);
      gvVentas.Cells[6,gvVentas.RowCount -1] := VarToStr(Qry['Terminal']);
      gvVentas.Cells[7,gvVentas.RowCount -1] := VarToStr(Qry['Recibido']);
      gvVentas.Cells[8,gvVentas.RowCount -1] := VarToStr(Qry['Entrega']);
      gvVentas.Cells[9,gvVentas.RowCount -1] := VarToStr(Qry['Nombre']);
      Qry.Next;
  end;

  Qry.RecNo := iCurrent;
end;



procedure TfrmVentas.TabSheet2Show(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmVentas.btnExportClick(Sender: TObject);
var sFileName: String;
begin

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(gvVentas,sFileName);

  end;
end;

procedure TfrmVentas.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
var XApp : Variant;
Sheet : Variant;
Row,col :Integer;
begin
      Try //Create the excel object
      Begin
            XApp:= CreateOleObject('Excel.Application');
            //XApp.Visible := True;
            XApp.Visible := False;
            XApp.DisplayAlerts := False;
      end;
      except
       showmessage('No se pudo abrir Microsoft Excel,  parece que no esta instalado en el sistema.');
       exit;
      end;

      XApp.Workbooks.Add(xlWorkSheet);
      Sheet := XApp.Workbooks[1].WorkSheets[1];
      Sheet.Name := 'Empleados';

      for Col := 1 to Grid.Columns.Count do
              Sheet.Cells[1,Col] := Grid.Columns[Col - 1].Header.Caption;

      for Row := 1 to Grid.RowCount do
                for Col := 1 to Grid.Columns.Count do
                        Sheet.Cells[Row + 1,Col] := Grid.Cells[Col - 1,Row - 1];


      Sheet.Cells.Select;
      Sheet.Cells.EntireColumn.AutoFit;

      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;
end;
procedure TfrmVentas.cmbProductosDropDown(Sender: TObject);
begin
BindProductos();
end;

procedure TfrmVentas.txtRequeridaExit(Sender: TObject);
begin
if (txtOrdenada.ReadOnly = False) and (txtOrdenada.Text = '') then
        txtOrdenada.Text := txtRequerida.Text
end;

procedure TfrmVentas.cmbProductosChange(Sender: TObject);
begin
txtOtras.Text := cmbProductos.Text;
end;

procedure TfrmVentas.Button5Click(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrImpresionOrden,qrImpresionOrden);

    qrImpresionOrden.QROrden.Caption := txtOrden.Text;
    qrImpresionOrden.QROrden2.Caption := txtOrden.Text;

    qrImpresionOrden.QRCode1.Caption := '*' + gsYear + txtOrden.Text + '*';
    qrImpresionOrden.QRCode2.Caption := '*' + gsYear + txtOrden.Text + '*';
    qrImpresionOrden.QRSemana.Caption := '';
    qrImpresionOrden.QRNumero.Caption := txtNumero.Text;
    qrImpresionOrden.QRTerminal.Caption := txtTerminal.Text;
    qrImpresionOrden.QRRecibido.Caption := txtRecibido.Text;
    qrImpresionOrden.QREntrega.Caption := deInterna.Text;
    qrImpresionOrden.QREntrega2.Caption := deInterna.Text;
    qrImpresionOrden.QRNombre.Caption := cmbEmpleados.Text;
    qrImpresionOrden.QRFirma.Caption := '';
    qrImpresionOrden.QRObs.Caption := txtObservaciones.Text;
    qrImpresionOrden.QRDesc.Caption := cmbProductos.Text;
    qrImpresionOrden.QRCompra.Caption := txtCompra.Text;
    qrImpresionOrden.QRProceso.Caption := txtProceso.Text;
    qrImpresionOrden.QRCantidad.Caption := txtOrdenada.Text;

    qrImpresionOrden.QRMsg.Caption := 'Forma: Larco-015' + #13 +
                                      'Nivel de Revisión: D' + #13 +
                                      'Retención: 1 año+uso';
    //qrImpresionOrden.Print;
    qrImpresionOrden.Preview;
    qrImpresionOrden.Free;
end;

procedure TfrmVentas.txtOrdenChange(Sender: TObject);
var SQLStr,sOrden,sInterna,sCompra,sEntrega : String;
sLast : integer;
Qry2 : TADOQuery;
begin
  if giOpcion <> 1 then
        exit;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  sOrden := TrimRight( StringReplace(txtOrden.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
  if Length(sOrden) = 3 then
  begin
        SQLStr := 'SELECT TOP 1 * FROM tblOrdenes WHERE SUBSTRING(ITE_Nombre,4,3) = ' +
                   QuotedStr( LeftStr(txtOrden.Text,3) ) + ' ORDER BY ITE_ID desc';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        if Qry2.RecordCount > 0 then
        begin
                sOrden := VarToStr(Qry2['ITE_Nombre']);
                sInterna := VarToStr(Qry2['Interna']);
                sEntrega := VarToStr(Qry2['Entrega']);
                sCompra := VarToStr(Qry2['OrdenCompra']);
                sLast := StrToInt( RightStr(sOrden,2) );

                txtOrden.Text := leftStr(txtOrden.Text,3) + '-' + Copy(sOrden,8,3) + '-' + FormatFloat('00',sLast + 1);
                txtCompra.Text := sCompra;
                deInterna.Date := StrToDate(sInterna);
                deEntrega.Date := StrToDate(sEntrega);
                txtRequisicion.Text := VarToStr(Qry2['Requisicion']);
                txtTerminal.Text := VarToStr(Qry2['Terminal']);
                txtOrden.SelStart := 8;
        end;

        Qry2.Close;
  end;

end;

procedure TfrmVentas.deInternaChange(Sender: TObject);
begin
if giOpcion <> 1 then
        Exit;

deEntrega.Date := DateAdd(deInterna.Date,1,daDays);
end;

procedure TfrmVentas.deInternaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;

   If Key = vk_up then
   begin
        deInterna.Date := DateAdd(deInterna.Date,1,daDays);
        deEntrega.Date := DateAdd(deInterna.Date,1,daDays);
   end;

   If Key = vk_down then
   begin
        deInterna.Date := DateAdd(deInterna.Date,-1,daDays);
        deEntrega.Date := DateAdd(deInterna.Date,1,daDays);
   end;

    if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
    begin
            btnCancelarClick(nil);
    end;

end;

procedure TfrmVentas.deEntregaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;

   If Key = vk_up then
   begin
        deEntrega.Date := DateAdd(deEntrega.Date,1,daDays);
   end;

   If Key = vk_down then
   begin
        deEntrega.Date := DateAdd(deEntrega.Date,-1,daDays);
   end;

end;

procedure TfrmVentas.cmbEmpleadosDropDown(Sender: TObject);
begin
BindEmpleados();
end;

function TfrmVentas.FormIsRunning(FormName: String):Boolean;
var i:Integer;
begin
  Result := False;

  for  i := 0 to Screen.FormCount - 1 do
  begin
        if Screen.Forms[i].Name = FormName Then
          begin
                Result:= True;
                Break;
          end;
  end;

end;

procedure TfrmVentas.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

        if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
        begin
                btnCancelarClick(nil);
        end;

end;

procedure TfrmVentas.chkPlanoClick(Sender: TObject);
begin
  if giOpcion <> 0 then begin
    cmbPlanos.Enabled := chkPlano.Checked;
    if not chkPlano.Checked then
      cmbPlanos.Text := '';
  end;
  
end;

procedure TfrmVentas.cmbPlanosDropDown(Sender: TObject);
begin
BindPlanos();
end;

procedure TfrmVentas.cmbPlanosChange(Sender: TObject);
var SQLStr, enStock, piezas, ordenes: String;
Qry2 : TADOQuery;
begin
  cmbPlanos.Text := Trim(cmbPlanos.Text);
  if cmbPlanos.Text = '' then
    Exit;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT P.PN_Id,P.PN_Numero, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) AS Entradas, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Salidas, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) - SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Cantidad ' +
            'FROM tblPlano P ' +
            'INNER JOIN tblStock S ON P.PN_Id = S.PN_Id AND P.PN_Numero = ' + QuotedStr(cmbPlanos.Text) + ' ' +
            'GROUP BY P.PN_Id,P.PN_Numero';


  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  enStock := '';
  piezas := '';
  ordenes := '';

  if Qry2.RecordCount > 0 then begin
      enStock := VarToStr(Qry2['Cantidad']);
  end;

  SQLStr := 'SELECT P.PN_Id,P.PN_Numero, COUNT(*) AS Ordenes, SUM(O.Requerida) AS Piezas ' +
            'FROM tblOrdenes O ' +
            'INNER JOIN tblPlano P ON O.PN_Id = P.PN_Id AND P.PN_Numero = ' + QuotedStr(cmbPlanos.Text) + ' ' +
            'WHERE O.Recibido <= GETDATE() AND O.Recibido >= DATEADD(MONTH, -6, GETDATE()) ' +
            'GROUP BY P.PN_Id,P.PN_Numero';

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
      piezas := VarToStr(Qry2['Piezas']);
      ordenes := VarToStr(Qry2['Ordenes']);
  end;

  gvNumPlano.ClearRows;
  gvNumPlano.AddRow();
  gvNumPlano.Cells[0, gvNumPlano.RowCount -1] := ordenes;
  gvNumPlano.Cells[1, gvNumPlano.RowCount -1] := piezas;
  gvNumPlano.Cells[2, gvNumPlano.RowCount -1] := enStock;

  SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden, ' +
            'O.Ordenada As Cantidad, O.Requerida As Cliente, ' +
            'T.Nombre AS Tarea, ' +
            'CASE WHEN I.ITS_Status = 0 THEN ''Listo'' ' +
            'WHEN I.ITS_Status = 1 THEN ''Activo'' ' +
            'WHEN I.ITS_Status = 2 THEN ''Terminado'' END AS Status, ' +
            'SUM(CASE WHEN MO_Cantidad IS NULL THEN 0 ELSE MO_Cantidad END) As Usado, ' +
            '(O.Ordenada + CASE WHEN StockParcialCantidad IS NULL THEN 0 ELSE StockParcialCantidad END) - ' +
            '(SUM(CASE WHEN MO_Cantidad IS NULL THEN 0 ELSE MO_Cantidad END) + Requerida) As Disponible ' +
            'FROM tblOrdenes O  ' +
            'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID ' +
            'AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL AND I.ITS_Status <> 9 ' +
            'AND LEFT(O.ITE_Nombre,2) = ' + QuotedStr(gsOYear) + ' AND O.ITE_Nombre <> ' + QuotedStr(gsYear + txtOrden.Text) + ' ' +
            'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ' +
            'INNER JOIN tblPlano P ON O.PN_Id = P.PN_Id AND P.PN_Numero = ' + QuotedStr(cmbPlanos.Text) + ' ' +
            'LEFT OUTER JOIN tblMergeOrdenes M ON O.ITE_Nombre = M.MO_ITE_Nombre ' +
            'GROUP BY O.ITE_Nombre, O.Ordenada, O.Requerida, T.Nombre,I.ITS_Status,StockParcialCantidad';

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvProdPlano.ClearRows;
  while not Qry2.Eof do begin
    gvProdPlano.AddRow();
    gvProdPlano.Cells[0, gvProdPlano.RowCount -1] := VarToStr(Qry2['Orden']);
    gvProdPlano.Cells[1, gvProdPlano.RowCount -1] := VarToStr(Qry2['Cantidad']);
    gvProdPlano.Cells[2, gvProdPlano.RowCount -1] := VarToStr(Qry2['Cliente']);
    gvProdPlano.Cells[3, gvProdPlano.RowCount -1] := VarToStr(Qry2['Disponible']);
    gvProdPlano.Cells[4, gvProdPlano.RowCount -1] := VarToStr(Qry2['Tarea']);
    gvProdPlano.Cells[5, gvProdPlano.RowCount -1] := VarToStr(Qry2['Status']);
    Qry2.Next;
  end;

  Qry2.Close;
  Qry2.Free;
end;

procedure TfrmVentas.EnableButtons();
begin
  if giOpcion = 0 then begin
    Nuevo.Enabled := True;
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    Button5.Enabled := True; //Imprimir
    if Qry.RecordCount > 0 Then
    begin
          Editar.Enabled := True;
          Borrar.Enabled := True;
          Buscar.Enabled := True;

          EnableFormButtons(gbButtons, sPermits);
    end;
  end
  else if giOpcion = 1 then begin
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    Button5.Enabled := False; //Imprimir
  end
  else if giOpcion = 2 then begin
    Nuevo.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    Button5.Enabled := False; //Imprimir
  end
  else if giOpcion = 3 then begin
    Nuevo.Enabled := False;
    Editar.Enabled := False;
    Buscar.Enabled := False;
    Button5.Enabled := False; //Imprimir
  end
  else if giOpcion = 4 then begin
    Nuevo.Enabled := False;
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Button5.Enabled := False; //Imprimir
  end;

  if giOpcion = 0 then begin
    btnAceptar.Enabled := False;
    btnCancelar.Enabled := False;
  end else begin
    btnAceptar.Enabled := True;
    btnCancelar.Enabled := True;
  end;

end;

procedure TfrmVentas.txtNumeroExit(Sender: TObject);
var SQLStr, enStock, piezas, ordenes: String;
Qry2 : TADOQuery;
begin
  txtNumero.Text := Trim(txtNumero.Text);
  if txtNumero.Text = '' then
    Exit;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PA.PN_Id,PA.PA_Alias, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) AS Entradas, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Salidas, ' +
            'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) - SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Cantidad ' +
            'FROM tblPlanoAlias PA ' +
            'INNER JOIN tblStock S ON PA.PN_Id = S.PN_Id AND PA.PA_Alias = ' + QuotedStr(txtNumero.Text) + ' ' +
            'GROUP BY PA.PN_Id,PA.PA_Alias';


  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  enStock := '';
  piezas := '';
  ordenes := '';

  if Qry2.RecordCount > 0 then begin
      enStock := VarToStr(Qry2['Cantidad']);
  end;

  SQLStr := 'SELECT COUNT(*) AS Ordenes, SUM(O.Requerida) AS Piezas ' +
            'FROM tblOrdenes O ' +
            'WHERE O.Recibido <= GETDATE() AND O.Recibido >= DATEADD(MONTH, -6, GETDATE()) AND O.Numero = ' + QuotedStr(txtNumero.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
      piezas := VarToStr(Qry2['Piezas']);
      ordenes := VarToStr(Qry2['Ordenes']);
  end;

  gvNumParte.ClearRows;
  gvNumParte.AddRow();
  gvNumParte.Cells[0, gvNumParte.RowCount -1] := ordenes;
  gvNumParte.Cells[1, gvNumParte.RowCount -1] := piezas;
  gvNumParte.Cells[2, gvNumParte.RowCount -1] := enStock;

  SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden, ' +
            'O.Ordenada As Cantidad, O.Requerida As Cliente, ' +
            'T.Nombre AS Tarea, ' +
            'CASE WHEN I.ITS_Status = 0 THEN ''Listo'' ' +
            'WHEN I.ITS_Status = 1 THEN ''Activo'' ' +
            'WHEN I.ITS_Status = 2 THEN ''Terminado'' END AS Status, ' +
            'SUM(CASE WHEN MO_Cantidad IS NULL THEN 0 ELSE MO_Cantidad END) As Usado, ' +
            '(O.Ordenada + CASE WHEN StockParcialCantidad IS NULL THEN 0 ELSE StockParcialCantidad END) - ' +
            '(SUM(CASE WHEN MO_Cantidad IS NULL THEN 0 ELSE MO_Cantidad END) + Requerida) As Disponible ' +
            'FROM tblOrdenes O  ' +
            'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID ' +
            'AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL AND I.ITS_Status <> 9 ' +
            'AND LEFT(O.ITE_Nombre,2) = ' + QuotedStr(gsOYear) + ' AND O.Numero = ' + QuotedStr(txtNumero.Text) +
            ' AND O.ITE_Nombre <> ' + QuotedStr(gsYear + txtOrden.Text) + ' ' +
            'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ' +
            'LEFT OUTER JOIN tblMergeOrdenes M ON O.ITE_Nombre = M.MO_ITE_Nombre ' +
            'GROUP BY O.ITE_Nombre, O.Ordenada, O.Requerida, T.Nombre,I.ITS_Status,StockParcialCantidad';

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvProdNumero.ClearRows;
  while not Qry2.Eof do begin
    gvProdNumero.AddRow();
    gvProdNumero.Cells[0, gvProdNumero.RowCount -1] := VarToStr(Qry2['Orden']);
    gvProdNumero.Cells[1, gvProdNumero.RowCount -1] := VarToStr(Qry2['Cantidad']);
    gvProdNumero.Cells[2, gvProdNumero.RowCount -1] := VarToStr(Qry2['Cliente']);
    gvProdNumero.Cells[3, gvProdNumero.RowCount -1] := VarToStr(Qry2['Disponible']);
    gvProdNumero.Cells[4, gvProdNumero.RowCount -1] := VarToStr(Qry2['Tarea']);
    gvProdNumero.Cells[5, gvProdNumero.RowCount -1] := VarToStr(Qry2['Status']);
    Qry2.Next;
  end;

  Qry2.Close;
  Qry2.Free;
end;

procedure TfrmVentas.txtOrdenKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   if (Key = vk_return) and (giopcion = 4)then begin // buscar
     btnAceptarClick(nil);
   end
   else if Key = vk_return then begin
     AppActivate(Application.Handle);
     SendKeys('{TAB}',False);
   end
   else if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then begin
     btnCancelarClick(nil);
   end;
end;

procedure TfrmVentas.chkStockParcialClick(Sender: TObject);
begin
  if giOpcion <> 0 then begin
    txtStockParcial.Enabled := chkStockParcial.Checked;
    if not chkStockParcial.Checked then
      txtStockParcial.Text := '';
  end;
end;

procedure TfrmVentas.chkMezclarClick(Sender: TObject);
begin
  if giOpcion <> 0 then begin
    txtOrdenMezclar.Enabled := chkMezclar.Checked;
    txtCantidadMezclar.Enabled := chkMezclar.Checked;
    AddOrden.Enabled := chkMezclar.Checked;
    DeleteOrden.Enabled := chkMezclar.Checked;
    gvMezclado.Enabled := chkMezclar.Checked;

    if not chkMezclar.Checked then begin
      gvMezclado.ClearRows;
      txtOrdenMezclar.Text := '';
      txtCantidadMezclar.Text := '';
    end
  end;
end;

procedure TfrmVentas.DeleteOrdenClick(Sender: TObject);
var i, cantidad : Integer;
begin

  for i:= 0 to gvProdNumero.RowCount - 1 do
  begin
     if gvProdNumero.Cells[0,i] = gvMezclado.Cells[0, gvMezclado.SelectedRow] then begin
        cantidad := StrToInt(gvProdNumero.Cells[3, i]) + StrToInt(gvMezclado.Cells[1, gvMezclado.SelectedRow]);
        gvProdNumero.Cells[3, i] := IntToStr(cantidad);
        break;
     end;
  end;

  for i:= 0 to gvProdPlano.RowCount - 1 do
  begin
     if gvProdPlano.Cells[0,i] = gvMezclado.Cells[0, gvMezclado.SelectedRow] then begin
        cantidad := StrToInt(gvProdPlano.Cells[3, i]) + StrToInt(gvMezclado.Cells[1, gvMezclado.SelectedRow]);
        gvProdPlano.Cells[3, i] := IntToStr(cantidad);
        break;
     end;
  end;

  gvMezclado.DeleteRow(gvMezclado.SelectedRow);
end;

procedure TfrmVentas.txtCantidadMezclarKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if Key = vk_return then
  begin
    AddOrdenClick(nil);
  end;
end;

procedure TfrmVentas.AddOrdenClick(Sender: TObject);
var sOrden, sOrdenTra : String;
found : boolean;
i,cantidad : integer;
begin
  sOrdenTra := Trim( StringReplace(txtOrden.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
  sOrden := Trim( StringReplace(txtOrdenMezclar.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
  if sOrden = '' then begin
      ShowMessage('La Orden es requerida.');
      Exit;
  end;

  if sOrden = sOrdenTra then begin
      ShowMessage('No se puede mezclar con la misma orden.');
      Exit;
  end;

  txtCantidadMezclar.Text := Trim(txtCantidadMezclar.Text);
  if txtCantidadMezclar.Text = '' then begin
      ShowMessage('La Cantidad es Requerida.');
      Exit;
  end;

  if not IsNumeric(txtCantidadMezclar.Text) then begin
    MessageDlg('La Cantidad debe de ser un valor numerico.', mtInformation,[mbOk], 0);
    Exit;
  end;

  for i:= 0 to gvMezclado.RowCount - 1 do
  begin
     if gvMezclado.Cells[0,i] = txtOrdenMezclar.Text then begin
        MessageDlg('La Orden ya existe.', mtInformation,[mbOk], 0);
        Exit;
     end;
  end;

  found := false;
  cantidad := 0;
  for i:= 0 to gvProdNumero.RowCount - 1 do
  begin
     if gvProdNumero.Cells[0,i] = txtOrdenMezclar.Text then begin
        found := True;
        cantidad := StrToInt(gvProdNumero.Cells[3, i]);
        break;
     end;
  end;

  for i:= 0 to gvProdPlano.RowCount - 1 do
  begin
     if gvProdPlano.Cells[0,i] = txtOrdenMezclar.Text then begin
        found := True;
        cantidad := StrToInt(gvProdPlano.Cells[3, i]);
        break;
     end;
  end;

  if found = false then begin
    MessageDlg('La Orden no es valida, no es una orden en produccion.', mtInformation,[mbOk], 0);
    Exit;
  end;

  if StrToInt(txtCantidadMezclar.Text) > cantidad then begin
    MessageDlg('La cantidad es mayor que lo disponible en esta orden.', mtInformation,[mbOk], 0);
    Exit;
  end;

  gvMezclado.AddRow(1);
  gvMezclado.Cells[0,gvMezclado.RowCount -1] := txtOrdenMezclar.Text;
  gvMezclado.Cells[1,gvMezclado.RowCount -1] := txtCantidadMezclar.Text;

  txtOrdenMezclar.Text := '';
  txtCantidadMezclar.Text := '';

  txtOrdenMezclar.SetFocus;
end;

procedure TfrmVentas.CopiarOrden1Click(Sender: TObject);
begin
  Clipboard.AsText := gvProdNumero.Cells[0, gvProdNumero.SelectedRow];
end;

procedure TfrmVentas.Copiar1Click(Sender: TObject);
begin
  Clipboard.AsText := gvProdPlano.Cells[0, gvProdPlano.SelectedRow];
end;

procedure TfrmVentas.MenuItem2Click(Sender: TObject);
begin
  if (giOpcion = 0) or (not chkMezclar.Checked) then Exit;

  txtOrdenMezclar.Text := gvProdNumero.Cells[0, gvProdNumero.SelectedRow];
  txtCantidadMezclar.Text := gvProdNumero.Cells[3, gvProdNumero.SelectedRow];
  txtCantidadMezclar.SetFocus;
end;

procedure TfrmVentas.Mezclar1Click(Sender: TObject);
begin
  if (giOpcion = 0) or (not chkMezclar.Checked) then Exit;

  txtOrdenMezclar.Text := gvProdPlano.Cells[0, gvProdPlano.SelectedRow];
  txtCantidadMezclar.Text := gvProdPlano.Cells[3, gvProdPlano.SelectedRow];
  txtCantidadMezclar.SetFocus;
end;

procedure TfrmVentas.InsertOrdenesMezclar();
var i : Integer;
SQLStr : String;
sDate: String;
begin
  if (giOpcion = 1) or (giOpcion = 2) then begin
      sDate := DateTimeToStr(Now);
      for i:= 0 to gvMezclado.RowCount - 1 do
      begin
            SQLStr := 'INSERT INTO tblMergeOrdenes(ITE_Nombre, MO_ITE_Nombre, MO_Cantidad, Update_Date, Update_User) ' +
                      'VALUES(' + QuotedStr(gsYear + txtOrden.Text) + ',' + QuotedStr(gsYear + gvMezclado.Cells[0,i]) +
                      ',' + gvMezclado.Cells[1,i] + ',' + QuotedStr(sDate) + ',' + frmMain.sUserLogin +')';

            conn.Execute(SQLStr);


      end;
  end;
end;

procedure TfrmVentas.BindOrdenesMezclar();
var Qry2 : TADOQuery;
SQLStr : String;
begin
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT * FROM tblMergeOrdenes WHERE ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvMezclado.ClearRows;
  while not Qry2.Eof do
  begin
    gvMezclado.AddRow(1);
    gvMezclado.Cells[0,gvMezclado.RowCount -1] := RightStr( VarToStr(Qry2['MO_ITE_Nombre']), Length(VarToStr(Qry2['MO_ITE_Nombre']))-3 );
    gvMezclado.Cells[1,gvMezclado.RowCount -1] := VarToStr(Qry2['MO_Cantidad']);

    Qry2.Next;
  end;

  Qry2.Close;
  Qry2.Free;
end;

end.
