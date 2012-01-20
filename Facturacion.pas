unit Facturacion;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors, ScrollView, CustomGridViewControl,Chris_Functions,
  CustomGridView, GridView,ADODB,DB,ComObj, ExtCtrls,LTCUtils,Larco_functions,Math,
  Menus,Clipbrd,sndkey32,All_Functions,QPrinters,IniFiles;

type
  TfrmFacturacion = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    ddlCliente: TComboBox;
    ddlOrdenCompra: TComboBox;
    Label4: TLabel;
    ddlOrden: TComboBox;
    txtFolio: TEdit;
    Label5: TLabel;
    deFecha: TDateEditor;
    GroupBox2: TGroupBox;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    txtCliente: TEdit;
    txtRFC: TEdit;
    txtDireccion: TEdit;
    txtTelefono: TEdit;
    txtCiudad: TEdit;
    txtOrden: TEdit;
    gvFacturacion: TGridView;
    GroupBox3: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    txtSubtotal: TEdit;
    txtIVA: TEdit;
    txtTotal: TEdit;
    txtLugar: TEdit;
    txtPedimento: TEdit;
    txtLetra: TEdit;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    GroupBox4: TGroupBox;
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button4: TButton;
    Button3: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    lblId: TLabel;
    btnCliente: TButton;
    lblCliente: TLabel;
    GroupBox5: TGroupBox;
    gvOrdenes: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    GroupBox6: TGroupBox;
    gvTrabajo: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    Label19: TLabel;
    txtTipoCambio: TEdit;
    chkDolares: TCheckBox;
    Label20: TLabel;
    gbNormales: TGroupBox;
    Label2: TLabel;
    txtOrdenCompra: TEdit;
    btnOrden: TButton;
    Label3: TLabel;
    txtOrdenes: TEdit;
    btnOrdenes: TButton;
    btnAdd: TButton;
    btnDelete: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    CopiarOrden1: TMenuItem;
    SaveDialog1: TSaveDialog;
    CopiarComo1: TMenuItem;
    Separadoporcomas1: TMenuItem;
    Encomillas1: TMenuItem;
    GroupBox8: TGroupBox;
    gvClientes: TGridView;
    Button5: TButton;
    Label21: TLabel;
    txtColonia: TEdit;
    Imprimir: TButton;
    lblAnio: TLabel;
    btnClear: TButton;
    Label22: TLabel;
    txtReq: TEdit;
    btnExternas: TButton;
    gbExternas: TGroupBox;
    Label23: TLabel;
    Label24: TLabel;
    btnNormales: TButton;
    txtCantidad: TEdit;
    txtDesc: TEdit;
    Label25: TLabel;
    txtNumero: TEdit;
    Label26: TLabel;
    txtUnitario: TEdit;
    chkDllsExt: TCheckBox;
    btnAddExternas: TButton;
    ddlIVAFactura: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindClients();
    procedure BindClientData(Clave: String; ClienteId: String);
    procedure BindOrdenesCompra();
    procedure BindOrdenesTrabajo();
    procedure BindData(Detail:Boolean);
    procedure BindDetail(Factura: String);
    procedure BindFacturaDetalle(Factura: String);
    procedure ClearData();
    procedure ClearClientData();
    procedure EnableControls(Value:Boolean);
    procedure ddlClienteDropDown(Sender: TObject);
    procedure ddlOrdenCompraDropDown(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure ddlClienteChange(Sender: TObject);
    procedure btnOrdenClick(Sender: TObject);
    procedure gbButtonsClick(Sender: TObject);
    procedure FormMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnOK2Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
      Grid:TGridView; Button: TButton);
    procedure btnOrdenesClick(Sender: TObject);
    procedure txtOrdenCompraChange(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure txtTipoCambioKeyPress(Sender: TObject; var Key: Char);
    procedure txtTipoCambioChange(Sender: TObject);
    procedure chkDolaresClick(Sender: TObject);
    //procedure txtIVAFacturaKeyPress(Sender: TObject; var Key: Char);
    procedure calcularTotal();
    //procedure txtIVAFacturaChange(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateFolio(Folio : String):Boolean;
    function ValidateData():Boolean;
    function BoolToStrInt(Value:Boolean):String;
    procedure ActualizarOrdenes(factura: String);
    procedure BuscarClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure gbNormalesClick(Sender: TObject);
    procedure GroupBox2Click(Sender: TObject);
    function getOrdenesAgregadas():String;
    procedure recalculate();
    procedure Exportar1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure Encomillas1Click(Sender: TObject);
    procedure Separadoporcomas1Click(Sender: TObject);
    procedure btnClienteClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure txtOrdenCompraExit(Sender: TObject);
    procedure gvFacturacionAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: String; var Accept: Boolean);
    procedure btnClearClick(Sender: TObject);
    procedure Split(const Delimiter: Char; Input: string; const Strings: TStrings);
    procedure btnExternasClick(Sender: TObject);
    procedure btnNormalesClick(Sender: TObject);
    procedure btnAddExternasClick(Sender: TObject);
    procedure ddlIVAFacturaChange(Sender: TObject);
    procedure ddlIVAFacturaKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmFacturacion: TfrmFacturacion;
  Conn : TADOConnection;
  Qry : TADOQuery;
  giOpcion : Integer;
  sPermits, gsPrinter : String;  
implementation

uses Main, ReporteFactura;

{$R *.dfm}

procedure TfrmFacturacion.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmFacturacion.BindClients();
var SQLStr : String;
Qry2 : TADOQuery;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT DISTINCT Clave FROM tblClientes ORDER BY Clave';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        ddlCliente.Clear;
        While not Qry2.Eof do
        begin
            ddlCliente.Items.Add(VarToStr(Qry2['Clave']));
            Qry2.Next;
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindClients');
    end;

    Qry2.Close;
end;

procedure TfrmFacturacion.BindOrdenesCompra();
var SQLStr, ordenes : String;
Qry2 : TADOQuery;
begin
    if ddlCliente.Text = '' then begin
        ShowMessage('Por favor selecciona un cliente primero.');
        Exit;
    end;

    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;
        ordenes := '';
        if gvFacturacion.RowCount > 0 then ordenes :=  getOrdenesAgregadas();

        SQLStr := 'Facturas_Ordenes_Compra ' +  QuotedStr(RightStr(lblAnio.Caption,2)) + ',' +
                  QuotedStr(ddlCliente.Text) +  ',' +  QuotedStr(ordenes);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        gvOrdenes.ClearRows;
        While not Qry2.Eof do
        begin
            gvOrdenes.AddRow(1);
            gvOrdenes.Cells[0,gvOrdenes.RowCount -1] := VarToStr(Qry2['OrdenCompra']);
            Qry2.Next;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindOrdenesCompra');
    end;

    Qry2.Close;
end;

procedure TfrmFacturacion.BindOrdenesTrabajo();
var SQLStr,ordenes : String;
Qry2 : TADOQuery;
begin
    if ddlCliente.Text = '' then begin
        ShowMessage('Selecciona un cliente primero');
        Exit;
    end;

    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        ordenes := '';
        if gvFacturacion.RowCount > 0 then ordenes :=  getOrdenesAgregadas();

        SQLStr := 'Facturas_Ordenes_Trabajo ' +  QuotedStr(RightStr(lblAnio.Caption,2)) + ',' +
                  QuotedStr(ddlCliente.Text) +  ',' + QuotedStr(txtOrdenCompra.Text) + ',' +
                  QuotedStr(ordenes);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        gvTrabajo.ClearRows;
        While not Qry2.Eof do
        begin
            gvTrabajo.AddRow(1);
            gvTrabajo.Cells[0,gvTrabajo.RowCount -1] := VarToStr(Qry2['ITE_Nombre']);
            gvTrabajo.Cells[2,gvTrabajo.RowCount -1] := VarToStr(Qry2['FD_Cantidad']);
            gvTrabajo.Cells[3,gvTrabajo.RowCount -1] := VarToStr(Qry2['FD_Desc']);
            gvTrabajo.Cells[4,gvTrabajo.RowCount -1] := VarToStr(Qry2['FD_Numero']);
            gvTrabajo.Cells[5,gvTrabajo.RowCount -1] := VarToStr(Qry2['DllText']);
            gvTrabajo.Cells[6,gvTrabajo.RowCount -1] := VarToStr(Qry2['Unitario']);
            gvTrabajo.Cells[7,gvTrabajo.RowCount -1] := VarToStr(Qry2['Stock']);
            Qry2.Next;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindOrdenesTrabajo');
    end;

    Qry2.Close;
end;

procedure TfrmFacturacion.ddlClienteDropDown(Sender: TObject);
var clave: String;
begin
    clave := ddlCliente.Text;
    BindClients();
    ddlCliente.Text := clave;
end;

procedure TfrmFacturacion.ddlOrdenCompraDropDown(Sender: TObject);
begin
    BindOrdenesCompra();
end;

procedure TfrmFacturacion.FormCreate(Sender: TObject);
var IniFile: TIniFile;
StartDDir: String;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    //ShowMessage(StartDDir);
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    gsPrinter := IniFile.ReadString('System','Printer','Epson');


    lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);

    SetRoundMode(rmUp);
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT * FROM tblFacturas WHERE YEAR(FAC_Fecha) = ' +
                    QuotedStr(lblAnio.Caption) + ' ORDER BY FAC_Folio';
    Qry.Open;
    if Qry.RecordCount > 0 then
        BindData(true)
    else
      begin
          Editar.Enabled := False;
          Borrar.Enabled := False;
          Buscar.Enabled := False;
          Imprimir.Enabled := False;
          btnCancelar.Enabled := False;
      end;
    ddlClienteDropDown(nil);
    giOpcion := 0;

  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmFacturacion.BindData(Detail:Boolean);
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['FAC_Id']);
    lblCliente.Caption := VarToStr(Qry['Cliente_Id']);
    txtFolio.Text := VarToStr(Qry['FAC_Folio']);
    deFecha.Text := VarToStr(Qry['FAC_Fecha']);
    if Detail =  true then
            gvFacturacion.ClearRows;

    chkDolares.Checked := StrToBool(VarToStr(Qry['FAC_Dolares']));
    txtTipoCambio.Text := VarToStr(Qry['FAC_TipoCambio']);
    //txtIVAFactura.Text := VarToStr(Qry['FAC_IVA']);
    ddlIVAFactura.Text := VarToStr(Qry['FAC_IVA']);
    txtLugar.Text := VarToStr(Qry['FAC_Expedicion']);
    txtPedimento.Text := VarToStr(Qry['FAC_Pedimento']);
    txtOrden.Text := VarToStr(Qry['FAC_OrdenCompra']);
    txtReq.Text := VarToStr(Qry['FAC_Req']);

    BindClientData('', lblCliente.Caption);

    if Detail =  true then
      BindFacturaDetalle(lblId.Caption);

end;

procedure TfrmFacturacion.ClearData();
begin
    GroupBox5.Visible := False;
    GroupBox6.Visible := False;

    ddlCliente.Text := '';
    txtOrdenCompra.Text := '';
    txtOrdenes.Text := '';
    ddlOrden.Text := '';
    txtFolio.Text := '';
    deFecha.Text := DateToStr(Now);

    ClearClientData();

    gvFacturacion.ClearRows;

    txtLugar.Text := '';
    txtPedimento.Text := '';
    txtOrden.Text := '';
    txtReq.Text := '';

    txtLetra.Text := '';
    txtSubtotal.Text := '';
    txtIVA.Text := '';
    txtTotal.Text := '';

    chkDolares.Checked := False;
    txtTipoCambio.Text := '1.00';
    //txtIVAFactura.Text := '10';
    ddlIVAFactura.Text := '11';
end;

procedure TfrmFacturacion.EnableControls(Value:Boolean);
begin
    ddlCliente.Enabled := not Value;
    btnAdd.Enabled := not Value;
    btnDelete.Enabled := not Value;    
    ddlOrden.Enabled := not Value;
    txtFolio.ReadOnly := Value;
    deFecha.Enabled := not Value;

    txtLugar.ReadOnly := Value;
    txtPedimento.ReadOnly := Value;
    txtOrden.ReadOnly := Value;
    txtReq.ReadOnly := Value;


    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
    btnCliente.Enabled := not Value;
    chkDolares.Enabled := not Value;
    txtTipoCambio.ReadOnly := Value;
    //txtIVAFactura.ReadOnly := Value;
    ddlIVAFactura.Enabled := not Value;

    btnClear.Enabled := not Value;
    btnExternas.Enabled := not Value;    
    gvFacturacion.Enabled := not Value;
    gbExternas.Visible := False;
    gbNormales.Visible := True;
end;


procedure TfrmFacturacion.Button1Click(Sender: TObject);
begin
    if Qry.RecordCount = 0 then
            Exit;

    Qry.First;
    BindData(true);
    btnAceptar.Enabled := False;
    btnCancelar.Enabled := False;

    Nuevo.Enabled := True;
    Editar.Enabled := True;
    Borrar.Enabled := True;
    Buscar.Enabled := True;
    Imprimir.Enabled := True;

  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmFacturacion.Button2Click(Sender: TObject);
begin
    if Qry.RecordCount = 0 then
            Exit;

    Qry.Prior;
    BindData(true);
    btnAceptar.Enabled := False;
    btnCancelar.Enabled := False;

    Nuevo.Enabled := True;
    Editar.Enabled := True;
    Borrar.Enabled := True;
    Buscar.Enabled := True;
    Imprimir.Enabled := True;

    EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmFacturacion.Button3Click(Sender: TObject);
begin
    if Qry.RecordCount = 0 then
            Exit;

    Qry.Next;
    BindData(true);
    btnAceptar.Enabled := False;
    btnCancelar.Enabled := False;

    Nuevo.Enabled := True;
    Editar.Enabled := True;
    Borrar.Enabled := True;
    Buscar.Enabled := True;
    Imprimir.Enabled := True;

    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmFacturacion.Button4Click(Sender: TObject);
begin
    if Qry.RecordCount = 0 then
            Exit;

    Qry.Last;
    BindData(true);
    btnAceptar.Enabled := False;
    btnCancelar.Enabled := False;

    Nuevo.Enabled := True;
    Editar.Enabled := True;
    Borrar.Enabled := True;
    Buscar.Enabled := True;
    Imprimir.Enabled := True;
    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmFacturacion.btnCancelarClick(Sender: TObject);
begin
    ClearData();
    EnableControls(True);
    btnOrden.Enabled := not True;
    btnOrdenes.Enabled := not True;
    Nuevo.Enabled := True;
    if Qry.RecordCount > 0 Then
    begin
          Editar.Enabled := True;
          Borrar.Enabled := True;
          Buscar.Enabled := True;
          Imprimir.Enabled := True;          
    end;
    BindData(true);
    giOpcion := 0;
    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmFacturacion.NuevoClick(Sender: TObject);
var SQLStr : String;
Qry2 : TADOQuery;
begin
    ClearData();
    EnableControls(False);
    btnOrden.Enabled := False;
    btnOrdenes.Enabled := False;
    giOpcion := 1;

    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    Imprimir.Enabled := False;
    ddlCliente.SetFocus;

    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT MAX(FAC_Folio) + 1 AS FAC_Folio FROM tblFacturas';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;
        txtFolio.Text := VarToStr(Qry2['FAC_Folio']);
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : NuevoClick');
    end;
    Qry2.Close;
end;

procedure TfrmFacturacion.EditarClick(Sender: TObject);
begin
    giOpcion := 2;
    EnableControls(False);
    btnOrden.Enabled := true;
    btnOrdenes.Enabled := true;

    Nuevo.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    Imprimir.Enabled := False;    
    ddlCliente.SetFocus;
end;

procedure TfrmFacturacion.BorrarClick(Sender: TObject);
begin
    giOpcion := 3;
    btnAceptar.Enabled := True;
    btnCancelar.Enabled := True;

    Editar.Enabled := False;
    Nuevo.Enabled := False;
    Buscar.Enabled := False;
    Imprimir.Enabled := False;    
end;

procedure TfrmFacturacion.ddlClienteChange(Sender: TObject);
begin
    BindClientData(ddlCliente.Text, '');

    btnOrden.Enabled := false;
    btnOrdenes.Enabled := false;
    btnAdd.Enabled := false;
    btnDelete.Enabled := false;
    if ddlCliente.Text <> '' then begin
        btnOrden.Enabled := true;
        btnOrdenes.Enabled := true;
        btnAdd.Enabled := true;
        btnDelete.Enabled := true;
        txtOrdenCompra.Text := 'Todos';
        txtOrdenes.Text := 'Todos';
    end;
end;

procedure  TfrmFacturacion.BindClientData(Clave: String; ClienteId: String);
var SQLStr : String;
Qry2 : TADOQuery;
begin
    btnCliente.Enabled := False;
    if ( (Clave = '') and (ClienteId = '')) then begin
        ClearClientData();
        Exit;
    end;

    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblClientes WHERE Id = ' + ClienteId;
        if '' <> Clave then
            SQLStr := 'SELECT * FROM tblClientes WHERE Clave = ' + QuotedStr(Clave);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        ClearClientData();
        ddlCliente.Text := VarToStr(Qry2['Clave']);
        txtCliente.Text := VarToStr(Qry2['Nombre']);
        txtRFC.Text := VarToStr(Qry2['RFC']);
        txtDireccion.Text := VarToStr(Qry2['Calle']) + ' NO ' + VarToStr(Qry2['Numero']);
        txtCiudad.Text := VarToStr(Qry2['Ciudad']) + '  ' + VarToStr(Qry2['Estado']) + '  ' +
                          VarToStr(Qry2['CP']);
        txtTelefono.Text := VarToStr(Qry2['Telefono']);
        //txtOrden.Text := VarToStr(Qry['FAC_Fecha']);
        lblCliente.Caption := VarToStr(Qry2['Id']);
        txtColonia.Text := VarToStr(Qry2['Colonia']);
        
        if Qry2.RecordCount > 1 Then
        begin
            btnCliente.Enabled := True;
            gvClientes.ClearRows;
            While not Qry2.Eof do
            begin
                gvClientes.AddRow(1);
                gvClientes.Cells[0,gvClientes.RowCount -1] := VarToStr(Qry2['Id']);
                gvClientes.Cells[1,gvClientes.RowCount -1] := VarToStr(Qry2['Clave']);
                gvClientes.Cells[2,gvClientes.RowCount -1] := VarToStr(Qry2['Nombre']);
                gvClientes.Cells[3,gvClientes.RowCount -1] := VarToStr(Qry2['RFC']);
                Qry2.Next;
            end;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindClients');
    end;

    Qry2.Close;
end;

procedure TfrmFacturacion.ClearClientData();
begin
    txtCliente.Text := '';
    txtRFC.Text := '';
    txtDireccion.Text := '';
    txtCiudad.Text := '';
    txtTelefono.Text := '';
end;

procedure TfrmFacturacion.btnOrdenClick(Sender: TObject);
begin
  GroupBox6.Visible := False;
  ShowSeleccionGrid(GroupBox5, CheckBox2, txtOrdenCompra, gvOrdenes, btnTodos2);
  if GroupBox5.Visible = True then
      BindOrdenesCompra();
end;

procedure TfrmFacturacion.gbButtonsClick(Sender: TObject);
begin
  GroupBox5.Visible := False;
  GroupBox6.Visible := False;
end;

procedure TfrmFacturacion.FormMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  GroupBox5.Visible := False;
  GroupBox6.Visible := False;
end;

procedure TfrmFacturacion.btnOK2Click(Sender: TObject);
begin
  ParseSelection(GroupBox5,CheckBox2,txtOrdenCompra,gvOrdenes);
end;

procedure TfrmFacturacion.CheckBox2Click(Sender: TObject);
begin
gvOrdenes.Enabled := not CheckBox2.Checked;
btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmFacturacion.btnTodos2Click(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodos2, gvOrdenes);
end;

procedure TfrmFacturacion.CheckBox1Click(Sender: TObject);
begin
gvTrabajo.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmFacturacion.btnTodosClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodos, gvTrabajo);
end;

procedure TfrmFacturacion.btnOKClick(Sender: TObject);
begin
  ParseSelection(GroupBox6,CheckBox1,txtOrdenes,gvTrabajo);
end;

procedure TfrmFacturacion.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
                                         Grid: TGridView);
var i: integer;
sOrdenes : String;
begin
  GroupBox.Visible := False;
  if CheckBox.Checked = True then begin
          TextBox.Text := 'Todos';
  end
  else begin
        sOrdenes := '';
        for i:= 0 to Grid.RowCount - 1 do
        begin
                if Grid.Cell[1,i].AsBoolean = True then
                begin
                        sOrdenes := sOrdenes + Grid.Cells[0,i] + ',';
                end;
        end;
        TextBox.Text := 'Todos';
        if sOrdenes <> '' then
        begin
                TextBox.Text :=  LeftStr(sOrdenes,Length(sOrdenes) - 1);
        end;
  end;

end;

procedure TfrmFacturacion.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
var i: integer;
begin
  if UT(Button.Caption) = UT('Seleccionar Todos') then begin
        Button.Caption := 'Deseleccionar Todos';
        for i:= 0 to Grid.RowCount - 1 do
        begin
                Grid.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        Button.Caption := 'Seleccionar Todos';
        for i:= 0 to Grid.RowCount - 1 do
        begin
                Grid.Cell[1,i].AsBoolean := False;
        end;
  end;


end;

procedure TfrmFacturacion.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
          TextBox: TEdit; Grid: TGridView; Button: TButton);
begin
  if GroupBox.Visible = True then
  begin
          GroupBox.Visible := False;
  end
  else begin
      GroupBox.Width := TextBox.Width;
      GroupBox.Top := gbNormales.Top + TextBox.Top + TextBox.Height;
      GroupBox.Left := TextBox.Left + 8;
      Grid.Width := GroupBox.Width - 12;
      Button.Width := GroupBox.Width - 12;

      GroupBox.Visible := True;
      CheckBox.Checked := False;
      Grid.Enabled := True;
      Button.Enabled := True;
  end;

end;

procedure TfrmFacturacion.btnOrdenesClick(Sender: TObject);
begin
  GroupBox5.Visible := False;
  ShowSeleccionGrid(GroupBox6, CheckBox1, txtOrdenes, gvTrabajo, btnTodos);
  if GroupBox6.Visible = True then
     BindOrdenesTrabajo();
end;

procedure TfrmFacturacion.txtOrdenCompraChange(Sender: TObject);
begin
txtOrdenes.Text := 'Todos';
end;

procedure TfrmFacturacion.btnAddClick(Sender: TObject);
var sUnitario: String;
dTipoCambio, dUnitario : Double;
i: integer;
bOrden: boolean;
SQLStr : String;
Qry2 : TADOQuery;
begin
    if nil <> Sender then
    begin
        if gvFacturacion.RowCount = 25 then begin
            ShowMessage('Esta factura contiene 25 ordenes no puedes agregar mas.');
            Exit;
        end;
    end;

    if ddlCliente.Text = '' then begin
        ShowMessage('Selecciona un cliente primero');
        Exit;
    end;

    if ((gvTrabajo.RowCount = 0) or (txtOrdenes.Text = 'Todos') or (txtOrdenes.Text = ''))then
        BindOrdenesTrabajo();

    if UT(txtTipoCambio.Text) = '' then txtTipoCambio.Text := '1.00';

    bOrden := false;
    if ((gvFacturacion.RowCount = 0) and (txtOrden.Text = '')) then
        bOrden := true;

    dTipoCambio := StrToFloat(txtTipoCambio.Text);
    if dTipoCambio <= 0 then dTipoCambio := 1.00;

    for i:= 0 to gvTrabajo.RowCount - 1 do
    begin
        if ((txtOrdenes.Text = '') or (txtOrdenes.Text = 'Todos') or (gvTrabajo.Cell[1,i].AsBoolean = True )) then
        begin
            gvFacturacion.AddRow(1);
            gvFacturacion.Cells[0,gvFacturacion.RowCount -1] := gvTrabajo.Cells[0,i];//ITE_Nombre
            gvFacturacion.Cells[1,gvFacturacion.RowCount -1] := gvTrabajo.Cells[2,i];//FD_Cantidad
            gvFacturacion.Cells[2,gvFacturacion.RowCount -1] := gvTrabajo.Cells[3,i];//FD_Desc
            gvFacturacion.Cells[3,gvFacturacion.RowCount -1] := gvTrabajo.Cells[4,i];//FD_Numero
            gvFacturacion.Cells[5,gvFacturacion.RowCount -1] := gvTrabajo.Cells[5,i];//Dolares

            sUnitario := gvTrabajo.Cells[6,i]; //Unitario

            if UT(sUnitario) = '' then sUnitario := '0.00';
            dUnitario := StrToFloat(sUnitario);
            if chkDolares.Checked then //Si la factura es en dolares
            begin
                if UT(gvTrabajo.Cells[5,i]) = 'N' then // y la cantidad capturada esta en pesos
                begin
                      dUnitario := RoundTo(dUnitario / dTipoCambio, -2);
                end;
            end
            else begin // si la factura es en pesos
                if UT(gvTrabajo.Cells[5,i]) = 'Y' then
                begin
                      dUnitario := RoundTo(dUnitario * dTipoCambio, -2);
                end;
            end;
            gvFacturacion.Cells[4,gvFacturacion.RowCount -1] := FormatFloat('#,###,##0.00', dUnitario);
            gvFacturacion.Cells[6,gvFacturacion.RowCount -1] :=
                    FormatFloat('#,###,##0.00', dUnitario * StrToFloat(gvTrabajo.Cells[2,i]) );

            gvFacturacion.Cells[7,gvFacturacion.RowCount -1] := gvTrabajo.Cells[6,i];
            gvFacturacion.Cells[8,gvFacturacion.RowCount -1] := gvTrabajo.Cells[3,i];
            gvFacturacion.Cells[9,gvFacturacion.RowCount -1] := gvTrabajo.Cells[7,i];
        end;

        if gvFacturacion.RowCount = 25 then Break;
    end;

    calcularTotal();

    //Cargar el numero de orden de compra de la primera orden.
    if bOrden = true then begin
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT OrdenCompra FROM tblOrdenes WHERE ITE_Nombre = ' +
                   QuotedStr(gvFacturacion.Cells[0,0]);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        txtOrden.Text := VarToStr(Qry2['OrdenCompra']);

        Qry2.Close;
    end;

end;

procedure TfrmFacturacion.txtTipoCambioKeyPress(Sender: TObject;
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

procedure TfrmFacturacion.txtTipoCambioChange(Sender: TObject);
begin
    if (Sender as TEdit).Text = '' then
      (Sender as TEdit).Text := '1.00'
    else
      (Sender as TEdit).Text := FormatFloat('######0.00', StrToFloat((Sender as TEdit).Text) );

    if (Sender as TEdit).Text = '0.00' then
      (Sender as TEdit).Text := '1.00';

    if gvFacturacion.RowCount > 0 then
        recalculate();
end;

procedure TfrmFacturacion.chkDolaresClick(Sender: TObject);
begin
    if gvFacturacion.RowCount > 0 then
        recalculate();

end;

procedure TfrmFacturacion.ddlIVAFacturaKeyPress(Sender: TObject;
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

{procedure TfrmFacturacion.txtIVAFacturaKeyPress(Sender: TObject;
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
end;}

procedure TfrmFacturacion.calcularTotal();
var sImporte: String;
dImporte, dIVA: Double;
i: Integer;
begin
        if gvFacturacion.RowCount <= 0 then
          Exit;

        dImporte := 0.00;
        for i:= 0 to gvFacturacion.RowCount - 1 do
        begin
                sImporte := StringReplace(gvFacturacion.Cells[6,i],',','',[rfReplaceAll, rfIgnoreCase]);
                dImporte := dImporte + StrToFloat(sImporte);
        end;

        txtSubTotal.Text := FormatFloat('###,##0.00', dImporte);

        //if txtIVAFactura.Text = '' then
        //   txtIVAFactura.Text := '10';

        dIVA := StrToFloat(ddlIVAFactura.Text);
        txtIVA.Text := FormatFloat('###,##0.00', RoundTo(dImporte * (dIVA / 100),-2 ));

        txtTotal.Text := FormatFloat('###,##0.00', dImporte + ( RoundTo(dImporte * (dIVA / 100),-2 ) ));
        txtLetra.Text := NumLetra(dImporte + ( RoundTo(dImporte * (dIVA / 100),-2 ) ) ,1,1);

        if StrPos(PChar(txtLetra.Text), PChar('pesos')) = nil then
          txtLetra.Text := txtLetra.Text + ' pesos 00/100 M.N.'
        else
          begin
              txtLetra.Text := LeftStr(txtLetra.Text,Pos('pesos', txtLetra.Text) - 2) + ' pesos ' +
                               RightStr(txtTotal.Text,2) + '/100 M.N.';
          end;

        txtLetra.Text := UpperCase( txtLetra.Text);
        if chkDolares.Checked then
        begin
            txtLetra.Text := StringReplace(txtLetra.Text,'PESOS','DOLARES',[rfReplaceAll, rfIgnoreCase]);
            txtLetra.Text := StringReplace(txtLetra.Text,'M.N.','M.A.',[rfReplaceAll, rfIgnoreCase]);
        end;

end;

{procedure TfrmFacturacion.txtIVAFacturaChange(Sender: TObject);
begin
calcularTotal();
end;}

procedure TfrmFacturacion.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['Cliente_Id'] := lblCliente.Caption;
        Qry['FAC_Folio'] := txtFolio.Text;
        Qry['FAC_Fecha'] := deFecha.Text;
        Qry['FAC_Expedicion'] := txtLugar.Text;
        Qry['FAC_Pedimento'] := txtPedimento.Text;
        Qry['FAC_Dolares'] := BoolToStrInt(chkDolares.Checked);
        Qry['FAC_TipoCambio'] := txtTipoCambio.Text;
        //Qry['FAC_IVA'] := txtIVAFactura.Text;
        Qry['FAC_IVA'] := ddlIVAFactura.Text;
        Qry['FAC_OrdenCompra'] := txtOrden.Text;
        Qry['FAC_Req'] := txtReq.Text;
        Qry.Post;

        BindData(false);
        ActualizarOrdenes(lblId.Caption);
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;
        Qry.Edit;
        Qry['Cliente_Id'] := lblCliente.Caption;
        Qry['FAC_Folio'] := txtFolio.Text;
        Qry['FAC_Fecha'] := deFecha.Text;
        Qry['FAC_Expedicion'] := txtLugar.Text;
        Qry['FAC_Pedimento'] := txtPedimento.Text;
        Qry['FAC_Dolares'] := BoolToStrInt(chkDolares.Checked);
        Qry['FAC_TipoCambio'] := txtTipoCambio.Text;
        //Qry['FAC_IVA'] := txtIVAFactura.Text;
        Qry['FAC_IVA'] := ddlIVAFactura.Text;
        Qry['FAC_OrdenCompra'] := txtOrden.Text;
        Qry['FAC_Req'] := txtReq.Text;
        Qry.Post;

        BindData(false);
        ActualizarOrdenes(lblId.Caption);
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar la factura con folio: ' +
                      txtFolio.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              Qry.Delete;
              gvFacturacion.ClearRows;
              ActualizarOrdenes(lblId.Caption);
        end;
  end
  else if giOpcion = 4 then
  begin
        ShowMessage('Esta opcion no esta disponible por el momento.');
  end;


  ClearData();
  EnableControls(True);
  btnOrden.Enabled := not True;
  btnOrdenes.Enabled := not True;

  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Imprimir.Enabled := False;
  Nuevo.Enabled := True;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
        Imprimir.Enabled := True;
  end;
  BindData(true);
  giOpcion := 0;
  EnableFormButtons(gbButtons, sPermits);
end;

function TfrmFacturacion.ValidateFolio(Folio : String):Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Result := False;
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT FAC_Folio FROM tblFacturas Where FAC_Folio = ' + Folio;

    if giOpcion = 2 then
        SQLStr :=  SQLStr + ' AND FAC_Id <> ' + lblId.Caption;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then
        Result := True;

    Qry2.Close;
end;

function TfrmFacturacion.ValidateData():Boolean;
var id : Integer;
begin
        Result := False;
        if txtFolio.Text = '' then
        begin
                ShowMessage('El folio no puede estar vacio.');
                Exit;
        end;

        if ddlCliente.Text = '' then
        begin
                ShowMessage('Por favor seleccione un cliente.');
                Exit;
        end;

        id := ddlCliente.Items.IndexOf(ddlCliente.Text);
        if id = -1 then
        begin
                ShowMessage('Por favor seleccione un cliente valido de la lista.(No lo escriba)');
                Exit;
        end;

        if ddlIVAFactura.Text = '' then
        begin
                ShowMessage('Por favor seleccione el IVA.');
                Exit;
        end;

        id := ddlIVAFactura.Items.IndexOf(ddlIVAFactura.Text);
        if id = -1 then
        begin
                ShowMessage('Por favor seleccione un IVA valido de la lista.(No lo escriba)');
                Exit;
        end;


        if ValidateFolio(txtFolio.Text) then
        begin
                ShowMessage('Ya existe una factura con este numero de folio.');
                Exit;
        end;

        if gvFacturacion.RowCount <= 0 then
        begin
                ShowMessage('Por favor agrege al menos una orden a esta factura.');
                Exit;
        end;

        Result := True;
end;

function TfrmFacturacion.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;

procedure TfrmFacturacion.ActualizarOrdenes(factura: String);
var i : Integer;
SQLStr, stock : String;
begin
  SQLStr := 'DELETE FROM tblFacturasDetalle WHERE FAC_ID = ' + factura;
  conn.Execute(SQLStr);

  for i:= 0 to gvFacturacion.RowCount - 1 do
  begin

      if (gvFacturacion.Cells[9,i] <> '2') then
      begin
        stock := '0';
        if gvFacturacion.Cells[9,i] <> '0' then stock := '1';

        SQLStr := 'INSERT INTO tblFacturasDetalle(FAC_ID,ITE_Nombre,FD_Cantidad,FD_Desc,FD_Numero,FD_Stock) ' +
                  'VALUES(' + factura + ',' + QuotedStr(gvFacturacion.Cells[0,i]) + ',' +
                  gvFacturacion.Cells[1,i] + ',' + QuotedStr(gvFacturacion.Cells[2,i]) + ',' +
                  QuotedStr(gvFacturacion.Cells[3,i]) + ',' + QuotedStr(stock) + ')';
      end
      else begin
        SQLStr := 'INSERT INTO tblFacturasDetalle(FAC_ID,ITE_Nombre,FD_Cantidad,FD_Desc,FD_Numero,FD_Stock,FD_Dolares,FD_Unitario) ' +
                  'VALUES(' + factura + ',' + QuotedStr(gvFacturacion.Cells[0,i]) + ',' +
                  gvFacturacion.Cells[1,i] + ',' + QuotedStr(gvFacturacion.Cells[2,i]) + ',' +
                  QuotedStr(gvFacturacion.Cells[3,i]) + ',' + QuotedStr(gvFacturacion.Cells[9,i]) + ',' +
                  QuotedStr(gvFacturacion.Cells[5,i]) + ',' + gvFacturacion.Cells[4,i] + ')';
      end;

      conn.Execute(SQLStr);
  end;

end;

procedure TfrmFacturacion.BuscarClick(Sender: TObject);
begin
        ShowMessage('Esta opcion no esta disponible por el momento.');
end;

procedure TfrmFacturacion.btnDeleteClick(Sender: TObject);
begin
  gvFacturacion.DeleteRow(gvFacturacion.SelectedRow);
  calcularTotal();
end;

procedure TfrmFacturacion.gbNormalesClick(Sender: TObject);
begin
  GroupBox5.Visible := False;
  GroupBox6.Visible := False;
end;

procedure TfrmFacturacion.GroupBox2Click(Sender: TObject);
begin
  GroupBox5.Visible := False;
  GroupBox6.Visible := False;
end;

function TfrmFacturacion.getOrdenesAgregadas():String;
var sOrdenes:String;
i: Integer;
begin
  sOrdenes := '';
  for i:= 0 to gvFacturacion.RowCount - 1 do
  begin
        sOrdenes := sOrdenes + gvFacturacion.Cells[0,i] + ',';
  end;

  if sOrdenes <> '' then
  begin
          sOrdenes :=  LeftStr(sOrdenes,Length(sOrdenes) - 1);
  end;

  result := sOrdenes;
end;

procedure TfrmFacturacion.BindDetail(Factura:String);
var SQLStr,sUnitario: String;
Qry2 : TADOQuery;
dTipoCambio, dUnitario : Double;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT *,CASE WHEN Dolares = 0 THEN ''N'' ELSE ''Y'' END AS DllText ' +
                  'FROM tblOrdenes WHERE FAC_Id = ' + Factura + ' ORDER BY ITE_Nombre';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        if UT(txtTipoCambio.Text) = '' then
                txtTipoCambio.Text := '1.00';

        dTipoCambio := StrToFloat(txtTipoCambio.Text);
        if dTipoCambio <= 0 then dTipoCambio := 1.00;

        //if ( (Qry2.RecordCount > 12) and (nil <> Sender))then
        //        ShowMessage('Selecciono mas de 12 ordenes, ' + #13 + 'solo se van a agregar las primeras 12 ordenadas por numero de orden.');

        gvFacturacion.ClearRows;
        While not Qry2.Eof do
        begin
            gvFacturacion.AddRow(1);
            gvFacturacion.Cells[0,gvFacturacion.RowCount -1] := VarToStr(Qry2['ITE_Nombre']);
            gvFacturacion.Cells[1,gvFacturacion.RowCount -1] := VarToStr(Qry2['FAC_Cantidad']);
            gvFacturacion.Cells[2,gvFacturacion.RowCount -1] := VarToStr(Qry2['FAC_Desc']);
            gvFacturacion.Cells[3,gvFacturacion.RowCount -1] := VarToStr(Qry2['FAC_Numero']);
            gvFacturacion.Cells[5,gvFacturacion.RowCount -1] := VarToStr(Qry2['DllText']);

            sUnitario := VarToStr(Qry2['Unitario']);

            if UT(sUnitario) = '' then sUnitario := '0.00';
            dUnitario := StrToFloat(sUnitario);
            if chkDolares.Checked then //Si la factura es en dolares
            begin
                if UT(VarToStr(Qry2['DllText'])) = 'N' then // y la cantidad capturada esta en pesos
                begin
                      dUnitario := RoundTo(dUnitario / dTipoCambio, -2);
                end;
            end
            else begin // si la factura es en pesos
                if UT(VarToStr(Qry2['DllText'])) = 'Y' then
                begin
                      dUnitario := RoundTo(dUnitario * dTipoCambio, -2);
                end;
            end;
            gvFacturacion.Cells[4,gvFacturacion.RowCount -1] := FormatFloat('#,###,##0.00', dUnitario);
            gvFacturacion.Cells[6,gvFacturacion.RowCount -1] :=
                    FormatFloat('#,###,##0.00', dUnitario * StrToFloat(VarToStr(Qry2['FAC_Cantidad'])) );

            gvFacturacion.Cells[7,gvFacturacion.RowCount -1] := VarToStr(Qry2['Unitario']);
            gvFacturacion.Cells[8,gvFacturacion.RowCount -1] := VarToStr(Qry2['FAC_Desc']);

            if gvFacturacion.RowCount = 25 then
            begin
                //ShowMessage('Ya se agregaron 12 ordenes no se pueden agregar mas.');
                break;
            end;
            Qry2.Next;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindDetail');
    end;

    Qry2.Close;
    calcularTotal();


end;


procedure TfrmFacturacion.BindFacturaDetalle(Factura: String);
var SQLStr,sUnitario: String;
Qry2 : TADOQuery;
dTipoCambio, dUnitario : Double;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'Facturas_Detalle ' + Factura;

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        if UT(txtTipoCambio.Text) = '' then
                txtTipoCambio.Text := '1.00';

        dTipoCambio := StrToFloat(txtTipoCambio.Text);
        if dTipoCambio <= 0 then dTipoCambio := 1.00;

        gvFacturacion.ClearRows;
        While not Qry2.Eof do
        begin
            gvFacturacion.AddRow(1);
            gvFacturacion.Cells[0,gvFacturacion.RowCount -1] := VarToStr(Qry2['ITE_Nombre']);
            gvFacturacion.Cells[1,gvFacturacion.RowCount -1] := VarToStr(Qry2['FD_Cantidad']);
            gvFacturacion.Cells[2,gvFacturacion.RowCount -1] := VarToStr(Qry2['FD_Desc']);
            gvFacturacion.Cells[3,gvFacturacion.RowCount -1] := VarToStr(Qry2['FD_Numero']);
            gvFacturacion.Cells[5,gvFacturacion.RowCount -1] := VarToStr(Qry2['DllText']);
            gvFacturacion.Cells[9,gvFacturacion.RowCount -1] := VarToStr(Qry2['Stock']);

            sUnitario := VarToStr(Qry2['Unitario']);

            if UT(sUnitario) = '' then sUnitario := '0.00';
            dUnitario := StrToFloat(sUnitario);
            if chkDolares.Checked then //Si la factura es en dolares
            begin
                if UT(VarToStr(Qry2['DllText'])) = 'N' then // y la cantidad capturada esta en pesos
                begin
                      dUnitario := RoundTo(dUnitario / dTipoCambio, -2);
                end;
            end
            else begin // si la factura es en pesos
                if UT(VarToStr(Qry2['DllText'])) = 'Y' then
                begin
                      dUnitario := RoundTo(dUnitario * dTipoCambio, -2);
                end;
            end;
            gvFacturacion.Cells[4,gvFacturacion.RowCount -1] := FormatFloat('#,###,##0.00', dUnitario);
            gvFacturacion.Cells[6,gvFacturacion.RowCount -1] :=
                    FormatFloat('#,###,##0.00', dUnitario * StrToFloat(VarToStr(Qry2['FD_Cantidad'])) );

            gvFacturacion.Cells[7,gvFacturacion.RowCount -1] := VarToStr(Qry2['Unitario']);
            gvFacturacion.Cells[8,gvFacturacion.RowCount -1] := VarToStr(Qry2['FD_Desc']);

            Qry2.Next;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindFacturaDetalle');
    end;

    Qry2.Close;
    calcularTotal();
end;



procedure TfrmFacturacion.recalculate();
var dTipoCambio, dUnitario : Double;
sUnitario : String;
i: integer;
begin
  if UT(txtTipoCambio.Text) = '' then
          txtTipoCambio.Text := '1.00';

  dTipoCambio := StrToFloat(txtTipoCambio.Text);
  if dTipoCambio <= 0 then dTipoCambio := 1.00;


  for i:= 0 to gvFacturacion.RowCount - 1 do
  begin
            sUnitario := StringReplace(gvFacturacion.Cells[7,i],',','',[rfReplaceAll, rfIgnoreCase]);

            if UT(sUnitario) = '' then sUnitario := '0.00';
            dUnitario := StrToFloat(sUnitario);
            if chkDolares.Checked then //Si la factura es en dolares
            begin
                if UT(gvFacturacion.Cells[5,i]) = 'N' then // y la cantidad capturada esta en pesos
                begin
                      dUnitario := RoundTo(dUnitario / dTipoCambio, -2);
                end;
            end
            else begin // si la factura es en pesos
                if UT(gvFacturacion.Cells[5,i]) = 'Y' then
                begin
                      dUnitario := RoundTo(dUnitario * dTipoCambio, -2);
                end;
            end;
            gvFacturacion.Cells[4,i] := FormatFloat('#,###,##0.00', dUnitario);
            gvFacturacion.Cells[6,i] :=
                    FormatFloat('#,###,##0.00', dUnitario * StrToFloat(gvFacturacion.Cells[1,i]) );

  end;
  calcularTotal();

end;


procedure TfrmFacturacion.Exportar1Click(Sender: TObject);
var sFileName: String;
begin

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(gvFacturacion,sFileName);

  end;
end;

procedure TfrmFacturacion.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmFacturacion.CopiarOrden1Click(Sender: TObject);
begin
    Clipboard.AsText := gvFacturacion.Cells[0,gvFacturacion.SelectedRow]
end;

procedure TfrmFacturacion.Encomillas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
        sText := '';
        for i:= 0 to gvFacturacion.RowCount - 1 do
                   sText := sText + QuotedStr(gvFacturacion.Cells[0,i]) + ',';

        Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
end;

procedure TfrmFacturacion.Separadoporcomas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
       sText := '';
       for i:= 0 to gvFacturacion.RowCount - 1 do
               sText := sText + gvFacturacion.Cells[0,i] + ',';

       Clipboard.AsText := LeftStr(sText,Length(sText) - 1);

end;

procedure TfrmFacturacion.btnClienteClick(Sender: TObject);
begin
  if GroupBox8.Visible = True then
  begin
          GroupBox8.Visible := False;
  end
  else begin
      //GroupBox8.Width := TextBox.Width;
      GroupBox8.Top := GroupBox2.Top +  txtCliente.Top + txtCliente.Height;
      GroupBox8.Left := txtCliente.Left + 8;

      GroupBox8.Visible := True;
  end;


end;

procedure TfrmFacturacion.Button5Click(Sender: TObject);
begin
    btnClienteClick(nil);
    BindClientData('',gvClientes.Cells[0,gvClientes.SelectedRow]);
end;

procedure TfrmFacturacion.ImprimirClick(Sender: TObject);
var dataSet: TADODataSet;
i,p: Integer;
desc, printerName: String;
Strings: TStringList;
list: TStrings;
begin
        Strings := TStringList.Create;
        dataSet := TADODataSet.Create(nil);
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Name := 'Cantidad';
        end;
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Size := 120;
          Name := 'Descripcion';
        end;
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Name := 'Numero';
        end;
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Name := 'Unitario';
        end;
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Name := 'Importe';
        end;
        dataSet.CreateDataSet;

        for i:= 0 to gvFacturacion.RowCount - 1 do
        begin
              desc := gvFacturacion.Cells[2,i];
              ShowMessage(IntToStr(InStr(0,desc,'/')));
              if Pos('/', desc) > 0 then begin

                  Strings.Clear;
                  Strings.Delimiter := '/';
                  Strings.DelimitedText := desc;
                  if Strings.count >= 2 then begin
                  dataSet.InsertRecord([gvFacturacion.Cells[1,i], Strings[0],
                                       gvFacturacion.Cells[3,i], gvFacturacion.Cells[4,i],
                                       gvFacturacion.Cells[6,i]]);

                  dataSet.InsertRecord(['', Strings[1],'', '','']);
                  end;
              end
              else begin
              dataSet.InsertRecord([gvFacturacion.Cells[1,i], gvFacturacion.Cells[2,i],
                                   gvFacturacion.Cells[3,i], gvFacturacion.Cells[4,i],
                                   gvFacturacion.Cells[6,i]]);
              end;

        end;

        Application.Initialize;
        Application.CreateForm(TqrFactura,qrFactura);

        list := Printer.Printers;
        for p := 0 to list.Count-1 do begin
            printerName := list[p];
            if Pos(gsPrinter, printerName) <> 0 then begin
                qrFactura.PrinterSettings.PrinterIndex := p;
                break;
            end;
        end;


        qrFactura.lblCliente.Caption := txtCliente.Text;
        qrFactura.lblRFC.Caption := txtRFC.Text;
        qrFactura.lblDireccion.Caption := txtDireccion.Text;
        qrFactura.lblCiudad.Caption := txtCiudad.Text;
        qrFactura.lblColonia.Caption := txtColonia.Text;
        qrFactura.lblTelefono.Caption := txtTelefono.Text;
        qrFactura.lblOrden.Caption := txtOrden.Text;
        qrFactura.lblReq.Caption := '';
        if txtReq.Text <> '' then
                qrFactura.lblReq.Caption := 'REQ. ' + txtReq.Text;

        qrFactura.lblFecha.Caption := FormatDateTime('dd/mm/yyyy', deFecha.Date);

        qrFactura.QRSubDetail1.DataSet := dataSet;
        qrFactura.lblCantidad.DataSet := dataSet;
        qrFactura.lblCantidad.DataField := 'Cantidad';

        qrFactura.lblDescripcion.DataSet := dataSet;
        qrFactura.lblDescripcion.DataField := 'Descripcion';

        qrFactura.lblNumero.DataSet := dataSet;
        qrFactura.lblNumero.DataField := 'Numero';

        qrFactura.lblUnitario.DataSet := dataSet;
        qrFactura.lblUnitario.DataField := 'Unitario';

        qrFactura.lblImporte.DataSet := dataSet;
        qrFactura.lblImporte.DataField := 'Importe';

        qrFactura.lblExpedicion.Caption := txtLugar.Text;
        qrFactura.lblPedimento.Caption := txtPedimento.Text;
        qrFactura.lblLetra.Caption := txtLetra.Text;
        qrFactura.lblSubTotal.Caption := txtSubTotal.Text;
        qrFactura.lblIVA.Caption := txtIVA.Text;
        qrFactura.lblTotal.Caption := txtTotal.Text;


        qrFactura.Preview;
        qrFactura.Free;
end;


procedure TfrmFacturacion.txtOrdenCompraExit(Sender: TObject);
begin
GroupBox5.Visible :=False;
end;

procedure TfrmFacturacion.gvFacturacionAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: String; var Accept: Boolean);
var sUnitario: String;
begin

  if ACol = 4 then
    sUnitario := StringReplace(Value,',','',[rfReplaceAll, rfIgnoreCase]);

  if ((ACol = 1) and (not IsNumeric(Value)) )then
  begin
          ShowMessage('La cantidad debe de ser numerica.');
          Accept := False;
          Exit;
  end;


  if ((ACol = 4) and (not IsNumeric(sUnitario)) )then
  begin
          ShowMessage('El precio unitario debe de ser numerico.');
          Accept := False;
          Exit;
  end;

  if ACol = 4 then gvFacturacion.Cells[7,Arow] := sUnitario;
  if ACol = 1 then begin
        gvFacturacion.Cells[6,ARow] :=
                FormatFloat('#,###,##0.00', StrToFloat(gvFacturacion.Cells[4,ARow]) * StrToFloat(Value) );

        calcularTotal();
  end;

end;

procedure TfrmFacturacion.btnClearClick(Sender: TObject);
begin
  gvFacturacion.ClearRows;
  calcularTotal();
end;

procedure TfrmFacturacion.Split(const Delimiter: Char; Input: string; const Strings: TStrings);
begin
   Assert(Assigned(Strings)) ;
   Strings.Clear;
   Strings.Delimiter := Delimiter;
   Strings.DelimitedText := Input;
end;

procedure TfrmFacturacion.btnExternasClick(Sender: TObject);
begin
gbNormales.Visible := False;
gbExternas.Top :=  gbNormales.Top;
gbExternas.Left :=  gbNormales.Left;
gbExternas.Visible := True;
end;

procedure TfrmFacturacion.btnNormalesClick(Sender: TObject);
begin
gbExternas.Visible := False;
gbNormales.Top :=  gbNormales.Top;
gbNormales.Left :=  gbNormales.Left;
gbNormales.Visible := True;
end;

procedure TfrmFacturacion.btnAddExternasClick(Sender: TObject);
var sUnitario, sDlls: String;
dTipoCambio, dUnitario : Double;
begin
  if txtCantidad.Text = '' then
  begin
          ShowMessage('Por favor captura la cantidad.');
          Exit;
  end;

  if txtDesc.Text = '' then
  begin
          ShowMessage('Por favor captura la Descripcion.');
          Exit;
  end;

  if txtNumero.Text = '' then
  begin
          ShowMessage('Por favor captura el Numero de parte.');
          Exit;
  end;

  if txtUnitario.Text = '' then
  begin
          ShowMessage('Por favor captura el Precio Unitario.');
          Exit;
  end;

  if gvFacturacion.RowCount = 25 then begin
      ShowMessage('Esta factura contiene 25 ordenes no puedes agregar mas.');
      Exit;
  end;

  gvFacturacion.AddRow(1);
  gvFacturacion.Cells[0,gvFacturacion.RowCount -1] := '';//ITE_Nombre
  gvFacturacion.Cells[1,gvFacturacion.RowCount -1] := txtCantidad.Text;//FD_Cantidad
  gvFacturacion.Cells[2,gvFacturacion.RowCount -1] := txtDesc.Text;//FD_Desc
  gvFacturacion.Cells[3,gvFacturacion.RowCount -1] := txtNumero.Text;//FD_Numero
  sDlls := 'N';
  if  chkDllsExt.Checked then sDlls := 'Y';
  gvFacturacion.Cells[5,gvFacturacion.RowCount -1] := sDlls;//Dolares

  sUnitario := txtUnitario.Text; //Unitario
  dTipoCambio := StrToFloat(txtTipoCambio.Text);
  if UT(sUnitario) = '' then sUnitario := '0.00';
  dUnitario := StrToFloat(sUnitario);
  if chkDolares.Checked then //Si la factura es en dolares
  begin
      if UT(sDlls) = 'N' then // y la cantidad capturada esta en pesos
      begin
            dUnitario := RoundTo(dUnitario / dTipoCambio, -2);
      end;
  end
  else begin // si la factura es en pesos
      if UT(sDlls) = 'Y' then
      begin
            dUnitario := RoundTo(dUnitario * dTipoCambio, -2);
      end;
  end;
  gvFacturacion.Cells[4,gvFacturacion.RowCount -1] := FormatFloat('#,###,##0.00', dUnitario);
  gvFacturacion.Cells[6,gvFacturacion.RowCount -1] :=
          FormatFloat('#,###,##0.00', dUnitario * StrToFloat(txtCantidad.Text));

  gvFacturacion.Cells[7,gvFacturacion.RowCount -1] := txtUnitario.Text;
  gvFacturacion.Cells[8,gvFacturacion.RowCount -1] := txtDesc.Text;
  gvFacturacion.Cells[9,gvFacturacion.RowCount -1] := '2';

  calcularTotal();

end;

procedure TfrmFacturacion.ddlIVAFacturaChange(Sender: TObject);
begin
calcularTotal();
end;

end.

