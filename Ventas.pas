unit Ventas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors,ImpresionOrden,Larco_Functions;

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
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
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
    Label17: TLabel;
    lblTarea: TLabel;
    Label18: TLabel;
    lblStatus: TLabel;
    lblStock: TLabel;
    lblAnio: TLabel;
    chkStock: TCheckBox;
    function FormIsRunning(FormName: String):Boolean;
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindProductos();
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

procedure TfrmVentas.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
Button5.Enabled := False;
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
    txtNumero.Text := '';
    txtTerminal.Text := '';
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
end;

procedure TfrmVentas.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
deEntrega.Text := DateToStr(Now);
deInterna.Text := DateToStr(Now);
txtRecibido.Text := DateToStr(Now);
txtOrden.SetFocus;
giOpcion := 1;

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
Button5.Enabled := False;
txtUnitario.Text := '0';
txtTotal.Text := '0';
end;

procedure TfrmVentas.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
Button5.Enabled := False;
txtOrden.SetFocus;
end;

procedure TfrmVentas.BuscarClick(Sender: TObject);
begin
ClearData();
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;
txtOrden.ReadOnly := False;
//EnableControls(False);
txtOrden.SetFocus;
giOpcion := 4;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
Button5.Enabled := False;
end;

procedure TfrmVentas.EnableControls(Value:Boolean);
begin
    txtOrden.ReadOnly := Value;
    txtProceso.ReadOnly := Value;
    txtRequerida.ReadOnly := Value;
    txtNumero.ReadOnly := Value;
    txtTerminal.ReadOnly := Value;
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

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;



procedure TfrmVentas.btnCancelarClick(Sender: TObject);
begin
ClearData();
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
giOpcion := 0;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmVentas.BindOrden();
var SQLStr : String;
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
    txtOrdenada.Text := VarToStr(Qry['Ordenada']);
    txtNumero.Text := VarToStr(Qry['Numero']);
    txtTerminal.Text := VarToStr(Qry['Terminal']);
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

    if Qry2.RecordCount > 0 then
    begin
        lblTarea.Caption := VarToStr(Qry2['Nombre']);
        lblStatus.Caption := VarToStr(Qry2['status']);
    end;

    Qry2.Close;
    Qry2.Free;

end;

procedure TfrmVentas.btnAceptarClick(Sender: TObject);
var SQLStr,sOrden : String;
Qry2 : TADOQuery;
sNew : String;
stock,cambio : boolean;
begin

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        sNew := gsYear + txtOrden.Text;
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
                  ',' + BoolToStrInt(chkStock.Checked);

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

            Qry.Locate('ITE_Nombre', sNew, [loPartialKey] );
            //Qry.Last;

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

        stock := StrToBool(VarToStr(Qry['Stock']));
        if stock <> chkStock.Checked then
                cambio := true;

        sNew := gsYear + txtOrden.Text;
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
                  ',' + BoolToStrInt(cambio) + ',' + BoolToStrInt(stock);

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

            //Qry.Locate('ITE_ID',lblId.Caption,[loPartialKey] )
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

EnableControls(True);
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
if giOpcion = 1 then
      NuevoClick(nil);

giOpcion := 0;
end;

function TfrmVentas.ValidateData():Boolean;
var i:Integer;
bfound : boolean;
begin
        result := True;

{        if not isnumeric(leftstr(txtOrden.Text,2)) then
          begin
            MessageDlg('A�io incorrecto: ' + leftstr(txtOrden.Text,2), mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if (leftstr(txtOrden.Text,2) <>  gsOYear) then
          begin
            MessageDlg('A�io incorrecto: ' + leftstr(txtOrden.Text,2) + ' el a�io no es el mismo que el predeterminado.', mtInformation,[mbOk], 0);
            result :=  False;
          end;
}
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

procedure TfrmVentas.Button2Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Prior;
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
                                      'Nivel de Revisi�n: D' + #13 +
                                      'Retenci�n: 1 a�o+uso';
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

end.