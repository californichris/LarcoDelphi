unit PorcentajeRetrabajoDinero;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors,ADODB,DB,IniFiles,ComObj,All_Functions,StrUtils,chris_Functions,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils,Larco_Functions ;

type
  TfrmRetrabajoDinero = class(TForm)
    GroupBox3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    lblLiberadas2: TLabel;
    lblCLiberadas2: TLabel;
    lblCRetrabajo2: TLabel;
    lblRetrabajo2: TLabel;
    lblTotal2: TLabel;
    lblCTotal2: TLabel;
    GroupBox4: TGroupBox;
    GridView1: TGridView;
    GridView2: TGridView;
    GroupBox2: TGroupBox;
    lblCLiberadas: TLabel;
    lblCRetrabajo: TLabel;
    lblCTotal: TLabel;
    lblLiberadas: TLabel;
    lblRetrabajo: TLabel;
    lblTotal: TLabel;
    lblCPorcentaje: TLabel;
    lblPorcentaje: TLabel;
    Chart1: TChart;
    Series1: TPieSeries;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label9: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label11: TLabel;
    lblTipo: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Detalle: TButton;
    Imprimir: TButton;
    cmbDetectado: TComboBox;
    cmbTareas: TComboBox;
    cmbEmpleados: TComboBox;
    chkDlls: TCheckBox;
    txtCliente: TEdit;
    Button2: TButton;
    txtTipo: TEdit;
    lblDineroLib: TLabel;
    Label5: TLabel;
    lblDineroRetrabajo: TLabel;
    Label6: TLabel;
    Label10: TLabel;
    lblDineroTotal: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure DetalleClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure chkDllsClick(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRetrabajoDinero: TfrmRetrabajoDinero;
  giLiberadas,giRetrabajo:Integer;
  gdLiberadas,gdRetrabajo,gdLiberadasDlls,gdRetrabajoDlls: Double;

implementation

uses ReporteDineroRetrabajoGraficaQr, Main;

{$R *.dfm}

procedure TfrmRetrabajoDinero.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action:= caFree;
end;

procedure TfrmRetrabajoDinero.FormCreate(Sender: TObject);
begin
    deFrom.Date := Now;
    deTo.Date := Now;

    BindComboEmpleados(gsConnString,cmbEmpleados);
    BindComboTareasDetectado(gsConnString,cmbTareas,cmbDetectado);

    cmbEmpleados.Text := 'Todos';
    cmbTareas.Text := 'Todos';
    cmbDetectado.Text := 'Todos';

    BindGridClientes(gsConnString, gvClientes);
    CheckBox1.Checked := False;
    btnOKClick(nil);
        
    Button1Click(nil);
end;

procedure TfrmRetrabajoDinero.Button1Click(Sender: TObject);
var  Qry : TADOQuery;
Conn : TADOConnection;
SQLStr : String;
iRetrabajo : Integer;
begin
    //Create Connection
    giLiberadas := 0;
    giRetrabajo := 0;
    gdLiberadas := 0.00;
    gdRetrabajo := 0.00;
    gdLiberadasDlls := 0.00;
    gdRetrabajoDlls := 0.00;
    if txtTipo.Text = '' then txtTipo.Text := '1.00';

    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT O.*,CASE WHEN O.Dolares = 0 THEN ''No'' ELSE ''Si'' END AS DllText ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_Id = T.[Id] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ''Calidad'' AND ITS_Status = 2 ' +
              'AND ITS_DTStop >= ' + QuotedStr(deFrom.Text) +
              ' AND ITS_DTStop <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99');

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND Substring(I.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := RightStr(VarToStr(Qry['ITE_Nombre']),10);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Ordenada']);
        giLiberadas := giLiberadas + StrToInt(GridView1.Cells[1,GridView1.RowCount -1]);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Unitario']);

        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['DllText']);

        if UT(GridView1.Cells[5,GridView1.RowCount -1]) = UT('No') then begin

            // si la orden esta en pesos la cantidad en pesos es igual a multiplicar la
            // cantidad por el precio unitario, la cantidad en dolares es igual a la cantidad por
            // el unitario entre el tipo de cambio....

            GridView1.Cells[3,GridView1.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView1.Cells[1,GridView1.RowCount -1]) *
                              StrToFloat(GridView1.Cells[2,GridView1.RowCount -1]) );

            GridView1.Cells[4,GridView1.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView1.Cells[1,GridView1.RowCount -1]) *
                              ( StrToFloat(GridView1.Cells[2,GridView1.RowCount -1]) /
                                StrToFloat(txtTipo.Text) ) );
        end
        else begin
            // si la orden esta en dolares la cantidad en dolares es igual a multiplicar la
            // cantidad por el precio unitario, la cantidad en pesos es igual a la cantidad por
            // el unitario por el tipo de cambio....

            GridView1.Cells[3,GridView1.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView1.Cells[1,GridView1.RowCount -1]) *
                              StrToFloat(GridView1.Cells[2,GridView1.RowCount -1]) *
                              StrToFloat(txtTipo.Text)  );

            GridView1.Cells[4,GridView1.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView1.Cells[1,GridView1.RowCount -1]) *
                              StrToFloat(GridView1.Cells[2,GridView1.RowCount -1]) );
        end;

        gdLiberadas := gdLiberadas + StrToFloat(GridView1.Cells[3,GridView1.RowCount -1]);
        gdLiberadasDlls := gdLiberadasDlls + StrToFloat(GridView1.Cells[4,GridView1.RowCount -1]);

        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Producto']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(Qry['Recibido']);
        GridView1.Cells[11,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cells[12,GridView1.RowCount -1] := VarToStr(Qry['Interna']);
        Qry.Next;
    End;


    //***************************************************************************************/
    SQLStr := 'SELECT RIGHT(S.ITE_Nombre,LEN(S.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, ' +
              'S.RET_Start,S.RET_Stop,S.RET_Area AS Area,E.Nombre As EmpleadoRes,RET_Motivo, ' +
              'O.Unitario,CASE WHEN O.Dolares = 0 THEN ''No'' ELSE ''Si'' END AS DllText ' +
              'FROM tblRetrabajo S ' +
              'INNER JOIN tblOrdenes O ON S.ITE_Nombre  = O.ITE_Nombre ' +
              'LEFT OUTER JOIN tblEmpleados E ON E.[Id] = S.RET_Empleado ' +
              'WHERE RET_Start >= ' + QuotedStr(deFrom.Text) +
              ' AND RET_Start <= ' + QuotedStr(deTo.Text + ' 23:59:59.99' );

    if cmbDetectado.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND RET_Detectado = ' + QuotedStr(cmbDetectado.Text) + ' ';
    end;

    if cmbTareas.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND RET_Area = ' + QuotedStr(cmbTareas.Text) + ' ';
    end;

    if cmbEmpleados.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND RET_Empleado = ' + QuotedStr(LeftStr(cmbEmpleados.Text,3)) + ' ';
    end;

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND Substring(S.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView2.ClearRows;
    While not Qry.Eof do
    Begin
        GridView2.AddRow(1);
        GridView2.Cells[0,GridView2.RowCount -1] := VarToStr(Qry['Orden']);
        GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry['Cantidad']);
        GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['Unitario']);

        GridView2.Cells[5,GridView2.RowCount -1] := VarToStr(Qry['DllText']);
         if UT(GridView2.Cells[5,GridView2.RowCount -1]) = UT('No') then begin

            // si la orden esta en pesos la cantidad en pesos es igual a multiplicar la
            // cantidad por el precio unitario, la cantidad en dolares es igual a la cantidad por
            // el unitario entre el tipo de cambio....

            GridView2.Cells[3,GridView2.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView2.Cells[1,GridView2.RowCount -1]) *
                              StrToFloat(GridView2.Cells[2,GridView2.RowCount -1]) );

            GridView2.Cells[4,GridView2.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView2.Cells[1,GridView2.RowCount -1]) *
                              ( StrToFloat(GridView2.Cells[2,GridView2.RowCount -1]) /
                                StrToFloat(txtTipo.Text) ) );
        end
        else begin
            // si la orden esta en dolares la cantidad en dolares es igual a multiplicar la
            // cantidad por el precio unitario, la cantidad en pesos es igual a la cantidad por
            // el unitario por el tipo de cambio....

            GridView2.Cells[3,GridView2.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView2.Cells[1,GridView2.RowCount -1]) *
                              StrToFloat(GridView2.Cells[2,GridView2.RowCount -1]) *
                              StrToFloat(txtTipo.Text)  );

            GridView2.Cells[4,GridView2.RowCount -1] := FormatFloat( '###0.00',
                              StrToFloat(GridView2.Cells[1,GridView2.RowCount -1]) *
                              StrToFloat(GridView2.Cells[2,GridView2.RowCount -1]) );
        end;


        giRetrabajo := giRetrabajo + StrToInt(GridView2.Cells[1,GridView2.RowCount -1]);
        gdRetrabajo := gdRetrabajo + StrToFloat(GridView2.Cells[3,GridView2.RowCount -1]);
        gdRetrabajoDlls := gdRetrabajoDlls + StrToFloat(GridView2.Cells[4,GridView2.RowCount -1]);

        GridView2.Cells[6,GridView2.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView2.Cells[7,GridView2.RowCount -1] := VarToStr(Qry['Numero']);
        GridView2.Cells[8,GridView2.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView2.Cells[9,GridView2.RowCount -1] := VarToStr(Qry['Fecha']);
        GridView2.Cells[10,GridView2.RowCount -1] := VarToStr(Qry['RET_Start']);
        GridView2.Cells[11,GridView2.RowCount -1] := VarToStr(Qry['RET_Stop']);
        GridView2.Cells[12,GridView2.RowCount -1] := VarToStr(Qry['Area']);
        GridView2.Cells[13,GridView2.RowCount -1] := VarToStr(Qry['EmpleadoRes']);
        GridView2.Cells[14,GridView2.RowCount -1] := VarToStr(Qry['RET_Motivo']);
        Qry.Next;
    End;

    lblLiberadas.Caption := VarToStr(GridView1.RowCount);
    lblLiberadas2.Caption := lblLiberadas.Caption;
    lblRetrabajo.Caption := VarToStr(GridView2.RowCount);
    lblRetrabajo2.Caption := lblRetrabajo.Caption;

    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblRetrabajo.Caption));
    lblTotal2.Caption := lblTotal.Caption;


    if chkDlls.Checked then
    begin
      lblDineroLib.Caption := FormatFloat('###0.00',gdLiberadasDlls);
      lblDineroRetrabajo.Caption := FormatFloat('####0.00',gdRetrabajoDlls);
    end
    else
    begin
      lblDineroLib.Caption := FormatFloat('###0.00',gdLiberadas);
      lblDineroRetrabajo.Caption := FormatFloat('####0.00',gdRetrabajo);
    end;

    lblDineroTotal.Caption := FormatFloat('#####0.00',StrToFloat(lblDineroLib.Caption) + StrToFloat(lblDineroRetrabajo.Caption));
    if lblDineroLib.Caption = '0.00' then
        lblPorcentaje.Caption := '0.00'
    else
        lblPorcentaje.Caption := FormatFloat('#####0.00', (StrToFloat(lblDineroRetrabajo.Caption) * 100) / StrToFloat(lblDineroLib.Caption) ) + '%';

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToFloat(lblDineroLib.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToFloat(lblDineroRetrabajo.Caption),'Retrabajadas',clRed);
    Application.ProcessMessages;
end;

procedure TfrmRetrabajoDinero.DetalleClick(Sender: TObject);
begin
if GroupBox2.Visible = True then
begin
        GroupBox2.Visible := False;
        GroupBox3.Visible := True;
        Detalle.Caption := 'Grafica';
end
else
begin
        GroupBox3.Visible := False;
        GroupBox2.Visible := True;
        Detalle.Caption := 'Detalle';
end;
end;

procedure TfrmRetrabajoDinero.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmRetrabajoDinero.btnOKClick(Sender: TObject);
var i: integer;
sClientes : String;
begin
  GroupBox5.Visible := False;
  if CheckBox1.Checked = True then begin
          txtCliente.Text := 'Todos';
  end
  else begin
        sClientes := '';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                if gvClientes.Cell[1,i].AsBoolean = True then
                begin
                        sClientes := sClientes + gvClientes.Cells[0,i] + ',';
                end;
        end;
        txtCliente.Text := 'Todos';
        if sClientes <> '' then
        begin
                txtCliente.Text :=  LeftStr(sClientes,Length(sClientes) - 1);
        end;
  end;
end;

procedure TfrmRetrabajoDinero.btnTodosClick(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos.Caption) = UT('Seleccionar Todos') then begin
        btnTodos.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos.Caption := 'Seleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmRetrabajoDinero.Exportar1Click(Sender: TObject);
var sFileName: String;
begin
if GridView1.RowCount = 0 then
begin
        ShowMessage('No hay informacion que exportar.');
        Exit;
end;

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(GridView1,sFileName);

  end;


end;

procedure TfrmRetrabajoDinero.MenuItem1Click(Sender: TObject);
var sFileName: String;
begin
if GridView2.RowCount = 0 then
begin
        ShowMessage('No hay informacion que exportar.');
        Exit;
end;

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(GridView2,sFileName);

  end;
end;

procedure TfrmRetrabajoDinero.Button2Click(Sender: TObject);
begin
  if GroupBox5.Visible = True then
  begin
          GroupBox5.Visible := False;
  end
  else begin
      GroupBox5.Visible := True;
      if txtCliente.Text = 'Todos' then
      begin
              CheckBox1.Checked := True;
              gvClientes.Enabled := False;
              btnTodos.Enabled := False;
      end
      else
      begin
              CheckBox1.Checked := False;
              gvClientes.Enabled := True;
              btnTodos.Enabled := True;
      end;

      GroupBox5.Top := txtCliente.Top + txtCliente.Height + 5;
      GroupBox5.Left := txtCliente.Left + 10;
  end;
end;

procedure TfrmRetrabajoDinero.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Retrabajo';

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

       showmessage('El archivo se creo exitosamente.');
end;


procedure TfrmRetrabajoDinero.chkDllsClick(Sender: TObject);
begin
    if chkDlls.Checked then
    begin
      lblDineroLib.Caption := FormatFloat('###0.00',gdLiberadasDlls);
      lblDineroRetrabajo.Caption := FormatFloat('####0.00',gdRetrabajoDlls);
    end
    else begin
      lblDineroLib.Caption := FormatFloat('###0.00',gdLiberadas);
      lblDineroRetrabajo.Caption := FormatFloat('####0.00',gdRetrabajo);
    end;

    lblDineroTotal.Caption := FormatFloat('#####0.00',StrToFloat(lblDineroLib.Caption) + StrToFloat(lblDineroRetrabajo.Caption));
    if lblDineroTotal.Caption = '0.00' then
        lblPorcentaje.Caption := '0.00'
    else
        lblPorcentaje.Caption := FormatFloat('#####0.00', (StrToFloat(lblDineroRetrabajo.Caption) * 100) / StrToFloat(lblDineroTotal.Caption) ) + '%';

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToFloat(lblDineroLib.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToFloat(lblDineroRetrabajo.Caption),'Retrabajadas',clRed);
    Application.ProcessMessages;
end;

procedure TfrmRetrabajoDinero.ImprimirClick(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrDineroRetrabajoGrafica,qrDineroRetrabajoGrafica);
    qrDineroRetrabajoGrafica.ReportTitle.Caption := 'Porcentaje de Retrabajo en Dinero desde ' + deFrom.Text + ' hasta ' + deTo.Text;

    qrDineroRetrabajoGrafica.lblLiberadas.Caption := lblCLiberadas.Caption + lblLiberadas.Caption ;
    qrDineroRetrabajoGrafica.lblScrap.Caption := lblCRetrabajo.Caption + lblRetrabajo.Caption;
    qrDineroRetrabajoGrafica.lblTotal.Caption := lblCTotal.Caption + lblTotal.Caption;

    qrDineroRetrabajoGrafica.lblDineroLiberadas.Caption := 'Cantidad Liberada : ' + lblDineroLib.Caption ;
    qrDineroRetrabajoGrafica.lblDineroScrap.Caption := 'Cantidad Retrabajada : ' + lblDineroRetrabajo.Caption;
    qrDineroRetrabajoGrafica.lblDineroTotal.Caption := 'Cantidad Total : ' + lblDineroTotal.Caption;

    qrDineroRetrabajoGrafica.lblPorcentaje.Caption := 'Porcentaje : ' + lblPorcentaje.Caption;

    qrDineroRetrabajoGrafica.QRChart1.Chart.Series[0].Clear;
    qrDineroRetrabajoGrafica.QRChart1.Chart.Series[0].Add(StrToFloat(lblDineroLib.Caption),'Liberadas',clBlue);
    qrDineroRetrabajoGrafica.QRChart1.Chart.Series[0].Add(StrToFloat(lblDineroRetrabajo.Caption),'Retrabajadas',clRed);

    //qrRetrabajoGrafica.Print;
    qrDineroRetrabajoGrafica.Preview;
    qrDineroRetrabajoGrafica.Free;
end;

end.
