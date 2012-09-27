unit PorcentajeRetrabajo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors,ADODB,DB,IniFiles,ComObj,All_Functions,StrUtils,chris_Functions,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils;

type
  TfrmRetrabajo = class(TForm)
    GroupBox3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    lblLiberadas2: TLabel;
    lblCLiberadas2: TLabel;
    lblTotal2: TLabel;
    lblCTotal2: TLabel;
    Label15: TLabel;
    lblPorcentaje2: TLabel;
    GroupBox4: TGroupBox;
    GridView1: TGridView;
    GridView2: TGridView;
    GroupBox2: TGroupBox;
    lblCLiberadas: TLabel;
    lblCTotal: TLabel;
    lblLiberadas: TLabel;
    lblTotal: TLabel;
    lblCPorcentaje: TLabel;
    lblPorcentaje: TLabel;
    Chart1: TChart;
    Series1: TPieSeries;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label9: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Detalle: TButton;
    Imprimir: TButton;
    cmbDetectado: TComboBox;
    cmbTareas: TComboBox;
    cmbEmpleados: TComboBox;
    chkPiezas: TCheckBox;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    lblCRetrabajo: TLabel;
    lblRetrabajo: TLabel;
    lblCRetrabajo2: TLabel;
    lblRetrabajo2: TLabel;
    Label3: TLabel;
    txtCliente: TEdit;
    Button2: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure DetalleClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure MenuItem1Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure BindEmpleados();
    procedure BindTareas();
    procedure chkPiezasClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindClientes();
    procedure Button2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRetrabajo: TfrmRetrabajo;
  giLiberadas,giRetrabajo:Integer;

implementation

uses ReporteRetrabajoGraficaQr, Main;

{$R *.dfm}


procedure TfrmRetrabajo.FormCreate(Sender: TObject);
begin
    deFrom.Date := Now;
    deTo.Date := Now;

    BindEmpleados;
    BindTareas;
    cmbTareas.Text := 'Todos';
    cmbDetectado.Text := 'Todos';
    cmbEmpleados.Text := 'Todos';
    BindClientes();
    CheckBox1.Checked := False;
    btnOKClick(nil);

    Button1Click(nil);
end;

procedure TfrmRetrabajo.Button1Click(Sender: TObject);
var  Qry : TADOQuery;
Conn : TADOConnection;
SQLStr : String;
begin
    //Create Connection
    giLiberadas := 0;
    giRetrabajo := 0;
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT O.* ' +
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
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Producto']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Recibido']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Interna']);
        Qry.Next;
    End;

    if chkPiezas.Checked then
       begin
          lblCLiberadas.Caption := '    Piezas Liberadas : ';
          lblCLiberadas2.Caption := '    Piezas Liberadas : ';
          lblLiberadas.Caption := VarToStr(giLiberadas);
          lblLiberadas2.Caption := lblLiberadas.Caption;
       end
    else
       begin
          lblCLiberadas.Caption := 'Ordenes Liberadas : ';
          lblCLiberadas2.Caption := 'Ordenes Liberadas : ';
          lblLiberadas.Caption := VarToStr(GridView1.RowCount);
          lblLiberadas2.Caption := lblLiberadas.Caption;
       end;

    SQLStr := 'SELECT RIGHT(S.ITE_Nombre,LEN(S.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, ' +
              'S.RET_Start,S.RET_Stop,S.RET_Area AS Area,E.Nombre As EmpleadoRes,RET_Motivo, ' +
              'S.RET_Detectado, D.Nombre As EmpleadoDetectado ' +
              'FROM tblRetrabajo S ' +
              'INNER JOIN tblOrdenes O ON S.ITE_Nombre  = O.ITE_Nombre ' +
              'LEFT OUTER JOIN tblEmpleados E ON E.[Id] = S.RET_Empleado ' +
              'LEFT OUTER JOIN tblEmpleados D ON D.[Id] = S.RET_EmpleadoDetectado ' +
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
        giRetrabajo := giRetrabajo + StrToInt(GridView2.Cells[1,GridView2.RowCount -1]);
        GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView2.Cells[3,GridView2.RowCount -1] := VarToStr(Qry['Numero']);
        GridView2.Cells[4,GridView2.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView2.Cells[5,GridView2.RowCount -1] := VarToStr(Qry['Fecha']);
        GridView2.Cells[6,GridView2.RowCount -1] := VarToStr(Qry['RET_Start']);
        GridView2.Cells[7,GridView2.RowCount -1] := VarToStr(Qry['RET_Stop']);
        GridView2.Cells[8,GridView2.RowCount -1] := VarToStr(Qry['Area']);
        GridView2.Cells[9,GridView2.RowCount -1] := VarToStr(Qry['EmpleadoRes']);
        GridView2.Cells[10,GridView2.RowCount -1] := VarToStr(Qry['RET_Detectado']);
        GridView2.Cells[11,GridView2.RowCount -1] := VarToStr(Qry['EmpleadoDetectado']);
        GridView2.Cells[12,GridView2.RowCount -1] := VarToStr(Qry['RET_Motivo']);
        Qry.Next;
    End;

    if chkPiezas.Checked then
       begin
          lblCRetrabajo.Caption := '    Piezas Retrabajadas : ';
          lblCTotal.Caption := '    Total de Piezas :';
          lblCRetrabajo2.Caption := '    Piezas Retrabajadas : ';
          lblCTotal2.Caption := '    Total de Piezas :';
          lblRetrabajo.Caption := VarToStr(giRetrabajo);
          lblRetrabajo2.Caption := lblRetrabajo.Caption;
       end
    else
       begin
          lblCRetrabajo.Caption := 'Ordenes Retrabajadas : ';
          lblCTotal.Caption := 'Total de Ordenes :';
          lblCRetrabajo2.Caption := 'Ordenes Retrabajadas : ';
          lblCTotal2.Caption := 'Total de Ordenes :';
          lblRetrabajo.Caption := VarToStr(GridView2.RowCount);
          lblRetrabajo2.Caption := lblRetrabajo.Caption;
       end;


    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblRetrabajo.Caption));
    if lblTotal.Caption = '0' then
        lblPorcentaje.Caption := '0'
    else
        lblPorcentaje.Caption := FormatFloat('######.00', (StrToInt(lblRetrabajo.Caption) * 100) / StrToInt(lblLiberadas.Caption) ) + '%';

    lblTotal2.Caption := lblTotal.Caption;
    lblPorcentaje2.Caption := lblPorcentaje.Caption;

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToInt(lblRetrabajo.Caption),'Retrabajadas',clRed);
    Application.ProcessMessages;
end;

procedure TfrmRetrabajo.DetalleClick(Sender: TObject);
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

procedure TfrmRetrabajo.Exportar1Click(Sender: TObject);
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

procedure TfrmRetrabajo.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmRetrabajo.MenuItem1Click(Sender: TObject);
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

procedure TfrmRetrabajo.ImprimirClick(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrRetrabajoGrafica,qrRetrabajoGrafica);
    qrRetrabajoGrafica.ReportTitle.Caption := 'Porcentaje de Retrabajo desde ' + deFrom.Text + ' hasta ' + deTo.Text;

    qrRetrabajoGrafica.lblLiberadas.Caption := Trim(lblCLiberadas.Caption) + ' ' + lblLiberadas.Caption ;
    qrRetrabajoGrafica.lblRetrabajo.Caption := Trim(lblCRetrabajo.Caption) + ' ' + lblRetrabajo.Caption;
    qrRetrabajoGrafica.lblTotal.Caption := Trim(lblCTotal.Caption) + ' ' + lblTotal.Caption;
    qrRetrabajoGrafica.lblPorcentaje.Caption := 'Porcentaje : ' + lblPorcentaje.Caption;

    qrRetrabajoGrafica.QRChart1.Chart.Series[0].Clear;
    qrRetrabajoGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    qrRetrabajoGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblRetrabajo.Caption),'Retrabajadas',clRed);

    //qrRetrabajoGrafica.Print;
    qrRetrabajoGrafica.Preview;
    qrRetrabajoGrafica.Free;
end;


procedure TfrmRetrabajo.BindEmpleados();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT ID,Nombre FROM tblEmpleados Order By Nombre';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    cmbEmpleados.Items.Clear;
    cmbEmpleados.Items.Add('Todos');
    cmbEmpleados.Items.Add('000 - Desconocido');
    While not Qry.Eof do
    Begin
        cmbEmpleados.Items.Add(FormatFloat('000',Qry['ID']) + ' - ' + Qry['Nombre']);
        Qry.Next;
    End;

    cmbEmpleados.Text := '';
    Qry.Close;
    Conn.Close;
end;

procedure TfrmRetrabajo.BindTareas();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    cmbTareas.Items.Clear;
    cmbDetectado.Items.Clear;
    cmbTareas.Items.Add('Todos');
    cmbDetectado.Items.Add('Todos');
    While not Qry.Eof do
    Begin
        cmbTareas.Items.Add(Qry['Nombre']);
        cmbDetectado.Items.Add(Qry['Nombre']);
        Qry.Next;
    End;

    cmbTareas.Text := '';
    cmbDetectado.Text := '';
    Qry.Close;
    Conn.Close;
end;


procedure TfrmRetrabajo.chkPiezasClick(Sender: TObject);
begin
    if chkPiezas.Checked then
       begin
          lblCLiberadas.Caption := '    Piezas Liberadas : ';
          lblCLiberadas2.Caption := '    Piezas Liberadas : ';
          lblLiberadas.Caption := VarToStr(giLiberadas);
          lblLiberadas2.Caption := lblLiberadas.Caption;
          lblCRetrabajo.Caption := '    Piezas Retrabajadas : ';
          lblCTotal.Caption := '    Total de Piezas :';
          lblCRetrabajo2.Caption := 'Piezas Retrabajadas : ';
          lblCTotal2.Caption := 'Total de Piezas :';
          lblRetrabajo.Caption := VarToStr(giRetrabajo);
          lblRetrabajo2.Caption := lblRetrabajo.Caption;
       end
    else
       begin
          lblCLiberadas.Caption := 'Ordenes Liberadas : ';
          lblCLiberadas2.Caption := 'Ordenes Liberadas : ';
          lblLiberadas.Caption := VarToStr(GridView1.RowCount);
          lblLiberadas2.Caption := lblLiberadas.Caption;
          lblCRetrabajo.Caption := 'Ordenes Retrabajadas : ';
          lblCTotal.Caption := 'Total de Ordenes :';
          lblCRetrabajo2.Caption := 'Ordenes Retrabajadas : ';
          lblCTotal2.Caption := 'Total de Ordenes :';
          lblRetrabajo.Caption := VarToStr(GridView2.RowCount);
          lblRetrabajo2.Caption := lblRetrabajo.Caption;
       end;

    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblRetrabajo.Caption));
    if lblTotal.Caption = '0' then
        lblPorcentaje.Caption := '0'
    else
        lblPorcentaje.Caption := FormatFloat('######.00', (StrToInt(lblRetrabajo.Caption) * 100) / StrToInt(lblTotal.Caption) ) + '%';

    lblTotal2.Caption := lblTotal.Caption;
    lblPorcentaje2.Caption := lblPorcentaje.Caption;

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToInt(lblRetrabajo.Caption),'Retrabajadas',clRed);
    Application.ProcessMessages;


end;
procedure TfrmRetrabajo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmRetrabajo.BindClientes();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT Distinct Clave FROM tblClientes Order By Clave';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvClientes.ClearRows;
    While not Qry.Eof do
    Begin
        gvClientes.AddRow(1);
        gvClientes.Cells[0,gvClientes.RowCount -1] := VarToStr(Qry['Clave']);
        if (slClientes.IndexOf(VarToStr(Qry['Clave'])) = -1) then begin
                gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
        end;
        
        Qry.Next;
    End;
end;


procedure TfrmRetrabajo.Button2Click(Sender: TObject);
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

procedure TfrmRetrabajo.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmRetrabajo.btnOKClick(Sender: TObject);
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

procedure TfrmRetrabajo.btnTodosClick(Sender: TObject);
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

end.
