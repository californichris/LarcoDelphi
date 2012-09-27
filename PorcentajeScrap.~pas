unit PorcentajeScrap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors,ADODB,DB,IniFiles,ComObj,All_Functions,StrUtils,chris_Functions,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, ScrollView,LTCUtils,
  CustomGridViewControl, CustomGridView, GridView, Menus;

type
  TfrmScrapPorcen = class(TForm)
    GroupBox1: TGroupBox;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    GroupBox2: TGroupBox;
    lblCLiberadas: TLabel;
    lblCScrap: TLabel;
    lblCTotal: TLabel;
    lblLiberadas: TLabel;
    lblScrap: TLabel;
    lblTotal: TLabel;
    lblCPorcentaje: TLabel;
    lblPorcentaje: TLabel;
    Chart1: TChart;
    Series1: TPieSeries;
    GroupBox3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    GroupBox4: TGroupBox;
    lblLiberadas2: TLabel;
    lblCLiberadas2: TLabel;
    lblCScrap2: TLabel;
    lblScrap2: TLabel;
    lblTotal2: TLabel;
    lblCTotal2: TLabel;
    Label15: TLabel;
    lblPorcentaje2: TLabel;
    GridView1: TGridView;
    GridView2: TGridView;
    Detalle: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    Imprimir: TButton;
    Label9: TLabel;
    cmbDetectado: TComboBox;
    Label12: TLabel;
    cmbTareas: TComboBox;
    Label13: TLabel;
    cmbEmpleados: TComboBox;
    chkPiezas: TCheckBox;
    Label3: TLabel;
    txtCliente: TEdit;
    Button2: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
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
    procedure BindClientes();
    procedure CheckBox1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScrapPorcen: TfrmScrapPorcen;
  giLiberadas,giScrap:Integer;
implementation

uses ReporteScrapGraficaQr, Main;

{$R *.dfm}

procedure TfrmScrapPorcen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmScrapPorcen.FormCreate(Sender: TObject);
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

procedure TfrmScrapPorcen.Button1Click(Sender: TObject);
var  Qry : TADOQuery;
Conn : TADOConnection;
SQLStr, parcial : String;
begin
    //Create Connection
    giLiberadas := 0;
    giScrap := 0;
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
        GridView1.Cells[0,GridView1.RowCount -1] := RightStr(VarToStr(Qry['ITE_Nombre']),10);;
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

    SQLStr := 'SELECT RIGHT(S.ITE_Nombre,LEN(S.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, ' +
              'RIGHT(S.SCR_NewItem,LEN(S.SCR_NewItem) - 3) AS NewOrden, ' +
              'S.SCR_Cantidad,S.SCR_Fecha,S.SCR_Tarea AS Area,E.Nombre As EmpleadoRes, '+
              'SCR_Motivo,SCR_Parcial,SCR_Repro,SCR_Detectado, SCR_EmpleadoDetectado,D.Nombre As EmpleadoDetectado ' +
              'FROM tblScrap S ' +
              'INNER JOIN tblOrdenes O ON S.ITE_Nombre  = O.ITE_Nombre ' +
              'LEFT OUTER JOIN tblEmpleados E ON E.[Id] = S.SCR_EmpleadoRes ' +
              'LEFT OUTER JOIN tblEmpleados D ON D.[Id] = S.SCR_EmpleadoDetectado ' +
              'WHERE SCR_Fecha >= ' + QuotedStr(deFrom.Text) +
              ' AND SCR_Fecha <= ' + QuotedStr(deTo.Text + ' 23:59:59.99' );

    if cmbDetectado.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND SCR_Detectado = ' + QuotedStr(cmbDetectado.Text) + ' ';
    end;

    if cmbTareas.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND SCR_Tarea = ' + QuotedStr(cmbTareas.Text) + ' ';
    end;

    if cmbEmpleados.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND SCR_EmpleadoRes = ' + QuotedStr(LeftStr(cmbEmpleados.Text,3)) + ' ';
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
        GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['SCR_Cantidad']);
        GridView2.Cells[3,GridView2.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView2.Cells[4,GridView2.RowCount -1] := VarToStr(Qry['Numero']);
        GridView2.Cells[5,GridView2.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView2.Cells[6,GridView2.RowCount -1] := VarToStr(Qry['Fecha']);
        GridView2.Cells[7,GridView2.RowCount -1] := VarToStr(Qry['NewOrden']);
        GridView2.Cells[8,GridView2.RowCount -1] := VarToStr(Qry['SCR_Fecha']);
        GridView2.Cells[9,GridView2.RowCount -1] := VarToStr(Qry['Area']);
        GridView2.Cells[10,GridView2.RowCount -1] := VarToStr(Qry['EmpleadoRes']);
        GridView2.Cells[11,GridView2.RowCount -1] := VarToStr(Qry['SCR_Detectado']);
        GridView2.Cells[12,GridView2.RowCount -1] := VarToStr(Qry['EmpleadoDetectado']);
        GridView2.Cells[13,GridView2.RowCount -1] := VarToStr(Qry['SCR_Motivo']);
        parcial := VarToStr(Qry['SCR_Parcial']);
        if parcial = '-1' then begin
          parcial := 'Si';
        end
        else begin
          parcial := 'No';
        end;
        GridView2.Cells[14,GridView2.RowCount -1] := parcial;
        GridView2.Cells[15,GridView2.RowCount -1] := VarToStr(Qry['SCR_Repro']);

        giScrap := giScrap + StrToInt(GridView2.Cells[2,GridView2.RowCount -1]);
        if GridView2.Cells[12,GridView2.RowCount -1] <> '0' then begin
                giLiberadas := giLiberadas +
                ( StrToInt(GridView2.Cells[1,GridView2.RowCount -1]) -  StrToInt(GridView2.Cells[2,GridView2.RowCount -1]));
        end;


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


    if chkPiezas.Checked then
       begin
          lblCScrap.Caption := '    Piezas Scrapeadas : ';
          lblCTotal.Caption := '    Total de Piezas :';
          lblCScrap2.Caption := '    Piezas Scrapeadas : ';
          lblCTotal2.Caption := '    Total de Piezas :';
          lblScrap.Caption := VarToStr(giScrap);
          lblScrap2.Caption := lblScrap.Caption;
       end
    else
       begin
          lblCScrap.Caption := 'Ordenes Scrapeadas : ';
          lblCTotal.Caption := 'Total de Ordenes :';
          lblCScrap2.Caption := 'Ordenes Scrapeadas : ';
          lblCTotal2.Caption := 'Total de Ordenes :';
          lblScrap.Caption := VarToStr(GridView2.RowCount);
          lblScrap2.Caption := lblScrap.Caption;
       end;


    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblScrap.Caption));
    if lblTotal.Caption = '0' then
        lblPorcentaje.Caption := '0'
    else
        lblPorcentaje.Caption := FormatFloat('######.00', (StrToInt(lblScrap.Caption) * 100) / StrToInt(lblTotal.Caption) ) + '%';

    lblTotal2.Caption := lblTotal.Caption;
    lblPorcentaje2.Caption := lblPorcentaje.Caption;

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToInt(lblScrap.Caption),'Scrapeadas',clRed);
    Application.ProcessMessages;
end;

procedure TfrmScrapPorcen.DetalleClick(Sender: TObject);
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

procedure TfrmScrapPorcen.Exportar1Click(Sender: TObject);
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

procedure TfrmScrapPorcen.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Scrap';

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


procedure TfrmScrapPorcen.MenuItem1Click(Sender: TObject);
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

procedure TfrmScrapPorcen.ImprimirClick(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrScrapGrafica,qrScrapGrafica);
    qrScrapGrafica.ReportTitle.Caption := 'Porcentaje de Scrap desde ' + deFrom.Text + ' hasta ' + deTo.Text;

    qrScrapGrafica.lblLiberadas.Caption := Trim(lblCLiberadas.Caption) + ' ' + lblLiberadas.Caption ;
    qrScrapGrafica.lblScrap.Caption := Trim(lblCScrap.Caption) + ' ' + lblScrap.Caption;
    qrScrapGrafica.lblTotal.Caption := Trim(lblCTotal.Caption) + ' ' + lblTotal.Caption;
    qrScrapGrafica.lblPorcentaje.Caption := 'Porcentaje : ' + lblPorcentaje.Caption;

    qrScrapGrafica.QRChart1.Chart.Series[0].Clear;
    qrScrapGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    qrScrapGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblScrap.Caption),'Scrapeadas',clRed);

    //qrScrapGrafica.Print;
    qrScrapGrafica.Preview;
    qrScrapGrafica.Free;
end;


procedure TfrmScrapPorcen.BindEmpleados();
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

procedure TfrmScrapPorcen.BindTareas();
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


procedure TfrmScrapPorcen.chkPiezasClick(Sender: TObject);
begin
    if chkPiezas.Checked then
       begin
          lblCLiberadas.Caption := '    Piezas Liberadas : ';
          lblCLiberadas2.Caption := '    Piezas Liberadas : ';
          lblLiberadas.Caption := VarToStr(giLiberadas);
          lblLiberadas2.Caption := lblLiberadas.Caption;
          lblCScrap.Caption := '    Piezas Scrapeadas : ';
          lblCTotal.Caption := '    Total de Piezas :';
          lblCScrap2.Caption := 'Piezas Scrapeadas : ';
          lblCTotal2.Caption := 'Total de Piezas :';
          lblScrap.Caption := VarToStr(giScrap);
          lblScrap2.Caption := lblScrap.Caption;
       end
    else
       begin
          lblCLiberadas.Caption := 'Ordenes Liberadas : ';
          lblCLiberadas2.Caption := 'Ordenes Liberadas : ';
          lblLiberadas.Caption := VarToStr(GridView1.RowCount);
          lblLiberadas2.Caption := lblLiberadas.Caption;
          lblCScrap.Caption := 'Ordenes Scrapeadas : ';
          lblCTotal.Caption := 'Total de Ordenes :';
          lblCScrap2.Caption := 'Ordenes Scrapeadas : ';
          lblCTotal2.Caption := 'Total de Ordenes :';
          lblScrap.Caption := VarToStr(GridView2.RowCount);
          lblScrap2.Caption := lblScrap.Caption;
       end;

    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblScrap.Caption));
    if lblTotal.Caption = '0' then
        lblPorcentaje.Caption := '0'
    else
        lblPorcentaje.Caption := FormatFloat('######.00', (StrToInt(lblScrap.Caption) * 100) / StrToInt(lblTotal.Caption) ) + '%';

    lblTotal2.Caption := lblTotal.Caption;
    lblPorcentaje2.Caption := lblPorcentaje.Caption;

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToInt(lblLiberadas.Caption),'Liberadas',clBlue);
    Chart1.Series[0].Add(StrToInt(lblScrap.Caption),'Scrapeadas',clRed);
    Application.ProcessMessages;


end;

procedure TfrmScrapPorcen.BindClientes();
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

procedure TfrmScrapPorcen.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmScrapPorcen.Button2Click(Sender: TObject);
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

procedure TfrmScrapPorcen.btnOKClick(Sender: TObject);
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

procedure TfrmScrapPorcen.btnTodosClick(Sender: TObject);
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
