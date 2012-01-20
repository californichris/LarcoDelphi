unit ReporteCumplimiento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors,ADODB,DB,IniFiles,ComObj,All_Functions,StrUtils,chris_Functions,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, ScrollView,LTCUtils,
  CustomGridViewControl, CustomGridView, GridView, Menus,Clipbrd;

type
  TfrmCumplimiento = class(TForm)
    GroupBox3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    lblLiberadas2: TLabel;
    lblCLiberadas2: TLabel;
    lblCScrap2: TLabel;
    lblScrap2: TLabel;
    lblTotal2: TLabel;
    lblCTotal2: TLabel;
    Label15: TLabel;
    lblPorcentaje2: TLabel;
    GroupBox4: TGroupBox;
    GridView1: TGridView;
    GridView2: TGridView;
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
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label9: TLabel;
    Label3: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Detalle: TButton;
    Imprimir: TButton;
    cmbTipo: TComboBox;
    txtCliente: TEdit;
    Button2: TButton;
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
    CopiarOrden1: TMenuItem;
    CopiarOrden2: TMenuItem;
    procedure Button2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure DetalleClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Exportar1Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure BindClientes();
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure Button1Click(Sender: TObject);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure CopiarOrden2Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCumplimiento: TfrmCumplimiento;
  giLiberadas,giScrap:Integer;
  s1,s2:String;

implementation

uses ReporteCumplimientoQr, Main;

{$R *.dfm}

procedure TfrmCumplimiento.Button2Click(Sender: TObject);
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

procedure TfrmCumplimiento.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmCumplimiento.btnOKClick(Sender: TObject);
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

procedure TfrmCumplimiento.btnTodosClick(Sender: TObject);
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

procedure TfrmCumplimiento.DetalleClick(Sender: TObject);
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

procedure TfrmCumplimiento.FormCreate(Sender: TObject);
begin
    deFrom.Date := Now;
    deTo.Date := Now;

    BindClientes();
    CheckBox1.Checked := False;
    btnOKClick(nil);

    Button1Click(nil);
end;

procedure TfrmCumplimiento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCumplimiento.Exportar1Click(Sender: TObject);
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

procedure TfrmCumplimiento.MenuItem1Click(Sender: TObject);
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

procedure TfrmCumplimiento.ExportGrid(Grid: TGridView;sFileName: String);
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

procedure TfrmCumplimiento.BindClientes();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
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


procedure TfrmCumplimiento.Button1Click(Sender: TObject);
var  Qry : TADOQuery;
Conn : TADOConnection;
SQLStr,sTarea,sFecha : String;
begin
    //Create Connection
    giLiberadas := 0;
    giScrap := 0;
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    sTarea := 'Calidad';
    sFecha := 'Interna';
    lblCLiberadas.Caption := 'Ordenes en Calidad o Despues :';
    lblCScrap.Caption := 'Ordenes antes de Calidad :';
    lblCLiberadas2.Caption := 'Ordenes en Calidad o Despues :';
    lblCScrap2.Caption := 'Ordenes antes de Calidad :';

    s1 := 'en Calidad o Despues';
    s2 := 'antes de Calidad';
    if UT(cmbTipo.Text) = UT('Cliente') Then begin
            sTarea := 'VentasFinal';
            sFecha := 'Entrega';
            lblCLiberadas.Caption := 'Ordenes A Tiempo :';
            lblCScrap.Caption := 'Ordenes No A Tiempo :';
            lblCLiberadas2.Caption := 'Ordenes A Tiempo :';
            lblCScrap2.Caption := 'Ordenes No A Tiempo :';
            s1 := 'A Tiempo';
            s2 := 'No A Tiempo';
    end;

    SQLStr := 'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_Id = T.[Id] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ' + QuotedStr(sTarea) + ' AND ITS_Status IS NOT NULL ' +
              'AND ITS_Status <> 9 ' +
              'AND ITS_DTStop <= CONVERT(VARCHAR(10),' +  sFecha + ',101) + '' 23:59:59.99'' ' +
              'AND I.ITE_NOMBRE NOT IN (SELECT ITE_NOMBRE FROM tblScrap) ' +
              'AND (' + sFecha + ' >= ' + QuotedStr(deFrom.Text) + ' ' +
              'AND ' + sFecha + ' <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99') + ') ';

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND Substring(I.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT O.*,I.ITS_DTStop ' + SQLStr + ' ORDER BY O.ITE_Nombre';
    Qry.Open;

{
    'en calidad o despues'
    'antes de calidad'

    'Terminados'
    'No Terminados'
}

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
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Interna']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['ITS_DTStop']);
        Qry.Next;
    End;

    // ************************************************************************************** //
    // ************************************************************************************** //

{    SQLStr := 'SELECT * ' +
              'FROM tblOrdenes O ' +
              'WHERE (' + sFecha + ' >= ' + QuotedStr(deFrom.Text) + ' ' +
              'AND ' + sFecha + ' <= ' + QuotedStr(deTo.Text + ' 23:59:59.99' ) + ') ' +
              'AND O.ITE_Nombre NOT IN (' + 'SELECT O.ITE_Nombre ' + SQLStr + ') ' +
              'AND O.ITE_Nombre NOT IN (SELECT ITE_Nombre FROM tblScrap) ';

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND Substring(O.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;
}

    SQLStr := 'CumplimientoNoATiempo ' + QuotedStr(deFrom.Text) + ',' +
               QuotedStr(deTo.Text  + ' 23:59:59.99') + ',';


    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + QuotedStr(txtCliente.Text) + ',';
    end
    else begin
        SQLStr := SQLStr + QuotedStr('') + ',';
    end;

    SQLStr := SQLStr + QuotedStr(cmbTipo.Text);


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView2.ClearRows;
    While not Qry.Eof do
    Begin
        GridView2.AddRow(1);
        GridView2.Cells[0,GridView2.RowCount -1] := RightStr(VarToStr(Qry['ITE_Nombre']),10);;
        GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry['Ordenada']);
        giLiberadas := giLiberadas + StrToInt(GridView2.Cells[1,GridView2.RowCount -1]);
        GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['Producto']);
        GridView2.Cells[3,GridView2.RowCount -1] := VarToStr(Qry['Numero']);
        GridView2.Cells[4,GridView2.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView2.Cells[5,GridView2.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        GridView2.Cells[6,GridView2.RowCount -1] := VarToStr(Qry['Recibido']);
        GridView2.Cells[7,GridView2.RowCount -1] := VarToStr(Qry['Interna']);
        GridView2.Cells[8,GridView2.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView2.Cells[9,GridView2.RowCount -1] := VarToStr(Qry['Calidad']);
        GridView2.Cells[10,GridView2.RowCount -1] := VarToStr(Qry['VentasFinal']);

        Qry.Next;
    End;

    lblLiberadas.Caption := VarToStr(GridView1.RowCount);
    lblLiberadas2.Caption := lblLiberadas.Caption;

    lblScrap.Caption := VarToStr(GridView2.RowCount);
    lblScrap2.Caption := lblScrap.Caption;


    lblTotal.Caption := IntToStr(StrToInt(lblLiberadas.Caption) + StrToInt(lblScrap.Caption));
    if lblTotal.Caption = '0' then
        lblPorcentaje.Caption := '0'
    else
        lblPorcentaje.Caption := FormatFloat('######.00', (StrToInt(lblScrap.Caption) * 100) / StrToInt(lblTotal.Caption) ) + '%';

    lblTotal2.Caption := lblTotal.Caption;
    lblPorcentaje2.Caption := lblPorcentaje.Caption;

    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToInt(lblLiberadas.Caption),s1,clBlue);
    Chart1.Series[0].Add(StrToInt(lblScrap.Caption),s2,clRed);
    Application.ProcessMessages;
end;

procedure TfrmCumplimiento.CopiarOrden1Click(Sender: TObject);
begin
      Clipboard.AsText := GridView1.Cells[0,GridView1.SelectedRow];
end;

procedure TfrmCumplimiento.CopiarOrden2Click(Sender: TObject);
begin
      Clipboard.AsText := GridView2.Cells[0,GridView2.SelectedRow];
end;

procedure TfrmCumplimiento.ImprimirClick(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrCumpliGrafica,qrCumpliGrafica);
    qrCumpliGrafica.ReportTitle.Caption := 'Porcentaje de Cumplimiento desde ' + deFrom.Text + ' hasta ' + deTo.Text;

    qrCumpliGrafica.lblLiberadas.Caption := lblCLiberadas.Caption + lblLiberadas.Caption ;
    qrCumpliGrafica.lblScrap.Caption := lblCScrap.Caption + lblScrap.Caption;
    qrCumpliGrafica.lblTotal.Caption := lblCTotal.Caption + lblTotal.Caption;
    qrCumpliGrafica.lblPorcentaje.Caption := 'Porcentaje : ' + lblPorcentaje.Caption;

    qrCumpliGrafica.QRChart1.Chart.Series[0].Clear;
    qrCumpliGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblLiberadas.Caption),s1,clBlue);
    qrCumpliGrafica.QRChart1.Chart.Series[0].Add(StrToInt(lblScrap.Caption),s2,clRed);

    //qrScrapGrafica.Print;
    qrCumpliGrafica.Preview;
    qrCumpliGrafica.Free;
end;

end.
