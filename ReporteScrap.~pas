unit ReporteScrap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ExtCtrls, CellEditors, ScrollView,ComCtrls,ComObj,
  CustomGridViewControl, CustomGridView, GridView,ReporteScrapQr;

type
  TfrmScrapReport = class(TForm)
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    deRecibido1: TDateEditor;
    deRecibido2: TDateEditor;
    Button1: TButton;
    cmbProductos: TComboBox;
    chkRecibido: TCheckBox;
    deInterna1: TDateEditor;
    chkInterna: TCheckBox;
    deInterna2: TDateEditor;
    deEntrega1: TDateEditor;
    chkEntrega: TCheckBox;
    deEntrega2: TDateEditor;
    cmbClientes: TComboBox;
    cmbPartes: TComboBox;
    Button2: TButton;
    btnBuscar: TButton;
    cmbOrdenes: TComboBox;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    chkScrap: TCheckBox;
    deScrap1: TDateEditor;
    deScrap2: TDateEditor;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure chkRecibidoClick(Sender: TObject);
    procedure chkInternaClick(Sender: TObject);
    procedure chkEntregaClick(Sender: TObject);
    procedure chkScrapClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    Procedure BindPartes(Query: String);
    Procedure BindOrdenes(Query: String);
    procedure BindGrid();
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure btnBuscarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    function FormIsRunning(FormName: String):Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScrapReport: TfrmScrapReport;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main;

{$R *.dfm}

procedure TfrmScrapReport.FormCreate(Sender: TObject);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Timer1.Interval := frmMain.iIntervalo;

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT CASE WHEN Min(Interna) IS NULL THEN GETDATE() ELSE Min(Interna) END As Interna ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    deInterna1.Date := Now;
    if Qry2.RecordCount > 0 then
            deInterna1.Date := StrToDateTime( VarToStr(Qry2['Interna']) ) ;

    deInterna2.Date := DateAdd(Now,5,daDays);

    deRecibido1.Date := Now;
    deRecibido2.Date := Now;

    deEntrega1.Date := Now;
    deEntrega2.Date := Now;
    deScrap1.Date := Now;
    deScrap2.Date := Now;

    cmbProductos.Text := 'Todos';
    cmbClientes.Text := 'Todos';
    cmbPartes.Text := 'Todos';
    cmbOrdenes.Text := 'Todos';

    BindGrid();
    BindProductos();
    BindClientes();

    Qry2.Close;
end;

procedure TfrmScrapReport.BindGrid();
var SQLStr,SQLWhere : String;
begin
    SQLStr := 'SELECT RIGHT(S.ITE_Nombre,LEN(S.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, ' +
              'RIGHT(S.SCR_NewItem,LEN(S.SCR_NewItem) - 3) AS NewOrden, ' +
              'S.SCR_Fecha,S.SCR_Tarea AS Area,E.Nombre As EmpleadoRes,SCR_Motivo ' +
              'FROM tblScrap S ' +
              'INNER JOIN tblOrdenes O ON S.ITE_Nombre  = O.ITE_Nombre ' +
              'LEFT OUTER JOIN tblEmpleados E ON E.[Id] = S.SCR_EmpleadoRes ';

    if chkRecibido.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Recibido >= ' + QuotedStr(deRecibido1.Text) +
                    ' and O.Recibido <= ' + QuotedStr(deRecibido2.Text) + ') ';
    end;

    if chkInterna.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Interna >= ' + QuotedStr(deInterna1.Text) +
                    ' and O.Interna <= ' + QuotedStr(deInterna2.Text) + ') ';
    end;

    if chkEntrega.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Entrega >= ' + QuotedStr(deEntrega1.Text) +
                    ' and O.Entrega <= ' + QuotedStr(deEntrega2.Text) + ') ';
    end;

    if chkScrap.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (S.SCR_Fecha >= ' + QuotedStr(deScrap1.Text) +
                    ' and S.SCR_Fecha, <= ' + QuotedStr(deScrap2.Text) + ') ';
    end;

    if cmbProductos.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' O.Producto = ' + QuotedStr(cmbProductos.Text) + ' ';
    end;

    if cmbClientes.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' SUBSTRING(O.ITE_Nombre,4,3) = ' + QuotedStr(cmbClientes.Text) + ' ';
    end;

    if cmbOrdenes.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' SUBSTRING(O.ITE_Nombre,8,3) = ' + QuotedStr(cmbOrdenes.Text) + ' ';
    end;

    if cmbPartes.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' O.Numero = ' + QuotedStr(cmbPartes.Text) + ' ';
    end;

    if SQLWhere <> '' then SQLStr := SQLStr + ' WHERE ' + SQLWhere;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Orden']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['Fecha']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['NewOrden']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['SCR_Fecha']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Area']);
        GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['EmpleadoRes']);
        GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(Qry['SCR_Motivo']);
        Qry.Next;
    End;

    BindPartes(SQLWhere);
    BindOrdenes(SQLWhere);
end;


procedure TfrmScrapReport.BindProductos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblProductos Order By Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbProductos.Items.Clear;
    cmbProductos.Items.Add('Todos');
    While not Qry2.Eof do
    Begin
        cmbProductos.Items.Add(Qry2['Nombre']);
        Qry2.Next;
    End;

    cmbProductos.Text := 'Todos';
    Qry2.Close;
end;

procedure TfrmScrapReport.BindClientes();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Clave FROM tblClientes Order By Clave';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbClientes.Items.Clear;
    cmbClientes.Items.Add('Todos');
    While not Qry2.Eof do
    Begin
        cmbClientes.Items.Add(Qry2['Clave']);
        Qry2.Next;
    End;

    cmbClientes.Text := 'Todos';
    Qry2.Close;
end;

procedure TfrmScrapReport.BindPartes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;


    SQLStr := 'SELECT DISTINCT(Numero) AS Numero ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';

    if Query <> '' then SQLStr := SQLStr + ' WHERE ' + Query;

    SQLStr := SQLStr + ' ORDER BY Numero ';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbPartes.Items.Clear;
    cmbPartes.Items.Add('Todos');
    While not Qry2.Eof do
    Begin
        cmbPartes.Items.Add(Qry2['Numero']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmScrapReport.BindOrdenes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;


    SQLStr := 'SELECT DISTINCT SUBSTRING(O.ITE_Nombre,8,3) AS Orden ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';

    if Query <> '' then SQLStr := SQLStr + ' WHERE ' + Query;

    SQLStr := SQLStr + ' ORDER BY SUBSTRING(O.ITE_Nombre,8,3) ';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbOrdenes.Items.Clear;
    cmbOrdenes.Items.Add('Todos');
    While not Qry2.Eof do
    Begin
        cmbOrdenes.Items.Add(Qry2['Orden']);
        Qry2.Next;
    End;

    Qry2.Close;
end;


procedure TfrmScrapReport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmScrapReport.chkRecibidoClick(Sender: TObject);
begin
deRecibido1.Enabled := chkRecibido.Checked;
deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmScrapReport.chkInternaClick(Sender: TObject);
begin
deInterna1.Enabled := chkInterna.Checked;
deInterna2.Enabled := chkInterna.Checked;
end;

procedure TfrmScrapReport.chkEntregaClick(Sender: TObject);
begin
deEntrega1.Enabled := chkEntrega.Checked;
deEntrega2.Enabled := chkEntrega.Checked;
end;

procedure TfrmScrapReport.chkScrapClick(Sender: TObject);
begin
deScrap1.Enabled := chkScrap.Checked;
deScrap2.Enabled := chkScrap.Checked;
end;

procedure TfrmScrapReport.ExportGrid(Grid: TGridView;sFileName: String);
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

procedure TfrmScrapReport.btnBuscarClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmScrapReport.Button1Click(Sender: TObject);
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

procedure TfrmScrapReport.Button2Click(Sender: TObject);
begin

    Application.Initialize;
    Application.CreateForm(TqrReporteScrap, qrReporteScrap);
    qrReporteScrap.ReportTitle.Caption := 'Reporte de Scrap';

    qrReporteScrap.QRSubDetail1.DataSet := Qry;
    qrReporteScrap.Field1.DataSet := Qry;
    qrReporteScrap.Field1.DataField := 'Orden';

    qrReporteScrap.Field2.DataSet := Qry;
    qrReporteScrap.Field2.DataField := 'Cantidad';

    qrReporteScrap.Field3.DataSet := Qry;
    qrReporteScrap.Field3.DataField := 'Descripcion';

    qrReporteScrap.Field4.DataSet := Qry;
    qrReporteScrap.Field4.DataField := 'Numero';

    qrReporteScrap.Field5.DataSet := Qry;
    qrReporteScrap.Field5.DataField := 'Fecha';

    qrReporteScrap.Field6.DataSet := Qry;
    qrReporteScrap.Field6.DataField := 'NewOrden';

    qrReporteScrap.Field7.DataSet := Qry;
    qrReporteScrap.Field7.DataField := 'SCR_Fecha';

    qrReporteScrap.Field8.DataSet := Qry;
    qrReporteScrap.Field8.DataField := 'Area';

    qrReporteScrap.Field9.DataSet := Qry;
    qrReporteScrap.Field9.DataField := 'EmpleadoRes';

    qrReporteScrap.Preview;
    qrReporteScrap.Free;
end;

function TfrmScrapReport.FormIsRunning(FormName: String):Boolean;
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


end.
