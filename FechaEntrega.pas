unit FechaEntrega;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint;

type
  TfrmEntrega = class(TForm)
    GridView1: TGridView;
    gbSearch: TGroupBox;
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
    Label1: TLabel;
    Label2: TLabel;
    cmbClientes: TComboBox;
    Label3: TLabel;
    cmbPartes: TComboBox;
    Button2: TButton;
    btnBuscar: TButton;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    cmbOrdenes: TComboBox;
    Label4: TLabel;
    lblCount: TLabel;
    GroupBox3: TGroupBox;
    gvOrdenes: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    Button3: TButton;
    txtCliente: TEdit;
    txtOrden: TEdit;
    Button7: TButton;
    txtProducto: TEdit;
    Button4: TButton;
    GroupBox2: TGroupBox;
    gvProds: TGridView;
    CheckBox3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    Button5: TButton;
    Button6: TButton;
    Label5: TLabel;
    GroupBox4: TGroupBox;
    gvTareas: TGridView;
    CheckBox4: TCheckBox;
    btnOK4: TButton;
    btnTodos4: TButton;
    txtTarea: TEdit;
    Button8: TButton;
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    Procedure BindPartes(Query: String);
    Procedure BindTareas(Query: String);    
    Procedure BindOrdenes(Query: String);
    procedure BindGrid();
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnBuscarClick(Sender: TObject);
    procedure GridView1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Timer1Timer(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure chkInternaClick(Sender: TObject);
    procedure chkRecibidoClick(Sender: TObject);
    procedure chkEntregaClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure btnOK4Click(Sender: TObject);
    procedure btnTodos4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEntrega: TfrmEntrega;
  giCantidad, giCantCliente:Integer;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main;

{$R *.dfm}

procedure TfrmEntrega.FormCreate(Sender: TObject);
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

    cmbProductos.Text := 'Todos';
    cmbClientes.Text := 'Todos';
    cmbPartes.Text := 'Todos';
    cmbOrdenes.Text := 'Todos';

    BindClientes();
    BindTareas('');
    CheckBox1.Checked := False;
    btnOKClick(nil);
    BindGrid();
    BindProductos();

    Qry2.Close;
end;

procedure TfrmEntrega.BindGrid();
var SQLStr,SQLWhere, SQLWhere2 : String;
begin
    lblCount.Caption := '';
    giCantidad := 0;
    giCantCliente := 0;
    SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Requerida As Cliente, O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, Entrega As Compromiso, ' +
              'T.Nombre AS Tarea,CASE WHEN I.ITS_Status = 0 THEN ''Listo'' ' +
              'WHEN I.ITS_Status = 1 THEN ''Activo'' ' +
              'WHEN I.ITS_Status = 2 THEN ''Terminado'' END AS Status, ' +
              'O.Observaciones,I.ITS_DTStart,E.Nombre,O.OrdenCompra, O.Recibido ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ' +
              'LEFT OUTER JOIN tblEmpleados E ON I.[USE_Login] = E.[ID] WHERE I.ITS_Status <> 9 ';

    SQLWhere := '';
    SQLWhere2 := '';
    if chkRecibido.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Recibido >= ' + QuotedStr(deRecibido1.Text) +
                    ' AND O.Recibido <= ' + QuotedStr(deRecibido2.Text) + ') ';
    end;

    if chkInterna.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Interna >= ' + QuotedStr(deInterna1.Text) +
                    ' AND O.Interna <= ' + QuotedStr(deInterna2.Text) + ') ';
    end;

    if chkEntrega.Checked then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' (O.Entrega >= ' + QuotedStr(deEntrega1.Text) +
                    ' AND O.Entrega <= ' + QuotedStr(deEntrega2.Text) + ') ';
    end;

    if txtProducto.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' O.Producto IN (''' +
        StringReplace(txtProducto.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtCliente.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' SUBSTRING(O.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtOrden.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' SUBSTRING(O.ITE_Nombre,8,3) IN (''' +
        StringReplace(txtOrden.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if cmbPartes.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        //if Pos('*', cmbPartes.Text) > 0 then 
        SQLWhere := SQLWhere + ' O.Numero = ' + QuotedStr(cmbPartes.Text) + ' ';
    end;

    if txtTarea.Text <> 'Todos' then
    begin
        if SQLWhere2 <> '' then SQLWhere2 := SQLWhere2 + ' AND ';
        //SQLWhere2 := SQLWhere2 + ' T.Nombre = ' + QuotedStr(cmbTareas.Text) + ' ';
        SQLWhere2 := SQLWhere2 + ' T.Nombre IN (''' +
        StringReplace(txtTarea.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if SQLWhere <> '' then SQLStr := SQLStr + ' AND ' + SQLWhere;
    if SQLWhere2 <> '' then SQLStr := SQLStr + ' AND ' + SQLWhere2;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Orden']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);
        giCantidad := giCantidad + StrToInt(GridView1.Cells[1,GridView1.RowCount -1]);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Cliente']);
        giCantCliente := giCantCliente + StrToInt(GridView1.Cells[2,GridView1.RowCount -1]);

        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Fecha']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Compromiso']);
        GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['Recibido']);

        GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(Qry['Tarea']);
        GridView1.Cells[11,GridView1.RowCount -1] := VarToStr(Qry['Status']);
        GridView1.Cells[12,GridView1.RowCount -1] := VarToStr(Qry['ITS_DTStart']);
        GridView1.Cells[13,GridView1.RowCount -1] := VarToStr(Qry['Nombre']);
        GridView1.Cells[14,GridView1.RowCount -1] := VarToStr(Qry['Observaciones']);
        Qry.Next;
    End;

    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(GridView1.RowCount) +
                        ' Cantidad Total Larco : ' + IntToStr(giCantidad) +
                        ' Cantidad Total Cliente : ' + IntToStr(giCantCliente);
    BindPartes(SQLWhere);
    BindOrdenes(SQLWhere);
end;


procedure TfrmEntrega.BindProductos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblProductos Order By Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvProds.ClearRows;
    While not Qry2.Eof do
    Begin
        gvProds.AddRow(1);
        gvProds.Cells[0,gvProds.RowCount -1] := VarToStr(Qry2['Nombre']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmEntrega.BindClientes();
var Qry2 : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Distinct Clave FROM tblClientes Order By Clave';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvClientes.ClearRows;
    While not Qry2.Eof do
    Begin
        gvClientes.AddRow(1);
        gvClientes.Cells[0,gvClientes.RowCount -1] := VarToStr(Qry2['Clave']);

        if (slClientes.IndexOf(VarToStr(Qry2['Clave'])) = -1) then begin
                gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
        end;

        Qry2.Next;
    End;
    
    Qry2.Close;
end;

procedure TfrmEntrega.BindPartes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;


    SQLStr := 'SELECT DISTINCT(Numero) AS Numero ' +
              'FROM tblOrdenes O ';// +
//              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
//              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';

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

procedure TfrmEntrega.BindTareas(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;


    SQLStr := 'SELECT DISTINCT(Nombre) AS Nombre FROM tblTareas T Order by Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvTareas.ClearRows;
    while not Qry2.Eof do
    begin
        gvTareas.AddRow(1);
        gvTareas.Cells[0,gvTareas.RowCount -1] := VarToStr(Qry2['Nombre']);

        Qry2.Next;
    end;

    Qry2.Close;
end;

procedure TfrmEntrega.BindOrdenes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
i : Integer;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;


    SQLStr := 'SELECT DISTINCT SUBSTRING(O.ITE_Nombre,8,3) AS Orden ' +
              'FROM tblOrdenes O '; //+
//              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
//              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';

    if Query <> '' then SQLStr := SQLStr + ' WHERE ' + Query;

    SQLStr := SQLStr + ' ORDER BY SUBSTRING(O.ITE_Nombre,8,3) ';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbOrdenes.Items.Clear;
    cmbOrdenes.Sorted := True;
    While not Qry2.Eof do
    Begin
        cmbOrdenes.Items.Add(Qry2['Orden']);
        Qry2.Next;
    End;

    gvOrdenes.ClearRows;
    for i:= 0 to cmbOrdenes.Items.Count - 1 do begin
        gvOrdenes.AddRow(1);
        gvOrdenes.Cells[0,gvOrdenes.RowCount -1] := cmbOrdenes.Items.Strings[i];
    end;

    Qry2.Close;
end;


procedure TfrmEntrega.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmEntrega.btnBuscarClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmEntrega.GridView1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
                BindGrid;

end;

procedure TfrmEntrega.Timer1Timer(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmEntrega.Button1Click(Sender: TObject);
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

procedure TfrmEntrega.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Ordenes';

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


procedure TfrmEntrega.Button2Click(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrRelacionEntrega, qrRelacionEntrega);
    qrRelacionEntrega.ReportTitle.Caption := 'Relacion de Fecha de Entrega ';

    qrRelacionEntrega.QRSubDetail1.DataSet := Qry;
    qrRelacionEntrega.Field1.DataSet := Qry;
    qrRelacionEntrega.Field1.DataField := 'Orden';

    qrRelacionEntrega.Field2.DataSet := Qry;
    qrRelacionEntrega.Field2.DataField := 'Cantidad';

    qrRelacionEntrega.Field3.DataSet := Qry;
    qrRelacionEntrega.Field3.DataField := 'Descripcion';

    qrRelacionEntrega.Field4.DataSet := Qry;
    qrRelacionEntrega.Field4.DataField := 'Numero';

    qrRelacionEntrega.Field5.DataSet := Qry;
    qrRelacionEntrega.Field5.DataField := 'Terminal';

    qrRelacionEntrega.Field6.DataSet := Qry;
    qrRelacionEntrega.Field6.DataField := 'Fecha';

    qrRelacionEntrega.Field7.DataSet := Qry;
    qrRelacionEntrega.Field7.DataField := 'Tarea';

    qrRelacionEntrega.Field8.DataSet := Qry;
    qrRelacionEntrega.Field8.DataField := 'Status';

    qrRelacionEntrega.Field9.DataSet := Qry;
    qrRelacionEntrega.Field9.DataField := 'Observaciones';

    qrRelacionEntrega.Preview;
    qrRelacionEntrega.Free;
end;

procedure TfrmEntrega.chkInternaClick(Sender: TObject);
begin
deInterna1.Enabled := chkInterna.Checked;
deInterna2.Enabled := chkInterna.Checked;
end;

procedure TfrmEntrega.chkRecibidoClick(Sender: TObject);
begin
deRecibido1.Enabled := chkRecibido.Checked;
deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmEntrega.chkEntregaClick(Sender: TObject);
begin
deEntrega1.Enabled := chkEntrega.Checked;
deEntrega2.Enabled := chkEntrega.Checked;
end;

procedure TfrmEntrega.Button3Click(Sender: TObject);
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

procedure TfrmEntrega.Button7Click(Sender: TObject);
begin
  if GroupBox3.Visible = True then
  begin
          GroupBox3.Visible := False;
  end
  else begin
      GroupBox3.Visible := True;
      if txtOrden.Text = 'Todos' then
      begin
              CheckBox2.Checked := True;
              gvOrdenes.Enabled := False;
              btnTodos2.Enabled := False;
      end
      else
      begin
              CheckBox2.Checked := False;
              gvOrdenes.Enabled := True;
              btnTodos2.Enabled := True;
      end;

      GroupBox3.Top := txtOrden.Top + txtOrden.Height + 5;
      GroupBox3.Left := txtOrden.Left + 10;
  end;
end;

procedure TfrmEntrega.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmEntrega.CheckBox2Click(Sender: TObject);
begin
gvOrdenes.Enabled := not CheckBox2.Checked;
btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmEntrega.btnOKClick(Sender: TObject);
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

procedure TfrmEntrega.btnOK2Click(Sender: TObject);
var i: integer;
sOrdenes : String;
begin
  GroupBox3.Visible := False;
  if CheckBox2.Checked = True then begin
          txtOrden.Text := 'Todos';
  end
  else begin
        sOrdenes := '';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                if gvOrdenes.Cell[1,i].AsBoolean = True then
                begin
                        sOrdenes := sOrdenes + gvOrdenes.Cells[0,i] + ',';
                end;
        end;
        txtOrden.Text := 'Todos';
        if sOrdenes <> '' then
        begin
                txtOrden.Text :=  LeftStr(sOrdenes,Length(sOrdenes) - 1);
        end;
  end;

end;

procedure TfrmEntrega.btnTodosClick(Sender: TObject);
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

procedure TfrmEntrega.btnTodos2Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos2.Caption) = UT('Seleccionar Todos') then begin
        btnTodos2.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                gvOrdenes.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos2.Caption := 'Seleccionar Todos';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                gvOrdenes.Cell[1,i].AsBoolean := False;
        end;
  end;

end;

procedure TfrmEntrega.Button4Click(Sender: TObject);
begin
  if GroupBox2.Visible = True then
  begin
          GroupBox2.Visible := False;
  end
  else begin
      GroupBox2.Visible := True;
      if txtProducto.Text = 'Todos' then
      begin
              CheckBox3.Checked := True;
              gvProds.Enabled := False;
              btnTodos3.Enabled := False;
      end
      else
      begin
              CheckBox3.Checked := False;
              gvProds.Enabled := True;
              btnTodos3.Enabled := True;
      end;

      GroupBox2.Top := txtProducto.Top + txtProducto.Height + 5;
      GroupBox2.Left := txtProducto.Left + 10;
  end;

end;

procedure TfrmEntrega.CheckBox3Click(Sender: TObject);
begin
gvProds.Enabled := not CheckBox3.Checked;
btnTodos3.Enabled := not CheckBox3.Checked;
end;

procedure TfrmEntrega.btnOK3Click(Sender: TObject);
var i: integer;
sProductos : String;
begin
  GroupBox2.Visible := False;
  if CheckBox3.Checked = True then begin
          txtProducto.Text := 'Todos';
  end
  else begin
        sProductos := '';
        for i:= 0 to gvProds.RowCount - 1 do
        begin
                if gvProds.Cell[1,i].AsBoolean = True then
                begin
                        sProductos := sProductos + gvProds.Cells[0,i] + ',';
                end;
        end;
        txtProducto.Text := 'Todos';
        if sProductos <> '' then
        begin
                txtProducto.Text :=  LeftStr(sProductos,Length(sProductos) - 1);
        end;
  end;
end;

procedure TfrmEntrega.btnTodos3Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos3.Caption) = UT('Seleccionar Todos') then begin
        btnTodos3.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvProds.RowCount - 1 do
        begin
                gvProds.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos3.Caption := 'Seleccionar Todos';
        for i:= 0 to gvProds.RowCount - 1 do
        begin
                gvProds.Cell[1,i].AsBoolean := False;
        end;
  end;


end;

procedure TfrmEntrega.Button5Click(Sender: TObject);
var i:Integer;
ChildControl: TControl;
begin
  for i:=0 to gbSearch.ControlCount - 1 do
  begin
        ChildControl := gbSearch.Controls[i];
        ShowMessage('Name:' + ChildControl.Name + ' ClassName:' + ChildControl.ClassName);
        if(ChildControl.ClassName = 'TCheckBox') then begin
                if (ChildControl as TCheckBox).Checked then
                        ShowMessage('checked')
                else
                        ShowMessage('Not checked');
        end;
  end;
end;

procedure TfrmEntrega.Button6Click(Sender: TObject);
var ChildControl: TControl;
begin
  ChildControl := gbSearch.FindChildControl('chkRecibido');
  if(ChildControl.ClassName = 'TCheckBox') then begin
   (ChildControl as TCheckBox).Checked := True;
  end;

end;

procedure TfrmEntrega.Button8Click(Sender: TObject);
begin
  if GroupBox4.Visible = True then
  begin
          GroupBox4.Visible := False;
  end
  else begin
      GroupBox4.Visible := True;
      if txtTarea.Text = 'Todos' then
      begin
              CheckBox4.Checked := True;
              gvTareas.Enabled := False;
              btnTodos4.Enabled := False;
      end
      else
      begin
              CheckBox4.Checked := False;
              gvTareas.Enabled := True;
              btnTodos4.Enabled := True;
      end;

      GroupBox4.Top := txtTarea.Top + txtTarea.Height + 5;
      GroupBox4.Left := txtTarea.Left + 10;
  end;
end;

procedure TfrmEntrega.CheckBox4Click(Sender: TObject);
begin
gvTareas.Enabled := not CheckBox4.Checked;
btnTodos4.Enabled := not CheckBox4.Checked;
end;

procedure TfrmEntrega.btnOK4Click(Sender: TObject);
var i: integer;
sTareas : String;
begin
  GroupBox4.Visible := False;
  if CheckBox4.Checked = True then begin
          txtTarea.Text := 'Todos';
  end
  else begin
        sTareas := '';
        for i:= 0 to gvTareas.RowCount - 1 do
        begin
                if gvTareas.Cell[1,i].AsBoolean = True then
                begin
                        sTareas := sTareas + gvTareas.Cells[0,i] + ',';
                end;
        end;
        txtTarea.Text := 'Todos';
        if sTareas <> '' then
        begin
                txtTarea.Text :=  LeftStr(sTareas,Length(sTareas) - 1);
        end;
  end;
end;

procedure TfrmEntrega.btnTodos4Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos4.Caption) = UT('Seleccionar Todos') then begin
        btnTodos4.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvTareas.RowCount - 1 do
        begin
                gvTareas.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos4.Caption := 'Seleccionar Todos';
        for i:= 0 to gvTareas.RowCount - 1 do
        begin
                gvTareas.Cell[1,i].AsBoolean := False;
        end;
  end;

end;

end.
