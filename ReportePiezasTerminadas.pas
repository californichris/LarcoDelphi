unit ReportePiezasTerminadas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint, Menus;

type
  TfrmPiezasTerminadas = class(TForm)
    GroupBox1: TGroupBox;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Label1: TLabel;
    txtProducto: TEdit;
    Button4: TButton;
    Button3: TButton;
    txtCliente: TEdit;
    Label2: TLabel;
    Label4: TLabel;
    txtOrden: TEdit;
    Button7: TButton;
    Label3: TLabel;
    btnBuscar: TButton;
    GridView1: TGridView;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    GroupBox3: TGroupBox;
    gvOrdenes: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    GroupBox2: TGroupBox;
    gvProds: TGridView;
    CheckBox3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    cmbOrdenes: TComboBox;
    lblCount: TLabel;
    SaveDialog1: TSaveDialog;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    Procedure BindOrdenes(Query: String);
    procedure BindGrid();
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPiezasTerminadas: TfrmPiezasTerminadas;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main;

{$R *.dfm}

procedure TfrmPiezasTerminadas.FormCreate(Sender: TObject);
begin
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  deFrom.Date := Now;
  deTo.Date := Now;

  BindGrid();
  BindProductos();
  BindClientes();

end;

procedure TfrmPiezasTerminadas.BindGrid();
var SQLStr,SQLWhere : String;
giCantidad,giCantCliente:Integer;
begin
    giCantidad := 0;
    giCantCliente := 0;

{    SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden,O.Ordenada As Cantidad, ' +
              'O.Requerida As Cliente, O.Producto As Descripcion,O.Numero,O.Terminal,Interna As Fecha, ' +
              'T.Nombre AS Tarea,CASE WHEN I.ITS_Status = 0 THEN ''Listo'' ' +
              'WHEN I.ITS_Status = 1 THEN ''Activo'' ' +
              'WHEN I.ITS_Status = 2 THEN ''Terminado'' END AS Status, ' +
              'O.Observaciones,I.ITS_DTStart,E.Nombre ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ' +
              'LEFT OUTER JOIN tblEmpleados E ON I.[USE_Login] = E.[ID] ';
}
    SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden, O.* ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_Id = T.[Id] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ''VentasFinal'' AND ITS_Status = 2 ' +
              'AND ITS_DTStop >= ' + QuotedStr(deFrom.Text) +
              ' AND ITS_DTStop <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99');


    SQLWhere := ' ';

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

    if SQLWhere <> '' then SQLStr := SQLStr + ' ' + SQLWhere;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Orden']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Ordenada']);
        giCantidad := giCantidad + StrToInt(GridView1.Cells[1,GridView1.RowCount -1]);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Requerida']);
        giCantCliente := giCantCliente + StrToInt(GridView1.Cells[2,GridView1.RowCount -1]);

        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Producto']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['Terminal']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Interna']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Observaciones']);
        Qry.Next;
    End;

    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(GridView1.RowCount) +
                        '   Cantidad Piezas Larco : ' + IntToStr(giCantidad) +
                        '   Cantidad Piezas Cliente : ' + IntToStr(giCantCliente) +
                        '   Diferencia : ' + IntToStr(giCantidad - giCantCliente);
    
    BindOrdenes(SQLWhere);
end;

procedure TfrmPiezasTerminadas.BindOrdenes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
i : Integer;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

{    SQLStr := 'SELECT DISTINCT SUBSTRING(O.ITE_Nombre,8,3) AS Orden ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';
}

    SQLStr := 'SELECT DISTINCT SUBSTRING(O.ITE_Nombre,8,3) AS Orden ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_Id = T.[Id] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ''VentasFinal'' AND ITS_Status = 2 ' +
              'AND ITS_DTStop >= ' + QuotedStr(deFrom.Text) +
              ' AND ITS_DTStop <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99');


    if Query <> '' then SQLStr := SQLStr + ' ' + Query;

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

procedure TfrmPiezasTerminadas.BindProductos();
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

procedure TfrmPiezasTerminadas.BindClientes();
var Qry2 : TADOQuery;
SQLStr : String;
begin
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
        Qry2.Next;
    End;
    
    Qry2.Close;
end;

procedure TfrmPiezasTerminadas.CheckBox1Click(Sender: TObject);
begin
  gvClientes.Enabled := not CheckBox1.Checked;
  btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmPiezasTerminadas.CheckBox2Click(Sender: TObject);
begin
  gvOrdenes.Enabled := not CheckBox2.Checked;
  btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmPiezasTerminadas.CheckBox3Click(Sender: TObject);
begin
  gvProds.Enabled := not CheckBox3.Checked;
  btnTodos3.Enabled := not CheckBox3.Checked;
end;

procedure TfrmPiezasTerminadas.btnOKClick(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnOK2Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnOK3Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnTodosClick(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnTodos2Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnTodos3Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmPiezasTerminadas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmPiezasTerminadas.Button4Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.Button3Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.Button7Click(Sender: TObject);
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

procedure TfrmPiezasTerminadas.btnBuscarClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmPiezasTerminadas.Exportar1Click(Sender: TObject);
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

end.
