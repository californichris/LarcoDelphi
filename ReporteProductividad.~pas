unit ReporteProductividad;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,
  Menus, Larco_functions, TeEngine, Series, TeeProcs, Chart, Clipbrd;

type
  TfrmProductividad = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    lblAnio: TLabel;
    btnDetalle: TButton;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    btnImprimir: TButton;
    btnBuscar: TButton;
    txtCliente: TEdit;
    txtProducto: TEdit;
    btnDesc: TButton;
    btnClientes: TButton;
    gbDesc: TGroupBox;
    gvDescs: TGridView;
    chkDesc: TCheckBox;
    btnDescOK: TButton;
    btnTodosDesc: TButton;
    gbClientes: TGroupBox;
    gvClientes: TGridView;
    chkClientes: TCheckBox;
    btnClientesOK: TButton;
    btnTodosClientes: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    CopiarOrden1: TMenuItem;
    CopiarComo1: TMenuItem;
    Separadoporcomas1: TMenuItem;
    Encomillas1: TMenuItem;
    OpenDialog1: TOpenDialog;
    SaveDialog2: TSaveDialog;
    Label3: TLabel;
    Label5: TLabel;
    GroupBox2: TGroupBox;
    Label6: TLabel;
    lblEntraron: TLabel;
    Label8: TLabel;
    lblSalieron: TLabel;
    Label7: TLabel;
    lblHabia: TLabel;
    cmbTarea1: TComboBox;
    cmbTarea2: TComboBox;
    Label9: TLabel;
    gbDetalle: TGroupBox;
    lblEntraron3: TLabel;
    lblSalieron3: TLabel;
    lblHabia3: TLabel;
    GridView1: TGridView;
    GridView2: TGridView;
    GridView3: TGridView;
    Chart1: TChart;
    Series1: TPieSeries;
    Label13: TLabel;
    lblEntraron2: TLabel;
    Label15: TLabel;
    lblSalieron2: TLabel;
    Label17: TLabel;
    lblHabia2: TLabel;
    Chart2: TChart;
    PieSeries1: TPieSeries;
    Label19: TLabel;
    Label20: TLabel;
    lblCargando: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    procedure BindTareas();
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
      Grid:TGridView; Button: TButton);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure btnDescClick(Sender: TObject);
    procedure btnClientesClick(Sender: TObject);
    procedure chkDescClick(Sender: TObject);
    procedure chkClientesClick(Sender: TObject);
    procedure btnDescOKClick(Sender: TObject);
    procedure btnClientesOKClick(Sender: TObject);
    procedure btnTodosDescClick(Sender: TObject);
    procedure btnTodosClientesClick(Sender: TObject);
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure btnBuscarClick(Sender: TObject);
    procedure BindGridAuto(SQLStr: String;  Grid:TGridView);
    function  getCantidad(column:Integer; Grid:TGridView):Integer;
    procedure btnDetalleClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure Separadoporcomas1Click(Sender: TObject);
    procedure Encomillas1Click(Sender: TObject);
    procedure deFromChange(Sender: TObject);
    procedure cmbTarea1Change(Sender: TObject);
    procedure AddToGridAuto(SQLStr: String;  Grid:TGridView);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProductividad: TfrmProductividad;
  Conn : TADOConnection;
  Qry : TADOQuery;
  
implementation

uses Main;

{$R *.dfm}

procedure TfrmProductividad.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmProductividad.FormCreate(Sender: TObject);
begin
    lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    deFrom.Date := Now;
    deTo.Date := Now;

    BindTareas();
    cmbTarea1.ItemIndex := 0;
    cmbTarea2.ItemIndex := 16;
    cmbTarea1.OnChange := cmbTarea1Change;
    cmbTarea2.OnChange := cmbTarea1Change;
    deFrom.OnChange := deFromChange;
    deTo.OnChange := deFromChange;

    BindClientes();
    chkClientes.Checked := False;
    btnClientesOKClick(nil);

    btnBuscarClick(nil);
end;

procedure TfrmProductividad.BindProductos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblProductos Order By Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvDescs.ClearRows;
    While not Qry2.Eof do
    Begin
        gvDescs.AddRow(1);
        gvDescs.Cells[0,gvDescs.RowCount -1] := VarToStr(Qry2['Nombre']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmProductividad.BindClientes();
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

procedure TfrmProductividad.BindTareas();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblTareas Order By TAS_Order';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbTarea1.Items.Clear;
    cmbTarea2.Items.Clear;
    While not Qry2.Eof do
    Begin
        cmbTarea1.Items.Add(VarToStr(Qry2['Nombre']));
        cmbTarea2.Items.Add(VarToStr(Qry2['Nombre']));
        Qry2.Next;
    End;

    cmbTarea1.Text := cmbTarea1.Items[0];
    cmbTarea2.Text := 'VentasFinal';
    Qry2.Close;
end;


procedure TfrmProductividad.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
          TextBox: TEdit; Grid: TGridView; Button: TButton);
begin
  if GroupBox.Visible = True then
  begin
          GroupBox.Visible := False;
  end
  else begin
      GroupBox.Width := TextBox.Width;
      GroupBox.Top := txtProducto.Parent.Top + TextBox.Top + TextBox.Height;
      GroupBox.Left := TextBox.Left + 8;
      Grid.Width := GroupBox.Width - 12;
      Button.Width := GroupBox.Width - 12;

      GroupBox.Visible := True;
      CheckBox.Checked := False;
      Grid.Enabled := True;
      Button.Enabled := True;
  end;

end;

procedure TfrmProductividad.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
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

procedure TfrmProductividad.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid: TGridView);
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


procedure TfrmProductividad.btnDescClick(Sender: TObject);
begin
  gbClientes.Visible := False;
  ShowSeleccionGrid(gbDesc, chkDesc, txtProducto, gvDescs, btnTodosDesc);
  if gbDesc.Visible = True then
      BindProductos();

end;

procedure TfrmProductividad.btnClientesClick(Sender: TObject);
begin
  gbDesc.Visible := False;
  ShowSeleccionGrid(gbClientes, chkClientes, txtCliente, gvClientes, btnTodosClientes);
  //if gbClientes.Visible = True then
  //    BindClientes();

end;

procedure TfrmProductividad.chkDescClick(Sender: TObject);
begin
  gvDescs.Enabled := not chkDesc.Checked;
  btnTodosDesc.Enabled := not chkDesc.Checked;
end;

procedure TfrmProductividad.chkClientesClick(Sender: TObject);
begin
  gvClientes.Enabled := not chkClientes.Checked;
  btnTodosClientes.Enabled := not chkClientes.Checked;
end;

procedure TfrmProductividad.btnDescOKClick(Sender: TObject);
begin
  ParseSelection(gbDesc,chkDesc,txtProducto,gvDescs);
end;

procedure TfrmProductividad.btnClientesOKClick(Sender: TObject);
begin
  ParseSelection(gbClientes,chkClientes,txtCliente,gvClientes);
end;

procedure TfrmProductividad.btnTodosDescClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosDesc, gvDescs);
end;

procedure TfrmProductividad.btnTodosClientesClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosClientes, gvClientes);
end;

procedure TfrmProductividad.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmProductividad.btnBuscarClick(Sender: TObject);
var SQLStr,SQLWhere : String;
begin

    if cmbTarea2.ItemIndex < cmbTarea1.ItemIndex then begin
       ShowMessage('La tarea final esta antes que la tarea inicial, seleccione el orden correctamente.');
       Exit;
    end;

    lblCargando.Visible := True;
    btnBuscar.Enabled := False;
    btnDetalle.Enabled := False;
    btnImprimir.Enabled := False;
    Application.ProcessMessages;

    SQLWhere := 'UPDATE tblItemTasks SET USE_LOGIN = ''25'' WHERE USE_LOGIN = ''System''';
    Conn.Execute(SQLWhere);

    SQLWhere := '';

    if txtProducto.Text <> 'Todos' then
    begin
        SQLWhere := SQLWhere + ' AND O.Producto IN (''' +
        StringReplace(txtProducto.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtCliente.Text <> 'Todos' then
    begin
        SQLWhere := SQLWhere + ' AND SUBSTRING(O.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    SQLStr := 'SELECT ' +
              'RIGHT(I.ITE_Nombre,10) AS Orden, O.Requerida As [Cant.Cliente], O.Ordenada As [Cant.Larco], ' +
              'O.Producto As Descripcion,O.Numero AS [No.Parte],O.Terminal AS Term, ' +
              'O.Interna As [F.Interna], I.ITS_DTStart [F.Entrada], O.Unitario, CASE WHEN O.Dolares = 1 THEN ''1'' ELSE ''0'' END AS [Dolares] ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ' + QuotedStr(cmbTarea1.Text) + ' ' +
              'AND (I.ITS_DTStart >= ' + QuotedStr(deFrom.Text) +
              ' AND I.ITS_DTStart <= ' + QuotedStr(deTo.Text  + ' 23:59:59.999' ) + ') ' +
              SQLWhere;

    BindGridAuto(SQLStr, GridView1);

    SQLStr := 'SELECT ' +
              'RIGHT(I.ITE_Nombre,10) AS Orden, O.Requerida As [Cant.Cliente], O.Ordenada As [Cant.Larco], ' +
              'O.Producto As Descripcion,O.Numero AS [No.Parte],O.Terminal AS Term, ' +
              'O.Interna As [F.Interna], I.ITS_DTStart [F.Entrada], I.ITS_DTStop [F.Salida],E.Nombre AS Empleado, O.Unitario, CASE WHEN O.Dolares = 1 THEN ''1'' ELSE ''0'' END AS [Dolares] ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'INNER JOIN tblEmpleados E ON I.USE_Login = E.[ID] ' +
              'WHERE T.Nombre = ' + QuotedStr(cmbTarea2.Text) + ' ' +
              'AND (I.ITS_DTStop >= ' + QuotedStr(deFrom.Text) +
              ' AND I.ITS_DTStop <= ' + QuotedStr(deTo.Text + ' 23:59:59.999' ) + ') ' +
              SQLWhere;

    BindGridAuto(SQLStr, GridView2);

    SQLStr := 'SELECT ' +
              'RIGHT(I.ITE_Nombre,10) AS Orden, O.Requerida As [Cant.Cliente], O.Ordenada As [Cant.Larco], ' +
              'O.Producto As Descripcion,O.Numero AS [No.Parte],O.Terminal AS Term, ' +
              'O.Interna As [F.Interna], I.ITS_DTStart [F.Entrada], I.ITS_DTStop [F.Salida], O.Unitario, CASE WHEN O.Dolares = 1 THEN ''1'' ELSE ''0'' END AS [Dolares] ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE I.ITS_Status NOT IN (9,2) AND T.Nombre = ' + QuotedStr(cmbTarea1.Text) + ' ' +
              'AND I.ITS_DTStart < ' + QuotedStr(deFrom.Text) + ' ' +
              SQLWhere;

    BindGridAuto(SQLStr, GridView3);

    SQLStr := 'SELECT ' +
              'RIGHT(I.ITE_Nombre,10) AS Orden, O.Requerida As [Cant.Cliente], O.Ordenada As [Cant.Larco], ' +
              'O.Producto As Descripcion,O.Numero AS [No.Parte],O.Terminal AS Term, ' +
              'O.Interna As [F.Interna], I.ITS_DTStart [F.Entrada], I.ITS_DTStop [F.Salida], O.Unitario, CASE WHEN O.Dolares = 1 THEN ''1'' ELSE ''0'' END AS [Dolares]  ' +
              'FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
              'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ' +
              'WHERE T.Nombre = ' + QuotedStr(cmbTarea2.Text) + ' ' +
              'AND (I.ITS_DTStop >= ' + QuotedStr(deFrom.Text) + ' ' +
              'AND I.ITS_DTStop <= ' + QuotedStr(deTo.Text + ' 23:59:59.999' ) + ') ' +
              'AND I.ITS_DTStart < '+ QuotedStr(deFrom.Text) +
              SQLWhere;

    AddToGridAuto(SQLStr, GridView3);

    lblEntraron.Caption := IntToStr(GridView1.RowCount);
    lblSalieron.Caption := IntToStr(GridView2.RowCount);
    lblHabia.Caption := IntToStr(GridView3.RowCount);

    lblEntraron2.Caption := IntToStr(getCantidad(1,GridView1));
    lblSalieron2.Caption := IntToStr(getCantidad(1,GridView2));
    lblHabia2.Caption := IntToStr(getCantidad(1,GridView3));

    lblEntraron3.Caption := 'Entradas - Ordenes: ' + lblEntraron.Caption + ', Piezas: ' +  lblEntraron2.Caption;
    lblSalieron3.Caption := 'Salidas - Ordenes: ' + lblSalieron.Caption + ', Piezas: ' +  lblSalieron2.Caption;
    lblHabia3.Caption := 'Habia - Ordenes: ' + lblHabia.Caption + ', Piezas: ' +  lblHabia2.Caption;


    Chart1.Series[0].Clear;
    Chart1.Series[0].Add(StrToFloat(lblEntraron.Caption),'Entradas',clBlue);
    Chart1.Series[0].Add(StrToFloat(lblSalieron.Caption),'Salidas',clLtGray);
    Chart1.Series[0].Add(StrToFloat(lblHabia.Caption),'Habia',clRed);
    Application.ProcessMessages;

    Chart2.Series[0].Clear;
    Chart2.Series[0].Add(StrToFloat(lblEntraron2.Caption),'Entradas',clBlue);
    Chart2.Series[0].Add(StrToFloat(lblSalieron2.Caption),'Salidas',clLtGray);
    Chart2.Series[0].Add(StrToFloat(lblHabia2.Caption),'Habia',clRed);
    Application.ProcessMessages;

    lblCargando.Visible := False;
    btnBuscar.Enabled := True;
    btnDetalle.Enabled := True;
    btnImprimir.Enabled := True;

end;

procedure TfrmProductividad.BindGridAuto(SQLStr: String;  Grid:TGridView);
var Qry2 : TADOQuery;
slColumns: TStringList;
i:integer;
begin
    slColumns := TStringList.Create;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    Grid.ClearRows;
    Grid.Columns.Clear;
    for i := 0 to Qry2.Fields.Count - 1 do begin
       slColumns.Add(Qry2.Fields[i].FieldName);
       Grid.Columns.Add(TTextualColumn);
       Grid.Columns[i].Header.Caption := Qry2.Fields[i].FieldName;
       Grid.Columns[i].Width := 100;
       if Qry2.Fields[i].DataType = ftInteger then
         Grid.Columns[i].SortType := stNumeric
       else if Qry2.Fields[i].DataType = ftString then
         Grid.Columns[i].SortType := stAlphabetic
       else if Qry2.Fields[i].DataType = ftDateTime then
         Grid.Columns[i].SortType := stDate;

    end;


    While not Qry2.Eof do
    Begin
        Grid.AddRow(1);
        for i := 0 to (slColumns.Count - 1) do
        begin
                Grid.Cells[i,Grid.RowCount -1] := VarToStr(Qry2[slColumns[i]]);
        end;
       Qry2.Next;
    End;
end;

procedure TfrmProductividad.AddToGridAuto(SQLStr: String;  Grid:TGridView);
var Qry2 : TADOQuery;
slColumns: TStringList;
i:integer;
begin
    slColumns := TStringList.Create;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    for i := 0 to Qry2.Fields.Count - 1 do begin
       slColumns.Add(Qry2.Fields[i].FieldName);
    end;

    While not Qry2.Eof do
    Begin
        Grid.AddRow(1);
        for i := 0 to (slColumns.Count - 1) do
        begin
                Grid.Cells[i,Grid.RowCount -1] := VarToStr(Qry2[slColumns[i]]);
        end;
       Qry2.Next;
    End;
end;


procedure TfrmProductividad.btnDetalleClick(Sender: TObject);
begin
if GroupBox2.Visible = True then
begin
        GroupBox2.Visible := False;
        gbDetalle.Visible := True;
        btnDetalle.Caption := 'Grafica';
end
else
begin
        gbDetalle.Visible := False;
        GroupBox2.Visible := True;
        btnDetalle.Caption := 'Detalle';
end;
end;

function  TfrmProductividad.getCantidad(column:Integer; Grid:TGridView):Integer;
var row, total: Integer;
begin
    total := 0;
    for row := 0 to Grid.RowCount - 1 do
       total := total + StrToInt(Grid.Cells[column, row]);

    result := total;
end;

procedure TfrmProductividad.Exportar1Click(Sender: TObject);
var sFileName: String;
Grid :TGridView;
begin
  Grid := ( ( (Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent as TGridView);


  if Grid.RowCount  = 0 then
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

    ExportGrid(Grid,sFileName);
  end;


end;

procedure TfrmProductividad.CopiarOrden1Click(Sender: TObject);
var Grid :TGridView;
begin
  Grid := ( ( (Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent as TGridView);

  Clipboard.AsText := Grid.Cells[0,Grid.SelectedRow]

end;

procedure TfrmProductividad.Separadoporcomas1Click(Sender: TObject);
var i : integer;
sText : String;
Grid :TGridView;
begin
  Grid := ( ( (Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent as TGridView);
  sText := '';
  for i:= 0 to Grid.RowCount - 1 do
         sText := sText + Grid.Cells[0,i] + ',';

  Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
end;

procedure TfrmProductividad.Encomillas1Click(Sender: TObject);
var i : integer;
sText : String;
Grid :TGridView;
begin
  Grid := ( ( (Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent as TGridView);
  sText := '';
  for i:= 0 to Grid.RowCount - 1 do
         sText := sText + QuotedStr(Grid.Cells[0,i]) + ',';

  Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
end;

procedure TfrmProductividad.deFromChange(Sender: TObject);
begin
    btnBuscarClick(nil);
end;

procedure TfrmProductividad.cmbTarea1Change(Sender: TObject);
begin
    btnBuscarClick(nil);
end;

end.
