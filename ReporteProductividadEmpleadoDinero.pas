unit ReporteProductividadEmpleadoDinero;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,
  Menus, Larco_functions, TeEngine, Series, TeeProcs, Chart, Clipbrd;


type
  TfrmProdEmpleadoDinero = class(TForm)
    gbDetalle: TGroupBox;
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    lblAnio: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    lblCargando: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    btnDetalle: TButton;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    btnImprimir: TButton;
    btnBuscar: TButton;
    txtCliente: TEdit;
    txtProducto: TEdit;
    btnDesc: TButton;
    btnClientes: TButton;
    txtTareas: TEdit;
    Button7: TButton;
    txtEmpleados: TEdit;
    Button1: TButton;
    gbClientes: TGroupBox;
    gvClientes: TGridView;
    chkClientes: TCheckBox;
    btnClientesOK: TButton;
    btnTodosClientes: TButton;
    gbDesc: TGroupBox;
    gvDescs: TGridView;
    chkDesc: TCheckBox;
    btnDescOK: TButton;
    btnTodosDesc: TButton;
    gbTareas: TGroupBox;
    gvTareas: TGridView;
    chkTareas: TCheckBox;
    btnTareasOK: TButton;
    btnTodosTareas: TButton;
    gbEmpleados: TGroupBox;
    gvEmpleados: TGridView;
    chkEmpleados: TCheckBox;
    btnEmpleadosOK: TButton;
    btnTodosEmpleados: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindEmpleados();
    procedure BindClientes();
    procedure BindTareas();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnDescClick(Sender: TObject);
    procedure btnClientesClick(Sender: TObject);
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
      Grid:TGridView; Button: TButton);
    procedure Button7Click(Sender: TObject);
    procedure btnClientesOKClick(Sender: TObject);
    procedure chkClientesClick(Sender: TObject);
    procedure btnTodosClientesClick(Sender: TObject);
    procedure btnDescOKClick(Sender: TObject);
    procedure chkDescClick(Sender: TObject);
    procedure btnTodosDescClick(Sender: TObject);
    procedure btnTareasOKClick(Sender: TObject);
    procedure chkTareasClick(Sender: TObject);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure btnTodosTareasClick(Sender: TObject);
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure Exportar1Click(Sender: TObject);
    function getWhere():String;
    procedure btnBuscarClick(Sender: TObject);
    procedure BindGridAuto(SQLStr: String;  Grid:TGridView);
    procedure Button1Click(Sender: TObject);
    procedure btnEmpleadosOKClick(Sender: TObject);
    procedure chkEmpleadosClick(Sender: TObject);
    procedure btnTodosEmpleadosClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProdEmpleadoDinero: TfrmProdEmpleadoDinero;
  Conn : TADOConnection;
  Qry : TADOQuery;
  
implementation

uses Main;

{$R *.dfm}

procedure TfrmProdEmpleadoDinero.FormCreate(Sender: TObject);
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

    //deFrom.OnChange := deFromChange;
    //deTo.OnChange := deFromChange;

    BindClientes();
    chkClientes.Checked := False;
    btnClientesOKClick(nil);
    BindTareas();
    BindEmpleados();
    //btnBuscarClick(nil);
end;

procedure TfrmProdEmpleadoDinero.BindProductos();
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

procedure TfrmProdEmpleadoDinero.BindEmpleados();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT DISTINCT Nombre FROM tblEmpleados ORDER BY Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvEmpleados.ClearRows;
    While not Qry2.Eof do
    Begin
        gvEmpleados.AddRow(1);
        gvEmpleados.Cells[0,gvEmpleados.RowCount -1] := VarToStr(Qry2['Nombre']);
        Qry2.Next;
    End;

    Qry2.Close;
end;


procedure TfrmProdEmpleadoDinero.BindClientes();
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

procedure TfrmProdEmpleadoDinero.BindTareas();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvTareas.ClearRows;
    While not Qry2.Eof do
    Begin
        gvTareas.AddRow(1);
        gvTareas.Cells[0,gvTareas.RowCount -1] := VarToStr(Qry2['Nombre']);
        Qry2.Next;
    End;

    Qry2.Close;
end;


procedure TfrmProdEmpleadoDinero.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
        Action := caFree;
end;

procedure TfrmProdEmpleadoDinero.btnDescClick(Sender: TObject);
begin
  gbClientes.Visible := False;
  gbTareas.Visible := False;
  gbEmpleados.Visible := False;  
  ShowSeleccionGrid(gbDesc, chkDesc, txtProducto, gvDescs, btnTodosDesc);
  if gbDesc.Visible = True then
      BindProductos();

end;

procedure TfrmProdEmpleadoDinero.btnClientesClick(Sender: TObject);
begin
  gbDesc.Visible := False;
  gbTareas.Visible := False;
  gbEmpleados.Visible := False;  
  ShowSeleccionGrid(gbClientes, chkClientes, txtCliente, gvClientes, btnTodosClientes);

end;

procedure TfrmProdEmpleadoDinero.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
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


procedure TfrmProdEmpleadoDinero.Button7Click(Sender: TObject);
begin
  gbDesc.Visible := False;
  gbClientes.Visible := False;
  gbEmpleados.Visible := False;  
  ShowSeleccionGrid(gbTareas, chkTareas, txtTareas, gvTareas, btnTodosTareas);
end;

procedure TfrmProdEmpleadoDinero.btnClientesOKClick(Sender: TObject);
begin
  ParseSelection(gbClientes,chkClientes,txtCliente,gvClientes);
end;

procedure TfrmProdEmpleadoDinero.chkClientesClick(Sender: TObject);
begin
  gvClientes.Enabled := not chkClientes.Checked;
  btnTodosClientes.Enabled := not chkClientes.Checked;
end;

procedure TfrmProdEmpleadoDinero.btnTodosClientesClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosClientes, gvClientes);
end;

procedure TfrmProdEmpleadoDinero.btnDescOKClick(Sender: TObject);
begin
  ParseSelection(gbDesc,chkDesc,txtProducto,gvDescs);
end;

procedure TfrmProdEmpleadoDinero.chkDescClick(Sender: TObject);
begin
  gvDescs.Enabled := not chkDesc.Checked;
  btnTodosDesc.Enabled := not chkDesc.Checked;
end;

procedure TfrmProdEmpleadoDinero.btnTodosDescClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosDesc, gvDescs);
end;

procedure TfrmProdEmpleadoDinero.btnTareasOKClick(Sender: TObject);
begin
  ParseSelection(gbTareas,chkTareas,txtTareas,gvTareas);
end;

procedure TfrmProdEmpleadoDinero.chkTareasClick(Sender: TObject);
begin
  gvTareas.Enabled := not chkTareas.Checked;
  btnTodosTareas.Enabled := not chkTareas.Checked;
end;

procedure TfrmProdEmpleadoDinero.btnTodosTareasClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosTareas, gvTareas);
end;

procedure TfrmProdEmpleadoDinero.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid: TGridView);
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

procedure TfrmProdEmpleadoDinero.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
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

procedure TfrmProdEmpleadoDinero.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmProdEmpleadoDinero.Exportar1Click(Sender: TObject);
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

function TfrmProdEmpleadoDinero.getWhere():String;
var SQLWhere: String;
begin

    SQLWhere := ' (I.ITS_DTStop >= ' + QuotedStr(deFrom.Text) + ' AND I.ITS_DTStop <= ' + QuotedStr(deTo.Text + ' 23:59:59.999') + ') ';

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

    if txtTareas.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' T.Nombre IN (''' +
        StringReplace(txtTareas.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtEmpleados.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' E.Nombre IN (''' +
        StringReplace(txtEmpleados.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;


    result := SQLWhere;
end;


procedure TfrmProdEmpleadoDinero.btnBuscarClick(Sender: TObject);
var SQLStr : String;
i:integer;
begin
    lblCargando.Visible := True;
    btnBuscar.Enabled := False;
    btnDetalle.Enabled := False;
    btnImprimir.Enabled := False;
    Application.ProcessMessages;

    SQLStr := 'Productividad_Empleado_Dinero ' + QuotedStr(getWhere());

    BindGridAuto(SQLStr, GridView1);

    gvEmpleados.ClearRows;
    for i:=0 to GridView1.RowCount -1  do begin
        if '' = GridView1.Cells[0,i] then begin
                GridView1.Cells[0,i] := 'Ninguno';
        end
        else begin
          gvEmpleados.AddRow(1);
          gvEmpleados.Cells[0,gvEmpleados.RowCount -1] := GridView1.Cells[0,i];
        end;
    end;

    lblCargando.Visible := False;
    btnBuscar.Enabled := True;
    btnDetalle.Enabled := True;
    btnImprimir.Enabled := True;
    Application.ProcessMessages;
end;

procedure TfrmProdEmpleadoDinero.BindGridAuto(SQLStr: String;  Grid:TGridView);
var Qry2 : TADOQuery;
slColumns: TStringList;
i:integer;
columna:String;
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
       columna := Qry2.Fields[i].FieldName;
       slColumns.Add(columna);
       if (columna <> 'Empleado') then begin
         columna := RightStr(columna, Length(columna) - 2);
       end;

       Grid.Columns.Add(TTextualColumn);
       Grid.Columns[i].Header.Caption := columna;
       Grid.Columns[i].Width := 80;
       if Qry2.Fields[i].DataType = ftInteger then
         Grid.Columns[i].SortType := stNumeric
       else if Qry2.Fields[i].DataType = ftString then
         Grid.Columns[i].SortType := stAlphabetic
       else if Qry2.Fields[i].DataType = ftDateTime then
         Grid.Columns[i].SortType := stDate;

    end;
    Grid.Columns[0].Width := 160;

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

procedure TfrmProdEmpleadoDinero.Button1Click(Sender: TObject);
begin
  gbDesc.Visible := False;
  gbClientes.Visible := False;
  gbTareas.Visible := False;
  if gbEmpleados.Visible = True then
  begin
          gbEmpleados.Visible := False;
  end
  else begin
      gbEmpleados.Top := txtProducto.Parent.Top + txtEmpleados.Top + txtEmpleados.Height;
      gbEmpleados.Left := txtEmpleados.Left + 8;
      gvEmpleados.Width := gbEmpleados.Width - 12;
      btnTodosEmpleados.Width := gbEmpleados.Width - 12;

      gbEmpleados.Visible := True;
      chkEmpleados.Checked := False;
      gvEmpleados.Enabled := True;
      btnTodosEmpleados.Enabled := True;
  end;
//  ShowSeleccionGrid(gbEmpleados, chkEmpleados, txtEmpleados, gvEmpleados, btnTodosEmpleados);
end;

procedure TfrmProdEmpleadoDinero.btnEmpleadosOKClick(Sender: TObject);
begin
  ParseSelection(gbEmpleados,chkEmpleados,txtEmpleados,gvEmpleados);
end;

procedure TfrmProdEmpleadoDinero.chkEmpleadosClick(Sender: TObject);
begin
  gvEmpleados.Enabled := not chkEmpleados.Checked;
  btnTodosEmpleados.Enabled := not chkEmpleados.Checked;
end;

procedure TfrmProdEmpleadoDinero.btnTodosEmpleadosClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosEmpleados, gvEmpleados);
end;

end.
