unit PendientesFacturar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors, ScrollView, CustomGridViewControl,
  CustomGridView, GridView, Menus,ADODB,DB, All_Functions,chris_Functions,LTCUtils,
  ComCtrls,ComObj,Larco_Functions;

type
  TfrmPendientesFact = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    deRecibido1: TDateEditor;
    deRecibido2: TDateEditor;
    Button1: TButton;
    chkRecibido: TCheckBox;
    deInterna1: TDateEditor;
    chkInterna: TCheckBox;
    deInterna2: TDateEditor;
    deEntrega1: TDateEditor;
    chkEntrega: TCheckBox;
    deEntrega2: TDateEditor;
    Button2: TButton;
    btnBuscar: TButton;
    txtCliente: TEdit;
    txtOrden: TEdit;
    txtProducto: TEdit;
    btnDesc: TButton;
    btnClientes: TButton;
    btnOrdenes: TButton;
    lblCount: TLabel;
    gbOrdenes: TGroupBox;
    gvOrdenes: TGridView;
    chkOrdenes: TCheckBox;
    btnOrdenesOK: TButton;
    btnTodosOrdenes: TButton;
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
    gvPend: TGridView;
    SaveDialog2: TSaveDialog;
    OpenDialog1: TOpenDialog;
    lblAnio: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    procedure BindOrdenes();
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
      Grid:TGridView; Button: TButton);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure btnDescClick(Sender: TObject);
    procedure chkDescClick(Sender: TObject);
    procedure chkClientesClick(Sender: TObject);
    procedure chkOrdenesClick(Sender: TObject);
    procedure btnDescOKClick(Sender: TObject);
    procedure btnClientesOKClick(Sender: TObject);
    procedure btnOrdenesOKClick(Sender: TObject);
    procedure btnTodosDescClick(Sender: TObject);
    procedure btnTodosClientesClick(Sender: TObject);
    procedure btnTodosOrdenesClick(Sender: TObject);
    procedure btnClientesClick(Sender: TObject);
    procedure chkRecibidoClick(Sender: TObject);
    procedure chkInternaClick(Sender: TObject);
    procedure chkEntregaClick(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure gbDescExit(Sender: TObject);
    procedure btnOrdenesClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPendientesFact: TfrmPendientesFact;
  Conn : TADOConnection;
  Qry : TADOQuery;
  SQLWhere : String;

implementation

uses Main, ReportePendientesFacurar;

{$R *.dfm}

procedure TfrmPendientesFact.FormCreate(Sender: TObject);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);

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

    BindGrid();
//    BindProductos();
//    BindClientes();

    Qry2.Close;

end;

procedure TfrmPendientesFact.BindProductos();
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

procedure TfrmPendientesFact.BindClientes();
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

procedure TfrmPendientesFact.BindOrdenes();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT DISTINCT O.OrdenCompra FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.ID ' +
              'INNER JOIN tblOrdenes O ON O.ITE_Nombre = I.ITE_Nombre ' +
              'WHERE T.Nombre = ''VentasFinal'' ' +
              'AND I.ITS_STATUS = 2 ' +
              'AND O.OrdenCompra IS NOT NULL AND RTRIM(O.OrdenCompra) <> '''' ';
               //AND LEFT(O.ITE_Nombre,2) = ' + RightStr(lblAnio.Caption,2) + ' ' +

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr + SQLWhere;
    Qry2.Open;

    gvOrdenes.ClearRows;
    While not Qry2.Eof do
    Begin
        gvOrdenes.AddRow(1);
        gvOrdenes.Cells[0,gvOrdenes.RowCount -1] := VarToStr(Qry2['OrdenCompra']);
        Qry2.Next;
    End;

    Qry2.Close;
end;


procedure TfrmPendientesFact.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmPendientesFact.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
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

procedure TfrmPendientesFact.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
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

procedure TfrmPendientesFact.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit;
                                         Grid: TGridView);
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


procedure TfrmPendientesFact.btnDescClick(Sender: TObject);
begin
  gbClientes.Visible := False;
  gbOrdenes.Visible := False;
  ShowSeleccionGrid(gbDesc, chkDesc, txtProducto, gvDescs, btnTodosDesc);
  if gbDesc.Visible = True then
      BindProductos();
end;

procedure TfrmPendientesFact.chkDescClick(Sender: TObject);
begin
  gvDescs.Enabled := not chkDesc.Checked;
  btnTodosDesc.Enabled := not chkDesc.Checked;
end;

procedure TfrmPendientesFact.chkClientesClick(Sender: TObject);
begin
  gvClientes.Enabled := not chkClientes.Checked;
  btnTodosClientes.Enabled := not chkClientes.Checked;
end;

procedure TfrmPendientesFact.chkOrdenesClick(Sender: TObject);
begin
  gvOrdenes.Enabled := not chkOrdenes.Checked;
  btnTodosOrdenes.Enabled := not chkOrdenes.Checked;
end;

procedure TfrmPendientesFact.btnDescOKClick(Sender: TObject);
begin
  ParseSelection(gbDesc,chkDesc,txtProducto,gvDescs);
end;

procedure TfrmPendientesFact.btnClientesOKClick(Sender: TObject);
begin
  ParseSelection(gbClientes,chkClientes,txtCliente,gvClientes);
end;

procedure TfrmPendientesFact.btnOrdenesOKClick(Sender: TObject);
begin
  ParseSelection(gbOrdenes,chkOrdenes,txtOrden,gvOrdenes);
end;

procedure TfrmPendientesFact.btnTodosDescClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosDesc, gvDescs);
end;

procedure TfrmPendientesFact.btnTodosClientesClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosClientes, gvClientes);
end;

procedure TfrmPendientesFact.btnTodosOrdenesClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosOrdenes, gvOrdenes);
end;

procedure TfrmPendientesFact.chkRecibidoClick(Sender: TObject);
begin
  deRecibido1.Enabled := chkRecibido.Checked;
  deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmPendientesFact.chkInternaClick(Sender: TObject);
begin
  deInterna1.Enabled := chkInterna.Checked;
  deInterna2.Enabled := chkInterna.Checked;
end;

procedure TfrmPendientesFact.chkEntregaClick(Sender: TObject);
begin
  deEntrega1.Enabled := chkEntrega.Checked;
  deEntrega2.Enabled := chkEntrega.Checked;
end;

procedure TfrmPendientesFact.btnBuscarClick(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmPendientesFact.BindGrid();
var SQLStr,SQLWhere2 : String;
begin
    SQLStr := 'SELECT RIGHT(O.ITE_Nombre,LEN(O.ITE_Nombre) - 3) AS Orden,* FROM tblItemTasks I ' +
              'INNER JOIN tblTareas T ON I.TAS_ID = T.ID ' +
              'INNER JOIN tblOrdenes O ON O.ITE_Nombre = I.ITE_Nombre ' +
              'WHERE T.Nombre = ''VentasFinal'' ' +
              'AND I.ITS_STATUS = 2 ' +
              'AND LEFT(O.ITE_Nombre,2) = ' + RightStr(lblAnio.Caption,2) + ' ';

    SQLWhere := '';
    SQLWhere2 := '';
    if chkRecibido.Checked then
    begin
        SQLWhere := SQLWhere + ' AND (O.Recibido >= ' + QuotedStr(deRecibido1.Text) +
                    ' and O.Recibido <= ' + QuotedStr(deRecibido2.Text) + ') ';
    end;

    if chkInterna.Checked then
    begin
        SQLWhere := SQLWhere + ' AND (O.Interna >= ' + QuotedStr(deInterna1.Text) +
                    ' and O.Interna <= ' + QuotedStr(deInterna2.Text) + ') ';
    end;

    if chkEntrega.Checked then
    begin
        SQLWhere := SQLWhere + ' AND (O.Entrega >= ' + QuotedStr(deEntrega1.Text) +
                    ' and O.Entrega <= ' + QuotedStr(deEntrega2.Text) + ') ';
    end;

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

    if txtOrden.Text <> 'Todos' then
    begin
        SQLWhere2 := SQLWhere2 + ' AND O.OrdenCompra IN (''' +
        StringReplace(txtOrden.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr + SQLWhere + SQLWhere2 + ' ORDER BY I.ITE_NOMBRE' ;
    Qry.Open;

    gvPend.ClearRows;
    While not Qry.Eof do
    Begin
        gvPend.AddRow(1);
        gvPend.Cells[0,gvPend.RowCount -1] := VarToStr(Qry['Orden']);
        gvPend.Cells[1,gvPend.RowCount -1] := VarToStr(Qry['Requerida']);
        gvPend.Cells[2,gvPend.RowCount -1] := VarToStr(Qry['Producto']);
        gvPend.Cells[3,gvPend.RowCount -1] := VarToStr(Qry['Numero']);
        gvPend.Cells[4,gvPend.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        gvPend.Cells[5,gvPend.RowCount -1] := VarToStr(Qry['Recibido']);
        gvPend.Cells[6,gvPend.RowCount -1] := VarToStr(Qry['Interna']);
        gvPend.Cells[7,gvPend.RowCount -1] := VarToStr(Qry['Entrega']);
        gvPend.Cells[8,gvPend.RowCount -1] := VarToStr(Qry['Unitario']);
        Qry.Next;
    End;

    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(gvPend.RowCount);
end;

procedure TfrmPendientesFact.gbDescExit(Sender: TObject);
begin
(Sender As TGroupBox).Visible := False;
end;

procedure TfrmPendientesFact.btnClientesClick(Sender: TObject);
begin
  gbDesc.Visible := False;
  gbOrdenes.Visible := False;
  ShowSeleccionGrid(gbClientes, chkClientes, txtCliente, gvClientes, btnTodosClientes);
  if gbClientes.Visible = True then
      BindClientes();
end;

procedure TfrmPendientesFact.btnOrdenesClick(Sender: TObject);
begin
  gbDesc.Visible := False;
  gbClientes.Visible := False;
  ShowSeleccionGrid(gbOrdenes, chkOrdenes, txtOrden, gvOrdenes, btnTodosOrdenes);
  if gbOrdenes.Visible = True then
      BindOrdenes();

end;

procedure TfrmPendientesFact.Button1Click(Sender: TObject);
var sFileName: String;
begin
if gvPend.RowCount = 0 then
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

    ExportGrid(gvPend,sFileName);

  end;


end;

procedure TfrmPendientesFact.ExportGrid(Grid: TGridView;sFileName: String);
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


procedure TfrmPendientesFact.Button2Click(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrPendientesFacturar, qrPendientesFacturar);
    qrPendientesFacturar.ReportTitle.Caption := 'Pendientes de Facurar. ';

    qrPendientesFacturar.QRSubDetail1.DataSet := Qry;
    qrPendientesFacturar.Field1.DataSet := Qry;
    qrPendientesFacturar.Field1.DataField := 'Orden';

    qrPendientesFacturar.Field2.DataSet := Qry;
    qrPendientesFacturar.Field2.DataField := 'Requerida';

    qrPendientesFacturar.Field3.DataSet := Qry;
    qrPendientesFacturar.Field3.DataField := 'Producto';

    qrPendientesFacturar.Field4.DataSet := Qry;
    qrPendientesFacturar.Field4.DataField := 'Numero';

    qrPendientesFacturar.Field5.DataSet := Qry;
    qrPendientesFacturar.Field5.DataField := 'OrdenCompra';

    qrPendientesFacturar.Preview;
    qrPendientesFacturar.Free;

end;

end.
