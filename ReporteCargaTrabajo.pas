unit ReporteCargaTrabajo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses;

type
  TfrmCargaTrabajo = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
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
    Button2: TButton;
    btnBuscar: TButton;
    Button3: TButton;
    txtCliente: TEdit;
    txtTareas: TEdit;
    Button7: TButton;
    txtProducto: TEdit;
    Button4: TButton;
    GroupBox3: TGroupBox;
    gvTareas: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    GroupBox2: TGroupBox;
    gvProds: TGridView;
    CheckBox3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    GridView1: TGridView;
    GroupBox4: TGroupBox;
    Label3: TLabel;
    cmbRenglon: TComboBox;
    Label5: TLabel;
    cmbColumna: TComboBox;
    SaveDialog1: TSaveDialog;
    chkPiezas: TCheckBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindClientes();
    procedure BindTareas();
    procedure BindGrid();
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure chkRecibidoClick(Sender: TObject);
    procedure chkInternaClick(Sender: TObject);
    procedure chkEntregaClick(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure GridView1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    function getIntValue(value: boolean):Integer;
    procedure chkPiezasClick(Sender: TObject);
    procedure GridView1SortColumn(Sender: TObject; ACol: Integer;
      Ascending: Boolean);
    procedure GridView1DblClick(Sender: TObject);
    function getWhere():String;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCargaTrabajo: TfrmCargaTrabajo;

implementation

uses Main, ReporteCargaTrabajoDetalle;

{$R *.dfm}

procedure TfrmCargaTrabajo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCargaTrabajo.FormCreate(Sender: TObject);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT CASE WHEN Min(Interna) IS NULL THEN GETDATE() ELSE Min(Interna) END As Interna ' +
                'FROM tblOrdenes O ' +
                'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      deInterna1.Date := Now;
      if Qry.RecordCount > 0 then
              deInterna1.Date := StrToDateTime( VarToStr(Qry['Interna']) ) ;

    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    deInterna2.Date := DateAdd(Now,5,daDays);

    deRecibido1.Date := Now;
    deRecibido2.Date := Now;

    deEntrega1.Date := Now;
    deEntrega2.Date := Now;

    cmbRenglon.ItemIndex := 1;
    cmbColumna.ItemIndex := 0;

    BindProductos();
    BindClientes();
    CheckBox1.Checked := False;
    btnOKClick(nil);

    BindTareas();
    BindGrid();
end;

procedure TfrmCargaTrabajo.BindProductos();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Nombre FROM tblProductos Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvProds.ClearRows;
      While not Qry.Eof do
      Begin
          gvProds.AddRow(1);
          gvProds.Cells[0,gvProds.RowCount -1] := VarToStr(Qry['Nombre']);
          Qry.Next;
      End;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure TfrmCargaTrabajo.BindClientes();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '060,062,699,799,899,999,960';

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
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
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure TfrmCargaTrabajo.BindTareas();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvTareas.ClearRows;
      While not Qry.Eof do
      Begin
          gvTareas.AddRow(1);
          gvTareas.Cells[0,gvTareas.RowCount -1] := VarToStr(Qry['Nombre']);
          Qry.Next;
      End;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure TfrmCargaTrabajo.Button4Click(Sender: TObject);
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

procedure TfrmCargaTrabajo.Button3Click(Sender: TObject);
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

procedure TfrmCargaTrabajo.Button7Click(Sender: TObject);
begin
  if GroupBox3.Visible = True then
  begin
          GroupBox3.Visible := False;
  end
  else begin
      GroupBox3.Visible := True;
      if txtTareas.Text = 'Todos' then
      begin
              CheckBox2.Checked := True;
              gvTareas.Enabled := False;
              btnTodos2.Enabled := False;
      end
      else
      begin
              CheckBox2.Checked := False;
              gvTareas.Enabled := True;
              btnTodos2.Enabled := True;
      end;

      GroupBox3.Top := txtTareas.Top + txtTareas.Height + 5;
      GroupBox3.Left := txtTareas.Left + 10;
  end;

end;

procedure TfrmCargaTrabajo.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmCargaTrabajo.CheckBox2Click(Sender: TObject);
begin
gvTareas.Enabled := not CheckBox2.Checked;
btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmCargaTrabajo.CheckBox3Click(Sender: TObject);
begin
gvProds.Enabled := not CheckBox3.Checked;
btnTodos3.Enabled := not CheckBox3.Checked;
end;

procedure TfrmCargaTrabajo.btnOKClick(Sender: TObject);
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

procedure TfrmCargaTrabajo.btnOK2Click(Sender: TObject);
var i: integer;
sTareas : String;
begin
  GroupBox3.Visible := False;
  if CheckBox2.Checked = True then begin
          txtTareas.Text := 'Todos';
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
        txtTareas.Text := 'Todos';
        if sTareas <> '' then
        begin
                txtTareas.Text :=  LeftStr(sTareas,Length(sTareas) - 1);
        end;
  end;

end;

procedure TfrmCargaTrabajo.btnOK3Click(Sender: TObject);
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

procedure TfrmCargaTrabajo.btnTodosClick(Sender: TObject);
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

procedure TfrmCargaTrabajo.btnTodos2Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos2.Caption) = UT('Seleccionar Todos') then begin
        btnTodos2.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvTareas.RowCount - 1 do
        begin
                gvTareas.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos2.Caption := 'Seleccionar Todos';
        for i:= 0 to gvTareas.RowCount - 1 do
        begin
                gvTareas.Cell[1,i].AsBoolean := False;
        end;
  end;

end;

procedure TfrmCargaTrabajo.btnTodos3Click(Sender: TObject);
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


procedure TfrmCargaTrabajo.chkRecibidoClick(Sender: TObject);
begin
deRecibido1.Enabled := chkRecibido.Checked;
deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmCargaTrabajo.chkInternaClick(Sender: TObject);
begin
deInterna1.Enabled := chkInterna.Checked;
deInterna2.Enabled := chkInterna.Checked;
end;

procedure TfrmCargaTrabajo.chkEntregaClick(Sender: TObject);
begin
deEntrega1.Enabled := chkEntrega.Checked;
deEntrega2.Enabled := chkEntrega.Checked;
end;

procedure TfrmCargaTrabajo.BindGrid();
const Fields : array[0..2] of PChar =
('T.Nombre','O.Producto','SUBSTRING(I.ITE_Nombre,4,3)');
var SQLStr : String;
i,iTotalColumn:integer;
slColumns: TStringList;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    if cmbRenglon.Text = cmbColumna.Text then begin
        ShowMessage('El renglon y la columna no puede ser el mismo campo.');
        cmbRenglon.SetFocus;
        Exit;
    end;


    slColumns := TStringList.Create;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'Carga_Trabajo ' + QuotedStr(Fields[cmbRenglon.ItemIndex]) + ',' + QuotedStr(Fields[cmbColumna.ItemIndex]);

      SQLStr := SQLStr + ',' + QuotedStr(getWhere()) + ',' + IntToStr(getIntValue(chkPiezas.Checked));

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      GridView1.ClearRows;
      GridView1.Columns.Clear;
      iTotalColumn := 0;
      for i := 0 to Qry.Fields.Count - 1 do begin
         slColumns.Add(Qry.Fields[i].FieldName);
         GridView1.Columns.Add(TTextualColumn);
         GridView1.Columns[i].Header.Caption := Qry.Fields[i].FieldName;

         if GridView1.Columns[i].Header.Caption = 'Total' then begin
                  iTotalColumn := i;
         end;

         if i > 0 then begin
                 GridView1.Columns[i].SortType := stNumeric;
         end
         else begin
                 GridView1.Columns[i].SortType := stAlphabetic;
         end;
         GridView1.Columns[i].Width := 70;
      end;


      While not Qry.Eof do
      Begin
          GridView1.AddRow(1);
          for i := 0 to (slColumns.Count - 1) do
          begin
                  GridView1.Cells[i,GridView1.RowCount -1] := VarToStr(Qry[slColumns[i]]);
          end;
         Qry.Next;
      End;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    GridView1.Columns[0].Header.Caption := cmbRenglon.Text;
    GridView1.Columns[0].Width := 100;

    for i := 0 to GridView1.RowCount - 1 do begin
        if GridView1.Cells[0,i] = 'Total' then begin
                GridView1.MoveRow(i, GridView1.RowCount - 1);
                Break;
        end;
    end;

   GridView1.Columns.Add(TTextualColumn);
   GridView1.Columns[GridView1.Columns.Count - 1].Header.Caption := 'Total';
   GridView1.Columns[GridView1.Columns.Count - 1].SortType := stNumeric;
   for i := 0 to GridView1.RowCount - 1 do begin
        GridView1.Cells[GridView1.Columns.Count - 1,i] := GridView1.Cells[iTotalColumn,i];
   end;

   GridView1.Columns.Delete(iTotalColumn);


   application.ProcessMessages;

end;


procedure TfrmCargaTrabajo.btnBuscarClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmCargaTrabajo.GridView1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
                BindGrid;
end;

procedure TfrmCargaTrabajo.Button1Click(Sender: TObject);
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

procedure TfrmCargaTrabajo.ExportGrid(Grid: TGridView;sFileName: String);
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

function TfrmCargaTrabajo.getIntValue(value: boolean):Integer;
begin
  if(value = true) then
    result := 1
  else
    result := 0;
end;

procedure TfrmCargaTrabajo.chkPiezasClick(Sender: TObject);
begin
        BindGrid();
end;

procedure TfrmCargaTrabajo.GridView1SortColumn(Sender: TObject;
  ACol: Integer; Ascending: Boolean);
  var i:Integer;
begin
    for i := 0 to GridView1.RowCount - 1 do begin
        if ( (GridView1.Cells[0,i] = 'Total') and (i <> GridView1.RowCount - 1))  then begin
                GridView1.MoveRow(i, GridView1.RowCount - 1);
                Break;
        end;
    end;
end;

procedure TfrmCargaTrabajo.GridView1DblClick(Sender: TObject);
const Fields : array[0..2] of PChar =
('T.Nombre','O.Producto','SUBSTRING(I.ITE_Nombre,4,3)');
var cellValue,SQLStr,ren, col, field : String;
i:Integer;
Conn : TADOConnection;
Qry : TADOQuery;
begin
     if(GridView1.SelectedColumn = 0) then
        Exit;

     cellValue := GridView1.Cells[GridView1.SelectedColumn, GridView1.SelectedRow];
     if(cellValue = '0') then
        Exit;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

       SQLStr := 'SELECT ';

       ren := cmbRenglon.Text;
       col := cmbColumna.Text;
       for i:=0 to cmbRenglon.Items.Count -1 do
       begin
          if ( (cmbRenglon.Items[i] <> ren) and (cmbRenglon.Items[i] <> col) ) then begin
                  field := Fields[i];
                  SQLStr := SQLStr + field + ' AS Nombre, ';
                  Break;
          end;
       end;

       if chkPiezas.Checked then begin
          SQLStr := SQLStr + 'SUM(O.Requerida) AS Cantidad';
       end
       else begin
          SQLStr := SQLStr + 'COUNT(*) AS Cantidad';
       end;

       SQLStr := SQLStr + ' FROM tblItemTasks I ' +
                 'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                 'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nombre ';

       SQLStr := SQLStr + ' WHERE ' + getWhere();

       SQLStr := SQLStr + ' AND ' + Fields[cmbRenglon.ItemIndex] + ' = ' +
                 QuotedStr(GridView1.Cells[0, GridView1.SelectedRow]) +
                 ' AND ' + Fields[cmbColumna.ItemIndex] + ' = ' +
                 QuotedStr(GridView1.Columns[GridView1.SelectedColumn].Header.Caption);

       SQLStr := SQLStr + ' GROUP BY ' + field;

       Qry.SQL.Clear;
       Qry.SQL.Text := SQLStr;
       Qry.Open;

       Application.CreateForm(TfrmCTDetail,frmCTDetail);

       While not Qry.Eof do
       Begin
          frmCTDetail.GridView1.AddRow(1);
          frmCTDetail.GridView1.Cells[0,frmCTDetail.GridView1.RowCount -1] := VarToStr(Qry['Nombre']);
          frmCTDetail.GridView1.Cells[1,frmCTDetail.GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);
          Qry.Next;
       end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

     frmCTDetail.GridView1.Columns[0].Header.Caption := cmbRenglon.Items[i];
     frmCTDetail.GridView1.Columns[1].Header.Caption := 'Ordenes';
     if chkPiezas.Checked then
         frmCTDetail.GridView1.Columns[1].Header.Caption := 'Piezas';

     frmCTDetail.GridView1.AddRow(1);
     frmCTDetail.GridView1.Cells[0,frmCTDetail.GridView1.RowCount -1] := 'Total';
     frmCTDetail.GridView1.Cells[1,frmCTDetail.GridView1.RowCount -1] := cellValue;

     frmCTDetail.ShowModal;

end;

function TfrmCargaTrabajo.getWhere():String;
var SQLWhere: String;
begin

    SQLWhere := ' I.ITS_Status IN (0,1,3) ';
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

    result := SQLWhere;
end;

end.
