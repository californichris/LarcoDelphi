unit ReporteTotalPiezasStock;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, sndkey32,
  ExtCtrls, StdCtrls, CellEditors, ScrollView, ComCtrls,ComObj,DateUtils,
  CustomGridViewControl, CustomGridView, GridView, Menus,Clipbrd,LTCUtils,Larco_functions;

type
  TfrmReporteTotalPiezasStock = class(TForm)
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    Button1: TButton;
    btnBuscar: TButton;
    Button3: TButton;
    txtDescripcion: TEdit;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    GroupBox5: TGroupBox;
    gvDesc: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    Copiar1: TMenuItem;
    OpenDialog1: TOpenDialog;
    Label4: TLabel;
    txtPlano: TEdit;
    Button2: TButton;
    GroupBox2: TGroupBox;
    gvPlanos: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    Label5: TLabel;
    txtCliente: TEdit;
    Button4: TButton;
    GroupBox3: TGroupBox;
    gvClientes: TGridView;
    CheckBox3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindDescripciones();
    procedure BindPlanos();
    procedure BindClientes();
    procedure BindGrid();
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure txtDescripcionKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure txtPlanoKeyPress(Sender: TObject; var Key: Char);
    procedure txtPlanoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmReporteTotalPiezasStock: TfrmReporteTotalPiezasStock;

implementation

uses Main;

{$R *.dfm}

procedure TfrmReporteTotalPiezasStock.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmReporteTotalPiezasStock.FormCreate(Sender: TObject);
begin
  deFrom.Date := StartOfAMonth(YearOf(Now), MonthOf(Now));
  deTo.Date :=  EndOfTheMonth(Now);
  BindDescripciones();
  BindPlanos();
  BindClientes();

  BindGrid();
end;

procedure TfrmReporteTotalPiezasStock.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
entradas, salidas, existencia : Integer;
begin
    entradas := 0;
    salidas := 0;
    existencia := 0;
    Conn := nil;
    Qry := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT P.PN_Descripcion, P.PN_Numero, ' +
                'CASE WHEN O.Numero IS NULL THEN '''' ELSE O.Numero END AS Numero, ' +
                'SUBSTRING(S.ITE_Nombre,4,3) AS Cliente, ' +
                'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) AS Entradas, ' +
                'SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Salidas, ' +
                'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) - SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Cantidad ' +
                'FROM tblPlano P ' +
                'INNER JOIN tblStock S ON P.PN_Id = S.PN_Id ' +
                'LEFT OUTER JOIN tblOrdenes O ON S.ITE_Nombre = O.ITE_Nombre ' + 
                'WHERE S.ST_Fecha >= ' + QuotedStr(deFrom.Text) + ' AND S.ST_Fecha <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99');

      if Pos('*', txtPlano.Text) <> 0 then begin
        SQLStr := SQLStr + ' AND P.PN_Numero LIKE ''' + StringReplace(txtPlano.Text, '*', '%', [rfReplaceAll, rfIgnoreCase]) + '''';
      end else if txtPlano.Text <> 'Todos' then begin
        SQLStr := SQLStr + ' AND P.PN_Numero IN (''' + StringReplace(txtPlano.Text, ',', ''',''', [rfReplaceAll, rfIgnoreCase]) + ''')';
      end;

      if Pos('*', txtDescripcion.Text) <> 0 then begin
        SQLStr := SQLStr + ' AND P.PN_Descripcion LIKE ''' + StringReplace(txtDescripcion.Text, '*', '%', [rfReplaceAll, rfIgnoreCase]) + '''';
      end else if txtDescripcion.Text <> 'Todos' then begin
        SQLStr := SQLStr + ' AND P.PN_Descripcion IN (''' + StringReplace(txtDescripcion.Text, ',', ''',''', [rfReplaceAll, rfIgnoreCase]) + ''')';
      end;

      if txtCliente.Text <> 'Todos' then begin
        SQLStr := SQLStr + ' AND SUBSTRING(S.ITE_Nombre,4,3) IN (''' + StringReplace(txtCliente.Text, ',', ''',''', [rfReplaceAll, rfIgnoreCase]) + ''')';
      end;

      SQLStr := SQLStr + ' GROUP BY P.PN_Descripcion, P.PN_Numero, SUBSTRING(S.ITE_Nombre,4,3), O.Numero';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      GridView1.ClearRows();
      while not Qry.Eof do begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['PN_Descripcion']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['PN_Numero']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Cliente']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Entradas']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['Salidas']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);

        entradas := entradas + StrToInt(VarToStr(Qry['Entradas']));
        salidas := salidas + StrToInt(VarToStr(Qry['Salidas']));
        existencia := existencia + StrToInt(VarToStr(Qry['Cantidad']));

        Qry.Next;
      end;

        GridView1.AddRow(1);
        GridView1.Cells[3,GridView1.RowCount -1] := 'Totales :';
        GridView1.Cells[4,GridView1.RowCount -1] := IntToStr(entradas);
        GridView1.Cells[5,GridView1.RowCount -1] := IntToStr(salidas);
        GridView1.Cells[6,GridView1.RowCount -1] := IntToStr(existencia);
    end
    finally
      CloseConns(Qry, Conn);
    end;
end;

procedure TfrmReporteTotalPiezasStock.BindDescripciones();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT DISTINCT PN_Descripcion FROM tblPlano ORDER BY PN_Descripcion';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvDesc.ClearRows;
      while not Qry.Eof do begin
        gvDesc.AddRow(1);
        gvDesc.Cells[0,gvDesc.RowCount -1] := VarToStr(Qry['PN_Descripcion']);
        Qry.Next;
      end;
    end
    finally
      CloseConns(Qry, Conn);
    end;
end;

procedure TfrmReporteTotalPiezasStock.BindPlanos();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT DISTINCT PN_Numero FROM tblPlano ORDER BY PN_Numero';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvPlanos.ClearRows;
      while not Qry.Eof do begin
        gvPlanos.AddRow(1);
        gvPlanos.Cells[0,gvPlanos.RowCount -1] := VarToStr(Qry['PN_Numero']);
        Qry.Next;
      end;
    end
    finally
      CloseConns(Qry, Conn);
    end;
end;

procedure TfrmReporteTotalPiezasStock.BindClientes();
var Qry : TADOQuery;
SQLStr : String;
//slClientes : TStringList;
begin
    //slClientes := TStringList.Create;
    //slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';
    Conn := nil;
    Qry := nil;
    try
    begin
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
          //if (slClientes.IndexOf(VarToStr(Qry['Clave'])) = -1) then begin
          //        gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
          //end;
          Qry.Next;
      End;
    end
    finally
      CloseConns(Qry, Conn);
    end;
end;

procedure TfrmReporteTotalPiezasStock.CheckBox1Click(Sender: TObject);
begin
gvDesc.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmReporteTotalPiezasStock.CheckBox2Click(Sender: TObject);
begin
gvPlanos.Enabled := not CheckBox2.Checked;
btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmReporteTotalPiezasStock.btnOKClick(Sender: TObject);
var i: integer;
sDescs : String;
begin
  GroupBox5.Visible := False;
  if CheckBox1.Checked = True then begin
    txtDescripcion.Text := 'Todos';
  end
  else begin
        sDescs := '';
        for i:= 0 to gvDesc.RowCount - 1 do
        begin
                if gvDesc.Cell[1,i].AsBoolean = True then
                begin
                        sDescs := sDescs + gvDesc.Cells[0,i] + ',';
                end;
        end;

        txtDescripcion.Text := 'Todos';
        if sDescs <> '' then
        begin
          txtDescripcion.Text :=  LeftStr(sDescs, Length(sDescs) - 1);
        end;
  end;
end;

procedure TfrmReporteTotalPiezasStock.btnOK2Click(Sender: TObject);
var i: integer;
sDescs : String;
begin
  GroupBox2.Visible := False;
  if CheckBox2.Checked = True then begin
    txtPlano.Text := 'Todos';
  end
  else begin
        sDescs := '';
        for i:= 0 to gvPlanos.RowCount - 1 do
        begin
                if gvPlanos.Cell[1,i].AsBoolean = True then
                begin
                        sDescs := sDescs + gvPlanos.Cells[0,i] + ',';
                end;
        end;

        txtPlano.Text := 'Todos';
        if sDescs <> '' then
        begin
          txtPlano.Text :=  LeftStr(sDescs, Length(sDescs) - 1);
        end;
  end;
end;

procedure TfrmReporteTotalPiezasStock.btnTodosClick(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos.Caption) = UT('Seleccionar Todos') then begin
        btnTodos.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvDesc.RowCount - 1 do
        begin
                gvDesc.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos.Caption := 'Seleccionar Todos';
        for i:= 0 to gvDesc.RowCount - 1 do
        begin
                gvDesc.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmReporteTotalPiezasStock.btnTodos2Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos2.Caption) = UT('Seleccionar Todos') then begin
        btnTodos2.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvPlanos.RowCount - 1 do
        begin
                gvPlanos.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos2.Caption := 'Seleccionar Todos';
        for i:= 0 to gvPlanos.RowCount - 1 do
        begin
                gvPlanos.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmReporteTotalPiezasStock.Button3Click(Sender: TObject);
begin
  if GroupBox5.Visible = True then
  begin
          GroupBox5.Visible := False;
  end
  else begin
      GroupBox5.Visible := True;
      if txtDescripcion.Text = 'Todos' then
      begin
              CheckBox1.Checked := True;
              gvDesc.Enabled := False;
              btnTodos.Enabled := False;
      end
      else
      begin
              CheckBox1.Checked := False;
              gvDesc.Enabled := True;
              btnTodos.Enabled := True;
      end;

      GroupBox5.Top := txtDescripcion.Top + txtDescripcion.Height + 5;
      GroupBox5.Left := txtDescripcion.Left + 5;
  end;

end;

procedure TfrmReporteTotalPiezasStock.Button2Click(Sender: TObject);
begin
  if GroupBox2.Visible = True then
  begin
          GroupBox2.Visible := False;
  end
  else begin
      GroupBox2.Visible := True;
      if txtPlano.Text = 'Todos' then
      begin
              CheckBox2.Checked := True;
              gvPlanos.Enabled := False;
              btnTodos2.Enabled := False;
      end
      else
      begin
              CheckBox2.Checked := False;
              gvPlanos.Enabled := True;
              btnTodos2.Enabled := True;
      end;

      GroupBox2.Top := txtPlano.Top + txtPlano.Height + 5;
      GroupBox2.Left := txtPlano.Left + 5;
  end;

end;

procedure TfrmReporteTotalPiezasStock.CheckBox3Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox3.Checked;
btnTodos3.Enabled := not CheckBox3.Checked;
end;

procedure TfrmReporteTotalPiezasStock.btnOK3Click(Sender: TObject);
var i: integer;
sClientes : String;
begin
  GroupBox3.Visible := False;
  if CheckBox3.Checked = True then begin
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

procedure TfrmReporteTotalPiezasStock.btnTodos3Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos3.Caption) = UT('Seleccionar Todos') then begin
        btnTodos3.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos3.Caption := 'Seleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmReporteTotalPiezasStock.Button4Click(Sender: TObject);
begin
  if GroupBox3.Visible = True then
  begin
          GroupBox3.Visible := False;
  end
  else begin
      GroupBox3.Visible := True;
      if txtCliente.Text = 'Todos' then
      begin
              CheckBox3.Checked := True;
              gvClientes.Enabled := False;
              btnTodos3.Enabled := False;
      end
      else
      begin
              CheckBox3.Checked := False;
              gvClientes.Enabled := True;
              btnTodos3.Enabled := True;
      end;

      GroupBox3.Top := txtCliente.Top + txtCliente.Height + 5;
      GroupBox3.Left := txtCliente.Left + 10;
  end;
end;

procedure TfrmReporteTotalPiezasStock.btnBuscarClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmReporteTotalPiezasStock.Button1Click(Sender: TObject);
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

procedure TfrmReporteTotalPiezasStock.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
  xlCSV = 6;
var XApp : Variant;
Sheet : Variant;
Row, col, startRow :Integer;
begin
      try //Create the excel object
      begin
            XApp:= CreateOleObject('Excel.Application');
            //XApp.Visible := True;
            XApp.Visible := False;
            XApp.DisplayAlerts := False;
      end;
      except
        ShowMessage('No se pudo abrir Microsoft Excel,  parece que no esta instalado en el sistema.');
        Exit;
      end;

      XApp.Workbooks.Add(xlWorkSheet);
      Sheet := XApp.Workbooks[1].WorkSheets[1];
      Sheet.Name := 'Total de piezas en Stock';

      Sheet.Cells[1,1] := 'Total de piezas en Stock';
      Sheet.Cells[2,1] := 'Periodo de: ' + deFrom.Text + ' hasta: ' + deTo.Text;
      XApp.Range['A1:A2'].Font.Bold := True;

      startRow := 4;

      for Col := 1 to Grid.Columns.Count do
              Sheet.Cells[startRow, Col] := Grid.Columns[Col - 1].Header.Caption;

      for Row := 1 to Grid.RowCount do
                for Col := 1 to Grid.Columns.Count do
                        Sheet.Cells[Row + startRow,Col] := Grid.Cells[Col - 1,Row - 1];


      Sheet.Cells.Select;
      Sheet.Cells.EntireColumn.AutoFit;                        

      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

      ShowMessage('El archivo se exporto exitosamente.');
end;

procedure TfrmReporteTotalPiezasStock.Copiar1Click(Sender: TObject);
begin
Button1Click(nil);
end;

procedure TfrmReporteTotalPiezasStock.txtDescripcionKeyDown(
  Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      BindGrid();
  end
end;

procedure TfrmReporteTotalPiezasStock.txtDescripcionKeyPress(
  Sender: TObject; var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmReporteTotalPiezasStock.txtPlanoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmReporteTotalPiezasStock.txtPlanoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      BindGrid();
  end
end;

end.
