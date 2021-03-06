unit ReporteCumplimientoFechaEntrega;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,Clipbrd,
  Menus;

type
  TfrmCumplimientoTiempoEntrega = class(TForm)
    gbSearch: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Imprimir: TButton;
    Button3: TButton;
    txtCliente: TEdit;
    btnClientes: TButton;
    GridView1: TGridView;
    gbClientes: TGroupBox;
    gvClientes: TGridView;
    chkClientes: TCheckBox;
    btnClientesOK: TButton;
    btnTodosClientes: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    Label4: TLabel;
    txtMeta: TEdit;
    CopiarOrden1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure BindGrid();
    procedure BindClientes();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Exportar1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure btnClientesClick(Sender: TObject);
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView; Button: TButton);
    procedure btnClientesOKClick(Sender: TObject);
    procedure chkClientesClick(Sender: TObject);
    procedure btnTodosClientesClick(Sender: TObject);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCumplimientoTiempoEntrega: TfrmCumplimientoTiempoEntrega;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main, ReporteCumplimientoFechaEntregaQr;

{$R *.dfm}

procedure TfrmCumplimientoTiempoEntrega.FormCreate(Sender: TObject);
begin
  deFrom.Date := DateAdd(Now,-5,daDays);
  deTo.Date := Now;

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  BindClientes();
  chkClientes.Checked := False;
  btnClientesOKClick(nil);

  BindGrid();
end;

procedure TfrmCumplimientoTiempoEntrega.BindGrid();
var SQLStr : String;
workdays, goal, total, ahead, onTime, delay:Integer;
aheadDays, delayDays, totalDays : Double;
begin
    SQLStr := 'SELECT O.ITE_Nombre, CONVERT(VARCHAR, O.Recibido, 101) AS Recibido, CONVERT(VARCHAR, O.Interna, 101) AS Interna, CONVERT(VARCHAR, IT.ITS_DTStop, 101) AS ITS_DTStop,  ' +
              'DATEDIFF(dd, O.Recibido, IT.ITS_DTStop) AS Days, ' +
              'DATEDIFF(ww, O.Recibido, IT.ITS_DTStop) * 2 AS Weekenddays, ' +
              '(SELECT COUNT(*) FROM tblNonWorkingDay WHERE NonWorkingDay BETWEEN O.Recibido AND IT.ITS_DTStop) AS NonWorkingDays, ' +
              'DATEDIFF(dd, O.Recibido, IT.ITS_DTStop) - (DATEDIFF(ww, O.Recibido, IT.ITS_DTStop) * 2) - (SELECT COUNT(*) FROM tblNonWorkingDay WHERE NonWorkingDay BETWEEN O.Recibido AND IT.ITS_DTStop) AS WorkDays ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks IT ON O.ITE_Nombre = IT.ITE_Nombre AND ' +
              '(IT.ITS_DTStop >= ' + QuotedStr(deFrom.Text) + ' AND IT.ITS_DTStop <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99') + ') ' +
              'INNER JOIN tblTareas T ON IT.TAS_Id = T.Id AND T.Nombre = ''VentasFinal'' AND ITS_DTStop IS NOT NULL ';

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + 'WHERE SUBSTRING(O.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    SQLStr := SQLStr + 'ORDER BY O.ITE_Nombre ';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    goal := StrToInt(txtMeta.Text);
    ahead := 0;
    onTime := 0;
    delay := 0;
    total := 0;
    aheadDays := 0;
    delayDays := 0;
    totalDays := 0;
    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0, GridView1.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        GridView1.Cells[1, GridView1.RowCount -1] := VarToStr(Qry['Recibido']);
        GridView1.Cells[2, GridView1.RowCount -1] := VarToStr(Qry['ITS_DTStop']);
        workdays := StrToInt(VarToStr(Qry['WorkDays']));
        GridView1.Cells[3, GridView1.RowCount -1] := VarToStr(Qry['WorkDays']);

        if workdays < goal then begin
          GridView1.Cells[4, GridView1.RowCount -1] := '+' + IntToStr(abs(workdays - goal));
          GridView1.Cell[4, GridView1.RowCount -1].Color := $00C6EFCE;
          ahead := ahead + 1;
          aheadDays := aheadDays + abs(workdays - goal);
        end
        else if workdays > goal then begin
          GridView1.Cells[4, GridView1.RowCount -1] := '-' + IntToStr(abs(workdays - goal));
          GridView1.Cell[4, GridView1.RowCount -1].Color := $00CEC7FF;
          delay := delay + 1;
          delayDays := delayDays + abs(workdays - goal);
        end
        else begin
          GridView1.Cells[4, GridView1.RowCount -1] := IntToStr(workdays - goal);
          onTime := onTime + 1;
        end;

        total := total + 1;
        totalDays := totalDays + workdays;
        Qry.Next;
    End;

    if (total > 0) then begin
      GridView1.AddRow(1);
      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Total Ordenes';
      GridView1.Cells[4, GridView1.RowCount -1] := IntToStr(total);

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Total de Ordenes Adelantadas';
      GridView1.Cells[4, GridView1.RowCount -1] := IntToStr(ahead);

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Total de Ordenes a Tiempo';
      GridView1.Cells[4, GridView1.RowCount -1] := IntToStr(onTime);

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Total de Ordenes Atrasadas';
      GridView1.Cells[4, GridView1.RowCount -1] := IntToStr(delay);

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Dias Promedio de Adelanto';
      if ahead > 0  then begin
        GridView1.Cells[4, GridView1.RowCount -1] := FormatFloat('###0.00', (aheadDays / ahead));
      end;

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Dias Promedio de Atraso';
      if delay > 0  then begin
        GridView1.Cells[4, GridView1.RowCount -1] := FormatFloat('###0.00', (delayDays / delay));
      end;

      GridView1.AddRow(1);
      GridView1.Cells[3, GridView1.RowCount -1] := 'Tiempo Promedio de Todas las Ordenes';
      GridView1.Cells[4, GridView1.RowCount -1] := FormatFloat('###0.00', (totalDays / total));
    end;

    Qry.Close;
end;

procedure TfrmCumplimientoTiempoEntrega.BindClientes();
var Qry2 : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '062';
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
        if (slClientes.IndexOf(VarToStr(Qry2['Clave'])) <> -1) then begin
                gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
        end;
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmCumplimientoTiempoEntrega.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCumplimientoTiempoEntrega.Exportar1Click(Sender: TObject);
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

procedure TfrmCumplimientoTiempoEntrega.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
  // Format Cells
  xlBottom = -4107;
  xlLeft = -4131;
  xlRight = -4152;
  xlTop = -4160;
  xlLandscape = 2;  
var XApp : Variant;
Sheet, ColumnRange : Variant;
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
        ShowMessage('No se pudo abrir Microsoft Excel,  al parecer no esta instalado en el sistema.');
        Exit;
      end;

      XApp.Workbooks.Add(xlWorkSheet);
      Sheet := XApp.Workbooks[1].WorkSheets[1];
      Sheet.Name := 'Ordenes';

      for Col := 1 to Grid.Columns.Count do begin
              Sheet.Cells[1,Col] := Grid.Columns[Col - 1].Header.Caption;
              Sheet.Cells[1,Col].WrapText := true;
      end;

      for Row := 1 to Grid.RowCount do
                for Col := 1 to Grid.Columns.Count do
                        Sheet.Cells[Row + 1,Col] := Grid.Cells[Col - 1,Row - 1];

      // Change the Column Width.
      ColumnRange := Sheet.Columns;
      ColumnRange.Columns[1].ColumnWidth := 12;

      ColumnRange.Columns[2].ColumnWidth := 12;

      ColumnRange.Columns[3].ColumnWidth := 12;

      ColumnRange.Columns[4].ColumnWidth := 36;

      ColumnRange.Columns[5].ColumnWidth := 16.2;
{
      ColumnRange.Columns[6].ColumnWidth := 7.45;

      ColumnRange.Columns[7].ColumnWidth := 12.30;

      ColumnRange.Columns[8].ColumnWidth := 24.30;
      ColumnRange.Columns[8].HorizontalAlignment := xlRight;

      ColumnRange.Columns[9].ColumnWidth := 4.57;

      ColumnRange.Columns[10].ColumnWidth := 6.30;

      ColumnRange.Columns[11].ColumnWidth := 7.45;

      ColumnRange.Columns[12].ColumnWidth := 7.60;
}

      XApp.ActiveSheet.PageSetup.LeftMargin  := XApp.InchesToPoints(0.2);
      XApp.ActiveSheet.PageSetup.RightMargin := XApp.InchesToPoints(0.2);
{
      XApp.ActiveSheet.PageSetup.TopMargin := XApp.InchesToPoints(0.25);
      XApp.ActiveSheet.PageSetup.BottomMargin := XApp.InchesToPoints(0.25);
      XApp.ActiveSheet.PageSetup.HeaderMargin := XApp.InchesToPoints(0);
      XApp.ActiveSheet.PageSetup.FooterMargin := XApp.InchesToPoints(0);

      XApp.ActiveSheet.PageSetup.Orientation := xlLandscape;
}

      //Sheet.Cells.Select;
      //Sheet.Cells.EntireColumn.AutoFit;

      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

      ShowMessage('El archivo se creo exitosamente.');
end;

procedure TfrmCumplimientoTiempoEntrega.btnClientesClick(Sender: TObject);
begin
  ShowSeleccionGrid(gbClientes, chkClientes, txtCliente, gvClientes, btnTodosClientes);
end;

procedure TfrmCumplimientoTiempoEntrega.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
          TextBox: TEdit; Grid: TGridView; Button: TButton);
begin
  if GroupBox.Visible = True then
  begin
          GroupBox.Visible := False;
  end
  else begin
      GroupBox.Width := TextBox.Width;
      GroupBox.Top := TextBox.Parent.Top + TextBox.Top + TextBox.Height;
      GroupBox.Left := TextBox.Left + 8;
      Grid.Width := GroupBox.Width - 12;
      Button.Width := GroupBox.Width - 12;

      GroupBox.Visible := True;
      CheckBox.Checked := False;
      Grid.Enabled := True;
      Button.Enabled := True;
  end;

end;

procedure TfrmCumplimientoTiempoEntrega.btnClientesOKClick(
  Sender: TObject);
begin
  ParseSelection(gbClientes,chkClientes,txtCliente,gvClientes);
end;

procedure TfrmCumplimientoTiempoEntrega.chkClientesClick(Sender: TObject);
begin
  gvClientes.Enabled := not chkClientes.Checked;
  btnTodosClientes.Enabled := not chkClientes.Checked;
end;

procedure TfrmCumplimientoTiempoEntrega.btnTodosClientesClick(
  Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosClientes, gvClientes);
end;

procedure TfrmCumplimientoTiempoEntrega.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
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

procedure TfrmCumplimientoTiempoEntrega.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid: TGridView);
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

procedure TfrmCumplimientoTiempoEntrega.Button1Click(Sender: TObject);
begin
  if txtMeta.Text = '' then begin
    ShowMessage('Tiempo Meta es requerido.');
    Exit;
  end;

  if not IsNumeric(txtMeta.Text) then begin
    ShowMessage('Tiempo Meta debe de ser numerico.');
    Exit;
  end;

  BindGrid();
end;

procedure TfrmCumplimientoTiempoEntrega.Button3Click(Sender: TObject);
begin
Exportar1Click(nil);
end;

procedure TfrmCumplimientoTiempoEntrega.CopiarOrden1Click(Sender: TObject);
var Grid :TGridView;
begin
  Grid := ( ( (Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent as TGridView);

  Clipboard.AsText := Grid.Cells[0,Grid.SelectedRow]
end;

procedure TfrmCumplimientoTiempoEntrega.ImprimirClick(Sender: TObject);
var dataSet: TADODataSet;
i: Integer;
sDisplayCols: TStringList;
begin
    sDisplayCols := TStringList.Create;
    sDisplayCols.CommaText := '"Orden de Trabajo","Fecha Entrada","Fecha Terminacion","Dias","Adelanto Atrazo"';

    dataSet := TADODataSet.Create(nil);
    for i := 0 to (sDisplayCols.Count - 1) do
    begin
        if i <> 3 then begin
          with dataSet.FieldDefs.AddFieldDef do
          begin
            DataType := ftString;
            Name := sDisplayCols[i];
          end;
        end
        else begin
          with dataSet.FieldDefs.AddFieldDef do
          begin
            DataType := ftMemo;
            Name := sDisplayCols[i];
          end;
        end
    end;
    dataSet.CreateDataSet;

    for i:= 0 to GridView1.RowCount - 1 do
    begin
      dataSet.InsertRecord([GridView1.Cells[0,i],GridView1.Cells[1,i], GridView1.Cells[2,i], GridView1.Cells[3,i], GridView1.Cells[4,i]]);
    end;

      Try
      Begin
        Application.Initialize;
        Application.CreateForm(TqrCumplimientoTiempoEntrega, qrCumplimientoTiempoEntrega);
      end;
      except
        i := 0; //do nothing form is already created.
      end;

    qrCumplimientoTiempoEntrega.QRSubDetail1.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblOrdendeTrabajo.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblOrdendeTrabajo.DataField := 'Orden de Trabajo';

    qrCumplimientoTiempoEntrega.lblFechaEntrada.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblFechaEntrada.DataField := 'Fecha Entrada';

    qrCumplimientoTiempoEntrega.lblFechaTerminacion.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblFechaTerminacion.DataField := 'Fecha Terminacion';

    qrCumplimientoTiempoEntrega.lblDiasUtilizados.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblDiasUtilizados.DataField := 'Dias';

    qrCumplimientoTiempoEntrega.lblDiasAdelantoAtrazo.DataSet := dataSet;
    qrCumplimientoTiempoEntrega.lblDiasAdelantoAtrazo.DataField := 'Adelanto Atrazo';

    qrCumplimientoTiempoEntrega.Preview;
    qrCumplimientoTiempoEntrega.Free;

end;

end.
