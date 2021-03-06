unit ReporteMaterialesPorOrden;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,
  Menus;

type
  TfrmMaterialesPorOrden = class(TForm)
    gbSearch: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Imprimir: TButton;
    chkDlls: TCheckBox;
    Button3: TButton;
    Label3: TLabel;
    txtCliente: TEdit;
    btnClientes: TButton;
    GridView1: TGridView;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    SaveDialog1: TSaveDialog;
    gbClientes: TGroupBox;
    gvClientes: TGridView;
    chkClientes: TCheckBox;
    btnClientesOK: TButton;
    btnTodosClientes: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindGrid();
    procedure BindClientes();
    procedure Button1Click(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView);
    procedure SelectOrUnselectAll(Button: TButton; Grid:TGridView);
    procedure chkClientesClick(Sender: TObject);
    procedure btnTodosClientesClick(Sender: TObject);
    procedure btnClientesOKClick(Sender: TObject);
    procedure btnClientesClick(Sender: TObject);
    procedure ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid:TGridView; Button: TButton);
    procedure ImprimirClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMaterialesPorOrden: TfrmMaterialesPorOrden;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main, ReporteMaterialesPorOrdenQr;

{$R *.dfm}

procedure TfrmMaterialesPorOrden.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmMaterialesPorOrden.FormCreate(Sender: TObject);
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

procedure TfrmMaterialesPorOrden.BindGrid();
var SQLStr, ordenActual, ordenAnterior, totalHerramienta : String;
i:Integer;
slColumns, sDisplayCols: TStringList;
w: Real;
sumTotalHerramienta, valorOrden, resumenTotalHerramienta, resumenTotalOrden : Double;
columnSize: array of Integer;
begin
    sDisplayCols := TStringList.Create;
    sDisplayCols.CommaText := '"Orden de Trabajo","Cant. Orden","Descripcion","No.Parte","Material","Valor Orden","Tecnico","Desc. Herramienta","Cant.","Precio Unitario","Total Herramienta","% Individual"';

    slColumns := TStringList.Create;
    SQLStr := 'SELECT O.ITE_Nombre AS [Orden de Trabajo], O.Ordenada AS [Cant. Orden],O.Producto AS [Descripcion], ' +
              'O.Numero AS [No.Parte],M.MAT_Numero AS [Material], ' +
              'O.Ordenada * O.Unitario AS [Valor Orden], E.Nombre AS [Tecnico], M.MAT_Descripcion AS [Desc. Herramienta], ' +
              'SD.SD_Cantidad AS [Cant.], M.MAT_UltimoCosto AS [Precio Unitario], ' +
              'ROUND(SD.SD_Cantidad * M.MAT_UltimoCosto,2) AS [Total Herramienta], ' +
              'CASE WHEN O.Ordenada * O.Unitario = 0.0 THEN 0.0 ELSE ROUND(((SD.SD_Cantidad * M.MAT_UltimoCosto) / (O.Ordenada * O.Unitario)) * 100,2) END AS [% Individual] ' +
              'FROM tblSalidasDetalle SD ' +
              'INNER JOIN tblSalidas S ON S.SAL_ID = SD.SAL_ID ' +
              'INNER JOIN tblOrdenes O ON O.ITE_Nombre = S.SAL_Orden ' +
              'INNER JOIN tblItemTasks IT ON IT.ITE_Nombre = O.ITE_Nombre AND IT.TAS_Id = 19 ' +
              'INNER JOIN tblMateriales M ON SD.MAT_ID = M.MAT_ID ' +
              'LEFT OUTER JOIN tblEmpleados E ON S.SAL_Solicitado = E.ID ' +
              'WHERE (IT.ITS_DTStop >= ' + QuotedStr(deFrom.Text) + ' AND IT.ITS_DTStop <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99') + ') ';

    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + ' AND SUBSTRING(O.ITE_Nombre,4,3) IN (''' +
        StringReplace(txtCliente.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    SQLStr := SQLStr + 'ORDER BY O.ITE_Nombre ';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    GridView1.Columns.Clear;
    SetLength(columnSize, Qry.Fields.Count);

    for i := 0 to Qry.Fields.Count - 1 do begin
       slColumns.Add(Qry.Fields[i].FieldName);

       columnSize[i] := Length(Qry.Fields[i].FieldName);

       GridView1.Columns.Add(TTextualColumn);
       GridView1.Columns[i].Header.Caption := Qry.Fields[i].FieldName;
       GridView1.Columns[i].Width := 100;
       if (sDisplayCols.IndexOf(Qry.Fields[i].FieldName) = -1) then
         GridView1.Columns[i].Visible := False;
    end;

    resumenTotalHerramienta := 0.0;
    resumenTotalOrden := 0.0;
    sumTotalHerramienta := 0.0;
    valorOrden := 0.0;
    ordenAnterior := '';
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        ordenActual :=  VarToStr(Qry['Orden de Trabajo']);
        if  (ordenActual <>  ordenAnterior) and ('' <> ordenAnterior) then
        begin
                GridView1.Cells[2,GridView1.RowCount -1] := 'Totales';
                GridView1.Cells[5,GridView1.RowCount -1] := FloatToStr(valorOrden);
                GridView1.Cells[10,GridView1.RowCount -1] := FloatToStr(sumTotalHerramienta);
                if valorOrden = 0 then
                  GridView1.Cells[11,GridView1.RowCount -1] := '0%'
                else
                  GridView1.Cells[11,GridView1.RowCount -1] := FormatFloat('###0.00', (sumTotalHerramienta / valorOrden) * 100) + '%';

                resumenTotalHerramienta := resumenTotalHerramienta + sumTotalHerramienta;
                resumenTotalOrden := resumenTotalOrden + valorOrden;
                sumTotalHerramienta := 0.0;
                GridView1.AddRow(1);
        end;

        totalHerramienta :=  VarToStr(Qry['Total Herramienta']);
        sumTotalHerramienta := sumTotalHerramienta + StrToFloat(totalHerramienta);

        valorOrden := StrToFloat(VarToStr(Qry['Valor Orden']));

        for i := 0 to (slColumns.Count - 1) do
        begin
                GridView1.Cells[i,GridView1.RowCount -1] := VarToStr(Qry[slColumns[i]]);
                if '% Individual' = slColumns[i] then
                        GridView1.Cells[i,GridView1.RowCount -1] := GridView1.Cells[i,GridView1.RowCount -1] + '%';
                if Length(GridView1.Cells[i,GridView1.RowCount -1]) > columnSize[i] then
                        columnSize[i] := Length(GridView1.Cells[i,GridView1.RowCount -1]);
        end;

        ordenAnterior := ordenActual;
        Qry.Next;
    End;

    if GridView1.RowCount > 0 then
    begin
      GridView1.AddRow(1);
      GridView1.Cells[2,GridView1.RowCount -1] := 'Totales';
      GridView1.Cells[5,GridView1.RowCount -1] := FloatToStr(valorOrden);
      GridView1.Cells[10,GridView1.RowCount -1] := FloatToStr(sumTotalHerramienta);
      if valorOrden = 0 then
        GridView1.Cells[11,GridView1.RowCount -1] := '0%'
      else
        GridView1.Cells[11,GridView1.RowCount -1] := FormatFloat('###0.00', (sumTotalHerramienta / valorOrden) * 100) + '%';

      resumenTotalHerramienta := resumenTotalHerramienta + sumTotalHerramienta;
      resumenTotalOrden := resumenTotalOrden + valorOrden;

      GridView1.AddRow(1);
      GridView1.AddRow(1);
      GridView1.Cells[2,GridView1.RowCount -1] := 'Resumen Total';
      GridView1.Cells[5,GridView1.RowCount -1] := FloatToStr(resumenTotalOrden);
      GridView1.Cells[10,GridView1.RowCount -1] := FloatToStr(resumenTotalHerramienta);
      if valorOrden = 0 then
        GridView1.Cells[11,GridView1.RowCount -1] := '0%'
      else
        GridView1.Cells[11,GridView1.RowCount -1] := FormatFloat('###0.00', (resumenTotalHerramienta / resumenTotalOrden) * 100) + '%';

    end;

    for i:= Low(columnSize) to High(columnSize) do
    begin
        w := 6.38;
        if columnSize[i] >= 45 then w := 5.8;
        GridView1.Columns[i].Width := Trunc(columnSize[i] * w);
    end;
end;

procedure TfrmMaterialesPorOrden.Button1Click(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmMaterialesPorOrden.Exportar1Click(Sender: TObject);
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

procedure TfrmMaterialesPorOrden.Button3Click(Sender: TObject);
begin
Exportar1Click(nil);
end;

procedure TfrmMaterialesPorOrden.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
  // Format Cells
  xlBottom = -4107;
  xlLeft = -4131;
  xlRight = -4152;
  xlTop = -4160;
  xlLandscape = 2;  
var XApp : Variant;
Sheet,ColumnRange : Variant;
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
      ColumnRange.Columns[1].ColumnWidth := 9;
      ColumnRange.Columns[1].HorizontalAlignment := xlRight;

      ColumnRange.Columns[2].ColumnWidth := 5;
      ColumnRange.Columns[2].HorizontalAlignment := xlRight;

      ColumnRange.Columns[3].ColumnWidth := 9.15;

      ColumnRange.Columns[4].ColumnWidth := 13.45;

      ColumnRange.Columns[5].ColumnWidth := 7;

      ColumnRange.Columns[6].ColumnWidth := 7.45;

      ColumnRange.Columns[7].ColumnWidth := 12.30;

      ColumnRange.Columns[8].ColumnWidth := 24.30;
      ColumnRange.Columns[8].HorizontalAlignment := xlRight;

      ColumnRange.Columns[9].ColumnWidth := 4.57;

      ColumnRange.Columns[10].ColumnWidth := 6.30;

      ColumnRange.Columns[11].ColumnWidth := 7.45;

      ColumnRange.Columns[12].ColumnWidth := 7.60;


      //Sheet.Cells.Select;
      //Sheet.Cells.EntireColumn.AutoFit;

      XApp.ActiveSheet.PageSetup.LeftMargin  := XApp.InchesToPoints(0.2);
      XApp.ActiveSheet.PageSetup.RightMargin := XApp.InchesToPoints(0.2);
      XApp.ActiveSheet.PageSetup.TopMargin := XApp.InchesToPoints(0.25);
      XApp.ActiveSheet.PageSetup.BottomMargin := XApp.InchesToPoints(0.25);
      XApp.ActiveSheet.PageSetup.HeaderMargin := XApp.InchesToPoints(0);
      XApp.ActiveSheet.PageSetup.FooterMargin := XApp.InchesToPoints(0);

      XApp.ActiveSheet.PageSetup.Orientation := xlLandscape;

      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

      ShowMessage('El archivo se creo exitosamente.');
end;

procedure TfrmMaterialesPorOrden.SelectOrUnselectAll(Button: TButton; Grid:TGridView);
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

procedure TfrmMaterialesPorOrden.ParseSelection(GroupBox: TGroupBox; CheckBox: TCheckBox; TextBox: TEdit; Grid: TGridView);
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

procedure TfrmMaterialesPorOrden.chkClientesClick(Sender: TObject);
begin
  gvClientes.Enabled := not chkClientes.Checked;
  btnTodosClientes.Enabled := not chkClientes.Checked;
end;

procedure TfrmMaterialesPorOrden.btnTodosClientesClick(Sender: TObject);
begin
  SelectOrUnselectAll(btnTodosClientes, gvClientes);
end;

procedure TfrmMaterialesPorOrden.BindClientes();
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

procedure TfrmMaterialesPorOrden.btnClientesOKClick(Sender: TObject);
begin
  ParseSelection(gbClientes,chkClientes,txtCliente,gvClientes);
end;

procedure TfrmMaterialesPorOrden.btnClientesClick(Sender: TObject);
begin
  ShowSeleccionGrid(gbClientes, chkClientes, txtCliente, gvClientes, btnTodosClientes);
end;

procedure TfrmMaterialesPorOrden.ShowSeleccionGrid(GroupBox: TGroupBox; CheckBox: TCheckBox;
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

procedure TfrmMaterialesPorOrden.ImprimirClick(Sender: TObject);
var dataSet: TADODataSet;
i,p: Integer;
desc, printerName: String;
Strings, sDisplayCols: TStringList;
list: TStrings;
begin
    sDisplayCols := TStringList.Create;
    sDisplayCols.CommaText := '"Orden de Trabajo","Cantidad Orden","Descripcion","No.Parte","Material","Valor Orden","Tecnico","Desc. Herramienta","Cantidad","Precio Unitario","Total Herramienta","% Individual"';

    dataSet := TADODataSet.Create(nil);
    for i := 0 to (sDisplayCols.Count - 1) do
    begin
        with dataSet.FieldDefs.AddFieldDef do
        begin
          DataType := ftString;
          Name := sDisplayCols[i];
        end;
    end;

    dataSet.CreateDataSet;
    Strings := TStringList.Create;

    for i:= 0 to GridView1.RowCount - 1 do
    begin
      dataSet.InsertRecord([GridView1.Cells[0,i],GridView1.Cells[1,i], GridView1.Cells[2,i], GridView1.Cells[3,i], GridView1.Cells[4,i],
      GridView1.Cells[5,i], GridView1.Cells[6,i], GridView1.Cells[7,i], GridView1.Cells[8,i], GridView1.Cells[9,i], GridView1.Cells[10,i],
      GridView1.Cells[11,i]]);
    end;

      Try
      Begin
        Application.Initialize;
        Application.CreateForm(TqrReporteMaterialesPorOrden,qrReporteMaterialesPorOrden);
      end;
      except
        i := 0; //do nothing form is already created.
      end;

//    qrFactura.lblFecha.Caption := FormatDateTime('dd/mm/yyyy', deFecha.Date);

    qrReporteMaterialesPorOrden.QRSubDetail1.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblOrdendeTrabajo.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblOrdendeTrabajo.DataField := 'Orden de Trabajo';

    qrReporteMaterialesPorOrden.lblCantidadOrden.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblCantidadOrden.DataField := 'Cantidad Orden';

    qrReporteMaterialesPorOrden.lblDescripcion.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblDescripcion.DataField := 'Descripcion';

    qrReporteMaterialesPorOrden.lblNoParte.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblNoParte.DataField := 'No.Parte';

    qrReporteMaterialesPorOrden.lblMaterial.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblMaterial.DataField := 'Material';

    qrReporteMaterialesPorOrden.lblValorOrden.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblValorOrden.DataField := 'Valor Orden';

    qrReporteMaterialesPorOrden.lblTecnico.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblTecnico.DataField := 'Tecnico';

    qrReporteMaterialesPorOrden.lblDescHerramienta.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblDescHerramienta.DataField := 'Desc. Herramienta';

    qrReporteMaterialesPorOrden.lblCantidad.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblCantidad.DataField := 'Cantidad';

    qrReporteMaterialesPorOrden.lblPrecioUnitario.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblPrecioUnitario.DataField := 'Precio Unitario';

    qrReporteMaterialesPorOrden.lblTotalHerramienta.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblTotalHerramienta.DataField := 'Total Herramienta';

    qrReporteMaterialesPorOrden.lblIndividual.DataSet := dataSet;
    qrReporteMaterialesPorOrden.lblIndividual.DataField := '% Individual';


    qrReporteMaterialesPorOrden.Preview;
    qrReporteMaterialesPorOrden.Free;
end;

end.
