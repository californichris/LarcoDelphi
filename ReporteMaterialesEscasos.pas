unit ReporteMaterialesEscasos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,
  Menus;

type
  TfrmEscasos = class(TForm)
    GridView1: TGridView;
    gbSearch: TGroupBox;
    Label9: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    txtMaterialIDs: TEdit;
    Button1: TButton;
    Imprimir: TButton;
    ddlTipo: TComboBox;
    txtID: TEdit;
    btnID: TButton;
    txtDesc: TEdit;
    btnDesc: TButton;
    txtFraccion: TEdit;
    btnFraccion: TButton;
    txtMaterial: TEdit;
    btnMaterial: TButton;
    Button3: TButton;
    GroupBox1: TGroupBox;
    gvOpciones: TGridView;
    chkTodos: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    GroupBox2: TGroupBox;
    gvOpciones2: TGridView;
    chkTodos2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    GroupBox3: TGroupBox;
    gvOpciones3: TGridView;
    chkTodos3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    GroupBox6: TGroupBox;
    gvOpciones6: TGridView;
    chkTodos6: TCheckBox;
    btnOK6: TButton;
    btnTodos6: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    Label4: TLabel;
    Label1: TLabel;
    txtValor: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindOpciones(field : String; table : String; Grid:TGridView);
    procedure BindMateriales();
    procedure BindTiposMaterial();
    procedure BindGrid();
    function getWhere():String;    
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure btnIDClick(Sender: TObject);
    procedure btnDescClick(Sender: TObject);
    procedure btnFraccionClick(Sender: TObject);
    procedure btnMaterialClick(Sender: TObject);
    procedure chkTodosClick(Sender: TObject);
    procedure chkTodos2Click(Sender: TObject);
    procedure chkTodos3Click(Sender: TObject);
    procedure chkTodos6Click(Sender: TObject);
    procedure btnOKSelected(groupBox : TGroupBox; checkBox : TCheckBox; textBox: TEdit; grid : TGridView);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnOK6Click(Sender: TObject);
    procedure SeleccionarTodos(button: TButton; grid : TGridView);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure btnTodos6Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    function getIntValue(value: boolean):Integer;
    procedure Exportar1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure ddlTipoChange(Sender: TObject);
    procedure txtValorKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEscasos: TfrmEscasos;
  Qry : TADOQuery;
  Conn : TADOConnection;

implementation

uses Main;

{$R *.dfm}

procedure TfrmEscasos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmEscasos.FormCreate(Sender: TObject);
begin
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  BindOpciones('MAT_Numero','tblMateriales',gvOpciones);
  BindOpciones('MAT_Descripcion','tblMateriales',gvOpciones2);
  BindOpciones('MAT_Fraccion','tblMateriales',gvOpciones3);
  BindTiposMaterial();
  BindGrid();
end;

procedure TfrmEscasos.BindGrid();
var SQLStr : String;
i:Integer;
slColumns: TStringList;
begin
    slColumns := TStringList.Create;

    SQLStr := 'MaterialesEscasos ' + QuotedStr(ddlTipo.Text) + ',' + QuotedStr(getWhere()) +
              ',' + QuotedStr(txtValor.Text);


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    GridView1.Columns.Clear;
    for i := 0 to Qry.Fields.Count - 1 do begin
       slColumns.Add(Qry.Fields[i].FieldName);
       GridView1.Columns.Add(TTextualColumn);
       GridView1.Columns[i].Header.Caption := Qry.Fields[i].FieldName;
       GridView1.Columns[i].Width := 100;
       if GridView1.Columns[i].Header.Caption = 'Descripcion' then
           GridView1.Columns[i].Width := 350
       else if (GridView1.Columns[i].Header.Caption = 'Existenia') or
               (GridView1.Columns[i].Header.Caption = 'Minima') or
               (GridView1.Columns[i].Header.Caption = 'Maxima') or
               (GridView1.Columns[i].Header.Caption = 'Ideal') then
           GridView1.Columns[i].Width := 60;
    end;

   GridView1.Columns.Add(TCheckBoxColumn);
   GridView1.Columns[GridView1.Columns.Count - 1].Header.Caption := '';
   GridView1.Columns[GridView1.Columns.Count - 1].Options := GridView1.Columns[GridView1.Columns.Count - 1].Options + [coEditing];
   GridView1.Columns[GridView1.Columns.Count - 1].Width := 40;

   GridView1.Columns.Add(TTextualColumn);
   GridView1.Columns[GridView1.Columns.Count - 1].Header.Caption := 'Cantidad';
   GridView1.Columns[GridView1.Columns.Count - 1].Options := GridView1.Columns[GridView1.Columns.Count - 1].Options + [coEditing];
   GridView1.Columns[GridView1.Columns.Count - 1].Width := 60;

    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        for i := 0 to (slColumns.Count - 1) do
        begin
                GridView1.Cells[i,GridView1.RowCount -1] := VarToStr(Qry[slColumns[i]]);
        end;
       Qry.Next;
    End;    


//    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(GridView1.RowCount) +
//                        '   Cantidad Piezas Larco : ' + IntToStr(giCantidad) +
//                        '   Cantidad Piezas Cliente : ' + IntToStr(giCantCliente) +
//                        '   Diferencia : ' + IntToStr(giCantidad - giCantCliente);

end;

function TfrmEscasos.getWhere():String;
var SQLWhere: String;
begin

    SQLWhere := '';
    if txtID.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' M.MAT_Numero IN (''' +
        StringReplace(txtID.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtDesc.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' M.MAT_Descripcion IN (''' +
        StringReplace(txtDesc.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtFraccion.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' M.MAT_Fraccion IN (''' +
        StringReplace(txtFraccion.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtMaterial.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' M.MAT_Tipo IN (' + txtMaterialIDs.Text + ') ';
    end;

    if SQLWhere <> '' then SQLWhere := ' AND ' + SQLWhere;
    result := SQLWhere;
end;


procedure TfrmEscasos.BindOpciones(field : String; table : String; Grid:TGridView);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT DISTINCT ' + field + ' FROM ' + table +' ORDER BY ' + field;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    Grid.ClearRows;
    While not Qry2.Eof do
    Begin
        Grid.AddRow(1);
        Grid.Cells[0,Grid.RowCount -1] := VarToStr(Qry2[field]);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmEscasos.BindTiposMaterial();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT TIP_ID, TIP_Descripcion FROM tblTiposMaterial ORDER BY TIP_Descripcion';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvOpciones6.ClearRows;
    While not Qry2.Eof do
    Begin
        gvOpciones6.AddRow(1);
        gvOpciones6.Cells[0,gvOpciones6.RowCount -1] := VarToStr(Qry2['TIP_ID']);
        gvOpciones6.Cells[1,gvOpciones6.RowCount -1] := VarToStr(Qry2['TIP_Descripcion']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmEscasos.BindMateriales();
begin
  BindOpciones('MAT_Numero','tblMateriales',gvOpciones)
end;

procedure TfrmEscasos.ExportGrid(Grid: TGridView;sFileName: String);
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

procedure TfrmEscasos.btnIDClick(Sender: TObject);
begin
  if GroupBox1.Visible = True then
  begin
          GroupBox1.Visible := False;
  end
  else begin
      GroupBox1.Visible := True;

      if txtID.Text = 'Todos' then
      begin
          chkTodos.Checked := True;
          gvOpciones.Enabled := False;
          btnTodos.Enabled := False;
      end
      else
      begin
          chkTodos.Checked := False;
          gvOpciones.Enabled := True;
          btnTodos.Enabled := True;
      end;

      GroupBox1.Top := txtID.Top + txtID.Height + 5;
      GroupBox1.Left := txtID.Left + 10;
  end;
end;

procedure TfrmEscasos.btnDescClick(Sender: TObject);
begin
  if GroupBox2.Visible = True then
  begin
          GroupBox2.Visible := False;
  end
  else begin
      GroupBox2.Visible := True;

      if txtDesc.Text = 'Todos' then
      begin
          chkTodos2.Checked := True;
          gvOpciones2.Enabled := False;
          btnTodos2.Enabled := False;
      end
      else
      begin
          chkTodos2.Checked := False;
          gvOpciones2.Enabled := True;
          btnTodos2.Enabled := True;
      end;

      GroupBox2.Top := txtDesc.Top + txtDesc.Height + 5;
      GroupBox2.Left := txtDesc.Left + 10;
  end;
end;

procedure TfrmEscasos.btnFraccionClick(Sender: TObject);
begin
  if GroupBox3.Visible = True then
  begin
          GroupBox3.Visible := False;
  end
  else begin
      GroupBox3.Visible := True;

      if txtFraccion.Text = 'Todos' then
      begin
          chkTodos3.Checked := True;
          gvOpciones3.Enabled := False;
          btnTodos3.Enabled := False;
      end
      else
      begin
          chkTodos3.Checked := False;
          gvOpciones3.Enabled := True;
          btnTodos3.Enabled := True;
      end;

      GroupBox3.Top := txtFraccion.Top + txtFraccion.Height + 5;
      GroupBox3.Left := txtFraccion.Left + 10;
  end;
end;

procedure TfrmEscasos.btnMaterialClick(Sender: TObject);
begin
  if GroupBox6.Visible = True then
  begin
          GroupBox6.Visible := False;
  end
  else begin
      GroupBox6.Visible := True;

      if txtMaterial.Text = 'Todos' then
      begin
          chkTodos6.Checked := True;
          gvOpciones6.Enabled := False;
          btnTodos6.Enabled := False;
      end
      else
      begin
          chkTodos6.Checked := False;
          gvOpciones6.Enabled := True;
          btnTodos6.Enabled := True;
      end;

      GroupBox6.Top := txtMaterial.Top + txtMaterial.Height + 5;
      GroupBox6.Left := txtMaterial.Left + 10;
  end;
end;

procedure TfrmEscasos.chkTodosClick(Sender: TObject);
begin
  gvOpciones.Enabled := not chkTodos.Checked;
  btnTodos.Enabled := not chkTodos.Checked;
end;

procedure TfrmEscasos.chkTodos2Click(Sender: TObject);
begin
  gvOpciones2.Enabled := not chkTodos2.Checked;
  btnTodos2.Enabled := not chkTodos2.Checked;
end;

procedure TfrmEscasos.chkTodos3Click(Sender: TObject);
begin
  gvOpciones3.Enabled := not chkTodos3.Checked;
  btnTodos3.Enabled := not chkTodos3.Checked;
end;

procedure TfrmEscasos.chkTodos6Click(Sender: TObject);
begin
  gvOpciones6.Enabled := not chkTodos6.Checked;
  btnTodos6.Enabled := not chkTodos6.Checked;
end;

procedure TfrmEscasos.btnOKSelected(groupBox : TGroupBox; checkBox : TCheckBox; textBox: TEdit; grid : TGridView);
var i: integer;
sOpciones : String;
begin
  groupBox.Visible := False;
  if checkBox.Checked = True then begin
          textBox.Text := 'Todos';
  end
  else begin
        sOpciones := '';
        for i:= 0 to grid.RowCount - 1 do
        begin
                if grid.Cell[1,i].AsBoolean = True then
                begin
                        sOpciones := sOpciones + grid.Cells[0,i] + ',';
                end;
        end;
        textBox.Text := 'Todos';
        if sOpciones <> '' then
        begin
                textBox.Text :=  LeftStr(sOpciones,Length(sOpciones) - 1);
        end;
  end;
end;

procedure TfrmEscasos.btnOKClick(Sender: TObject);
begin
  btnOKSelected(GroupBox1, chkTodos, txtID, gvOpciones);
end;

procedure TfrmEscasos.btnOK2Click(Sender: TObject);
begin
  btnOKSelected(GroupBox2, chkTodos2, txtDesc, gvOpciones2);
end;

procedure TfrmEscasos.btnOK3Click(Sender: TObject);
begin
  btnOKSelected(GroupBox3, chkTodos3, txtFraccion, gvOpciones3);
end;

procedure TfrmEscasos.btnOK6Click(Sender: TObject);
var i: integer;
sOpciones, sOpcionesIDs : String;
begin
  GroupBox6.Visible := False;
  if chkTodos6.Checked = True then begin
          txtMaterial.Text := 'Todos';
          txtMaterialIds.Text := 'Todos';
  end
  else begin
        sOpciones := '';
        for i:= 0 to gvOpciones6.RowCount - 1 do
        begin
                if gvOpciones6.Cell[2,i].AsBoolean = True then
                begin
                        sOpcionesIDs := sOpcionesIDs + gvOpciones6.Cells[0,i] + ',';
                        sOpciones := sOpciones + gvOpciones6.Cells[1,i] + ',';
                end;
        end;
        txtMaterial.Text := 'Todos';
        txtMaterialIds.Text := 'Todos';
        if sOpciones <> '' then
        begin
                txtMaterial.Text :=  LeftStr(sOpciones,Length(sOpciones) - 1);
                txtMaterialIds.Text :=  LeftStr(sOpcionesIDs,Length(sOpcionesIDs) - 1);
        end;
  end;
end;

procedure TfrmEscasos.SeleccionarTodos(button: TButton; grid : TGridView);
var i: integer;
begin
  if UT(button.Caption) = UT('Seleccionar Todos') then begin
        button.Caption := 'Deseleccionar Todos';
        for i:= 0 to grid.RowCount - 1 do
        begin
                grid.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        button.Caption := 'Seleccionar Todos';
        for i:= 0 to grid.RowCount - 1 do
        begin
                grid.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmEscasos.btnTodosClick(Sender: TObject);
begin
  SeleccionarTodos(btnTodos, gvOpciones);
end;

procedure TfrmEscasos.btnTodos2Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos2, gvOpciones2);
end;

procedure TfrmEscasos.btnTodos3Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos3, gvOpciones3);
end;

procedure TfrmEscasos.btnTodos6Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos6, gvOpciones6);
end;

procedure TfrmEscasos.Button1Click(Sender: TObject);
begin
  BindGrid();
end;

function TfrmEscasos.getIntValue(value: boolean):Integer;
begin
  if(value = true) then
    result := 1
  else
    result := 0;
end;

procedure TfrmEscasos.Exportar1Click(Sender: TObject);
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

procedure TfrmEscasos.Button3Click(Sender: TObject);
begin
Exportar1Click(nil);
end;

procedure TfrmEscasos.ImprimirClick(Sender: TObject);
begin
  ShowMessage('Esta opcion no esta disponible por el momento');
end;

procedure TfrmEscasos.ddlTipoChange(Sender: TObject);
begin
txtValor.Enabled := False;
if 'Menor a Valor Especifico' = ddlTipo.Text then begin
   txtValor.Enabled := True;
end;

end;

procedure TfrmEscasos.txtValorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['0'..'9'] then
      begin
      end
  else if (Key = Chr(vk_Back)) then
      begin
      end
  else if (Key in ['.']) then
      begin
          if StrPos(PChar((Sender as TEdit).Text), '.') <> nil then
            Key := #0;
      end
 else
          Key := #0;
end;

end.
