unit ReporteEntradasSalidasLarco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,ReporteRelacion,LTCUtils, GridPrint,Columns,ColumnClasses,
  Menus;

type
  TfrmESLarco = class(TForm)
    GridView1: TGridView;
    gbSearch: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label9: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    lblColumnas: TLabel;
    txtMaterialIDs: TEdit;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Imprimir: TButton;
    ddlTipo: TComboBox;
    txtID: TEdit;
    btnID: TButton;
    txtDesc: TEdit;
    btnDesc: TButton;
    txtFraccion: TEdit;
    btnFraccion: TButton;
    txtTipoEntrada: TEdit;
    btnTipoEntrada: TButton;
    txtTipoImp: TEdit;
    btnTipoImp: TButton;
    txtMaterial: TEdit;
    btnMaterial: TButton;
    chkDlls: TCheckBox;
    Button2: TButton;
    Button3: TButton;
    GroupBox3: TGroupBox;
    gvOpciones3: TGridView;
    chkTodos3: TCheckBox;
    btnOK3: TButton;
    btnTodos3: TButton;
    GroupBox4: TGroupBox;
    gvOpciones4: TGridView;
    chkTodos4: TCheckBox;
    btnOK4: TButton;
    btnTodos4: TButton;
    GroupBox5: TGroupBox;
    gvOpciones5: TGridView;
    chkTodos5: TCheckBox;
    btnOK5: TButton;
    btnTodos5: TButton;
    GroupBox7: TGroupBox;
    gvOpciones7: TGridView;
    btnOK7: TButton;
    btnTodos7: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    Label10: TLabel;
    txtPais: TEdit;
    btnPais: TButton;
    GroupBox8: TGroupBox;
    gvOpciones8: TGridView;
    chkTodos8: TCheckBox;
    btnOK8: TButton;
    btnTodos8: TButton;
    txtPaisIDs: TEdit;
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
    GroupBox6: TGroupBox;
    gvOpciones6: TGridView;
    chkTodos6: TCheckBox;
    btnOK6: TButton;
    btnTodos6: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindOpciones(field : String; table : String; Grid:TGridView);
    procedure BindMateriales();
    procedure BindTiposEntradaAndImp();
    procedure BindTiposMaterial();
    procedure BindPaises();
    procedure BindGrid();
    function getWhere():String;    
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure btnIDClick(Sender: TObject);
    procedure btnDescClick(Sender: TObject);
    procedure btnFraccionClick(Sender: TObject);
    procedure btnTipoEntradaClick(Sender: TObject);
    procedure btnTipoImpClick(Sender: TObject);
    procedure btnMaterialClick(Sender: TObject);
    procedure chkTodosClick(Sender: TObject);
    procedure chkTodos2Click(Sender: TObject);
    procedure chkTodos3Click(Sender: TObject);
    procedure chkTodos4Click(Sender: TObject);
    procedure chkTodos5Click(Sender: TObject);
    procedure chkTodos6Click(Sender: TObject);
    procedure btnOKSelected(groupBox : TGroupBox; checkBox : TCheckBox; textBox: TEdit; grid : TGridView);
    procedure btnOKClick(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnOK3Click(Sender: TObject);
    procedure btnOK4Click(Sender: TObject);
    procedure btnOK5Click(Sender: TObject);
    procedure btnOK6Click(Sender: TObject);
    procedure SeleccionarTodos(button: TButton; grid : TGridView);
    procedure btnTodosClick(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure btnTodos3Click(Sender: TObject);
    procedure btnTodos4Click(Sender: TObject);
    procedure btnTodos5Click(Sender: TObject);
    procedure btnTodos6Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ddlTipoChange(Sender: TObject);
    function getIntValue(value: boolean):Integer;
    procedure Button2Click(Sender: TObject);
    procedure btnTodos7Click(Sender: TObject);
    procedure btnOK7Click(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure btnPaisClick(Sender: TObject);
    procedure chkTodos8Click(Sender: TObject);
    procedure btnTodos8Click(Sender: TObject);
    procedure btnOK8Click(Sender: TObject);
    procedure GridView1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmESLarco: TfrmESLarco;
  Qry : TADOQuery;
  Conn : TADOConnection;
  gbFirstTime : Boolean;

implementation

uses Main;

{$R *.dfm}

procedure TfrmESLarco.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmESLarco.FormCreate(Sender: TObject);
begin
  gbFirstTime := True;
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  deFrom.Date := Now;
  deTo.Date := Now;


  BindOpciones('MAT_Numero','tblMateriales',gvOpciones);
  BindOpciones('MAT_Descripcion','tblMateriales',gvOpciones2);
  BindOpciones('MAT_Fraccion','tblMateriales',gvOpciones3);
  BindTiposMaterial();
  BindPaises();
  BindTiposEntradaAndImp();
  BindGrid();
end;

procedure TfrmESLarco.BindGrid();
var SQLStr : String;
giCantidad, giCantCliente, i, iCol:Integer;
slColumns: TStringList;
begin
    giCantidad := 0;
    giCantCliente := 0;
    iCol := 0;
    slColumns := TStringList.Create;

    SQLStr := 'EntradasSalidasLarco ' + QuotedStr(ddlTipo.Text) + ',' + QuotedStr(deFrom.Text) + ',' +
               QuotedStr(deTo.Text  + ' 23:59:59.99') + ',' + IntToStr(getIntValue(chkDlls.Checked)) +
               ',' + QuotedStr(getWhere());


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    GridView1.Columns.Clear;
    gvOpciones7.ClearRows;
    for i := 0 to Qry.Fields.Count - 1 do begin
       //if 'Id' <> Qry.Fields[i].FieldName then begin
         slColumns.Add(Qry.Fields[i].FieldName);
       //end;

       GridView1.Columns.Add(TTextualColumn);
       GridView1.Columns[i].Header.Caption := Qry.Fields[i].FieldName;
       GridView1.Columns[i].Width := 100;

       if 'Id' = Qry.Fields[i].FieldName then begin
         GridView1.Columns[i].Visible := False;
       end;

       if 'Id' <> Qry.Fields[i].FieldName then begin
         gvOpciones7.AddRow(1);
         gvOpciones7.Cells[0,gvOpciones7.RowCount -1] := Qry.Fields[i].FieldName;
         gvOpciones7.Cell[1,gvOpciones7.RowCount -1].AsBoolean := True;
       end;
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


    if 'Entradas vs Salidas' = ddlTipo.Text then begin
        for i := 1 to GridView1.RowCount -1 do begin
          if GridView1.Cells[0, i] = GridView1.Cells[0, i - 1] then begin
            GridView1.DeleteRow(i);
          end;
        end;

        for i := 0 to GridView1.RowCount -1 do begin
          if GridView1.Cells[17, i] = '' then begin
            for iCol := 16 to GridView1.Columns.Count - 1 do begin
              GridView1.Cells[iCol, i] := '';
            end;
          end;
        end;

        {for i := 1 to GridView1.RowCount -1 do begin
          if (GridView1.Cells[2, i] = GridView1.Cells[2, i - 1]) and (GridView1.Cells[18, i] = GridView1.Cells[18, i - 1]) then begin
            for iCol := 17 to GridView1.Columns.Count - 1 do begin
              GridView1.Cells[iCol, i] := '';
            end;
          end;
        end;
         }

    end;

    if GridView1.RowCount > 0 then begin
        GridView1.SelectCell(0,0);
        if (not gbFirstTime) then begin
                GridView1.SetFocus;
                gbFirstTime := False;
        end;
    end    
//    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(GridView1.RowCount) +
//                        '   Cantidad Piezas Larco : ' + IntToStr(giCantidad) +
//                        '   Cantidad Piezas Cliente : ' + IntToStr(giCantCliente) +
//                        '   Diferencia : ' + IntToStr(giCantidad - giCantCliente);

end;

function TfrmESLarco.getWhere():String;
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

    if txtTipoEntrada.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' E.ENT_Nacional IN (''' +
        StringReplace(txtTipoEntrada.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtTipoImp.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' E.ENT_TipoImp IN (''' +
        StringReplace(txtTipoImp.Text,',',''',''',[rfReplaceAll, rfIgnoreCase]) + ''') ';
    end;

    if txtMaterial.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' M.MAT_Tipo IN (' + txtMaterialIDs.Text + ') ';
    end;

    if txtPais.Text <> 'Todos' then
    begin
        if SQLWhere <> '' then SQLWhere := SQLWhere + ' AND ';
        SQLWhere := SQLWhere + ' PA.PAIS_ID IN (' + txtPaisIDs.Text + ') ';
    end;

    if SQLWhere <> '' then SQLWhere := ' AND ' + SQLWhere;
    result := SQLWhere;
end;


procedure TfrmESLarco.BindOpciones(field : String; table : String; Grid:TGridView);
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

procedure TfrmESLarco.BindTiposMaterial();
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

procedure TfrmESLarco.BindPaises();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT PAIS_ID, PAIS_Nombre FROM tblPaises ORDER BY PAIS_Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvOpciones8.ClearRows;
    While not Qry2.Eof do
    Begin
        gvOpciones8.AddRow(1);
        gvOpciones8.Cells[0,gvOpciones8.RowCount -1] := VarToStr(Qry2['PAIS_ID']);
        gvOpciones8.Cells[1,gvOpciones8.RowCount -1] := VarToStr(Qry2['PAIS_Nombre']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmESLarco.BindTiposEntradaAndImp();
begin
    gvOpciones4.ClearRows;
    gvOpciones4.AddRow(1);
    gvOpciones4.Cells[0,gvOpciones4.RowCount -1] := 'Importado';
    gvOpciones4.AddRow(1);
    gvOpciones4.Cells[0,gvOpciones4.RowCount -1] := 'Nacional';

    gvOpciones5.ClearRows;
    gvOpciones5.AddRow(1);
    gvOpciones5.Cells[0,gvOpciones5.RowCount -1] := 'Importacion Temporal';
    gvOpciones5.AddRow(1);
    gvOpciones5.Cells[0,gvOpciones5.RowCount -1] := 'Importacion Definitiva';
end;

procedure TfrmESLarco.BindMateriales();
begin
  BindOpciones('MAT_Numero','tblMateriales',gvOpciones)
end;

procedure TfrmESLarco.ExportGrid(Grid: TGridView;sFileName: String);
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

procedure TfrmESLarco.btnIDClick(Sender: TObject);
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

procedure TfrmESLarco.btnDescClick(Sender: TObject);
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

procedure TfrmESLarco.btnFraccionClick(Sender: TObject);
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

procedure TfrmESLarco.btnTipoEntradaClick(Sender: TObject);
begin
  if GroupBox4.Visible = True then
  begin
          GroupBox4.Visible := False;
  end
  else begin
      GroupBox4.Visible := True;

      if txtTipoEntrada.Text = 'Todos' then
      begin
          chkTodos4.Checked := True;
          gvOpciones4.Enabled := False;
          btnTodos4.Enabled := False;
      end
      else
      begin
          chkTodos4.Checked := False;
          gvOpciones4.Enabled := True;
          btnTodos4.Enabled := True;
      end;

      GroupBox4.Top := txtTipoEntrada.Top + txtTipoEntrada.Height + 5;
      GroupBox4.Left := txtTipoEntrada.Left + 10;
  end;
end;

procedure TfrmESLarco.btnTipoImpClick(Sender: TObject);
begin
  if GroupBox5.Visible = True then
  begin
          GroupBox5.Visible := False;
  end
  else begin
      GroupBox5.Visible := True;

      if txtTipoImp.Text = 'Todos' then
      begin
          chkTodos5.Checked := True;
          gvOpciones5.Enabled := False;
          btnTodos5.Enabled := False;
      end
      else
      begin
          chkTodos5.Checked := False;
          gvOpciones5.Enabled := True;
          btnTodos5.Enabled := True;
      end;

      GroupBox5.Top := txtTipoImp.Top + txtTipoImp.Height + 5;
      GroupBox5.Left := txtTipoImp.Left + 10;
  end;
end;

procedure TfrmESLarco.btnMaterialClick(Sender: TObject);
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

procedure TfrmESLarco.chkTodosClick(Sender: TObject);
begin
  gvOpciones.Enabled := not chkTodos.Checked;
  btnTodos.Enabled := not chkTodos.Checked;
end;

procedure TfrmESLarco.chkTodos2Click(Sender: TObject);
begin
  gvOpciones2.Enabled := not chkTodos2.Checked;
  btnTodos2.Enabled := not chkTodos2.Checked;
end;

procedure TfrmESLarco.chkTodos3Click(Sender: TObject);
begin
  gvOpciones3.Enabled := not chkTodos3.Checked;
  btnTodos3.Enabled := not chkTodos3.Checked;
end;

procedure TfrmESLarco.chkTodos4Click(Sender: TObject);
begin
  gvOpciones4.Enabled := not chkTodos4.Checked;
  btnTodos4.Enabled := not chkTodos4.Checked;
end;

procedure TfrmESLarco.chkTodos5Click(Sender: TObject);
begin
  gvOpciones5.Enabled := not chkTodos5.Checked;
  btnTodos5.Enabled := not chkTodos5.Checked;
end;

procedure TfrmESLarco.chkTodos6Click(Sender: TObject);
begin
  gvOpciones6.Enabled := not chkTodos6.Checked;
  btnTodos6.Enabled := not chkTodos6.Checked;
end;

procedure TfrmESLarco.btnOKSelected(groupBox : TGroupBox; checkBox : TCheckBox; textBox: TEdit; grid : TGridView);
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

procedure TfrmESLarco.btnOKClick(Sender: TObject);
begin
  btnOKSelected(GroupBox1, chkTodos, txtID, gvOpciones);
end;

procedure TfrmESLarco.btnOK2Click(Sender: TObject);
begin
  btnOKSelected(GroupBox2, chkTodos2, txtDesc, gvOpciones2);
end;

procedure TfrmESLarco.btnOK3Click(Sender: TObject);
begin
  btnOKSelected(GroupBox3, chkTodos3, txtFraccion, gvOpciones3);
end;

procedure TfrmESLarco.btnOK4Click(Sender: TObject);
begin
  btnOKSelected(GroupBox4, chkTodos4, txtTipoEntrada, gvOpciones4);
end;

procedure TfrmESLarco.btnOK5Click(Sender: TObject);
begin
  btnOKSelected(GroupBox5, chkTodos5, txtTipoImp, gvOpciones5);
end;

procedure TfrmESLarco.btnOK6Click(Sender: TObject);
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

procedure TfrmESLarco.SeleccionarTodos(button: TButton; grid : TGridView);
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

procedure TfrmESLarco.btnTodosClick(Sender: TObject);
begin
  SeleccionarTodos(btnTodos, gvOpciones);
end;

procedure TfrmESLarco.btnTodos2Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos2, gvOpciones2);
end;

procedure TfrmESLarco.btnTodos3Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos3, gvOpciones3);
end;

procedure TfrmESLarco.btnTodos4Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos4, gvOpciones4);
end;

procedure TfrmESLarco.btnTodos5Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos5, gvOpciones5);
end;

procedure TfrmESLarco.btnTodos6Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos6.Caption) = UT('Seleccionar Todos') then begin
        btnTodos6.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvOpciones6.RowCount - 1 do
        begin
                gvOpciones5.Cell[2,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos6.Caption := 'Seleccionar Todos';
        for i:= 0 to gvOpciones5.RowCount - 1 do
        begin
                gvOpciones5.Cell[2,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmESLarco.Button1Click(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmESLarco.ddlTipoChange(Sender: TObject);
begin
  btnTipoEntrada.Enabled := True;
  btnTipoImp.Enabled := True;
  if ddlTipo.Text = 'Salidas' then begin
        txtTipoEntrada.Text := 'Todos';
        txtTipoImp.Text := 'Todos';

        btnTipoEntrada.Enabled := False;
        btnTipoImp.Enabled := False;
  end;

end;

function TfrmESLarco.getIntValue(value: boolean):Integer;
begin
  if(value = true) then
    result := 1
  else
    result := 0;
end;

procedure TfrmESLarco.Button2Click(Sender: TObject);
begin
  if GroupBox7.Visible = True then
  begin
          GroupBox7.Visible := False;
  end
  else begin
      GroupBox7.Visible := True;

      gvOpciones7.Enabled := True;
      btnTodos7.Enabled := True;

      GroupBox7.Top := lblColumnas.Top + lblColumnas.Height + 5;
      GroupBox7.Left := lblColumnas.Left + 10;
  end;
end;

procedure TfrmESLarco.btnTodos7Click(Sender: TObject);
begin
  SeleccionarTodos(btnTodos7, gvOpciones7);
end;

procedure TfrmESLarco.btnOK7Click(Sender: TObject);
var i, inc: integer;
begin
        for i:= 0 to gvOpciones7.RowCount - 1 do
        begin
                inc := 0;
                if ddlTipo.Text = 'Entradas vs Salidas' then inc := 1;

                if gvOpciones7.Cell[1,i].AsBoolean = False then
                begin
                        GridView1.Columns[i + inc].Visible := False;
                end
                else begin
                        GridView1.Columns[i + inc].Visible := True;
                end;
        end;
        GroupBox7.Visible := False;
end;

procedure TfrmESLarco.Exportar1Click(Sender: TObject);
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

procedure TfrmESLarco.Button3Click(Sender: TObject);
begin
Exportar1Click(nil);
end;

procedure TfrmESLarco.ImprimirClick(Sender: TObject);
begin
  ShowMessage('Esta opcion no esta disponible por el momento');
end;

procedure TfrmESLarco.btnPaisClick(Sender: TObject);
begin
  if GroupBox8.Visible = True then
  begin
          GroupBox8.Visible := False;
  end
  else begin
      GroupBox8.Visible := True;

      if txtPais.Text = 'Todos' then
      begin
          chkTodos8.Checked := True;
          gvOpciones8.Enabled := False;
          btnTodos8.Enabled := False;
      end
      else
      begin
          chkTodos8.Checked := False;
          gvOpciones8.Enabled := True;
          btnTodos8.Enabled := True;
      end;

      GroupBox8.Top := txtPais.Top + txtPais.Height + 5;
      GroupBox8.Left := txtPais.Left + 10;
  end;
end;

procedure TfrmESLarco.chkTodos8Click(Sender: TObject);
begin
  gvOpciones8.Enabled := not chkTodos8.Checked;
  btnTodos8.Enabled := not chkTodos8.Checked;
end;

procedure TfrmESLarco.btnTodos8Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos8.Caption) = UT('Seleccionar Todos') then begin
        btnTodos8.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvOpciones6.RowCount - 1 do
        begin
                gvOpciones8.Cell[2,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos8.Caption := 'Seleccionar Todos';
        for i:= 0 to gvOpciones8.RowCount - 1 do
        begin
                gvOpciones8.Cell[2,i].AsBoolean := False;
        end;
  end;

end;

procedure TfrmESLarco.btnOK8Click(Sender: TObject);
var i: integer;
sOpciones, sOpcionesIDs : String;
begin
  GroupBox8.Visible := False;
  if chkTodos8.Checked = True then begin
          txtPais.Text := 'Todos';
          txtPaisIDs.Text := 'Todos';
  end
  else begin
        sOpciones := '';
        for i:= 0 to gvOpciones8.RowCount - 1 do
        begin
                if gvOpciones8.Cell[2,i].AsBoolean = True then
                begin
                        sOpcionesIDs := sOpcionesIDs + gvOpciones8.Cells[0,i] + ',';
                        sOpciones := sOpciones + gvOpciones8.Cells[1,i] + ',';
                end;
        end;
        txtPais.Text := 'Todos';
        txtPaisIDs.Text := 'Todos';
        if sOpciones <> '' then
        begin
                txtPais.Text :=  LeftStr(sOpciones,Length(sOpciones) - 1);
                txtPaisIDs.Text :=  LeftStr(sOpcionesIDs,Length(sOpcionesIDs) - 1);
        end;
  end;

end;

procedure TfrmESLarco.GridView1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_LEFT then begin
      //ShowMessage(IntToStr(GridView1.SelectedColumn));
      if (GridView1.SelectedColumn - 8) > 1 then begin
        GridView1.ScrollToColumn(GridView1.Columns[GridView1.SelectedColumn - 8]);
        GridView1.SelectCell(GridView1.SelectedColumn - 8, GridView1.SelectedRow);
      end
      else begin
        GridView1.ScrollToColumn(GridView1.Columns[0]);
        GridView1.SelectCell(1, GridView1.SelectedRow);
      end;
      Application.ProcessMessages;
  end
  else if Key = VK_RIGHT then begin
      //ShowMessage(IntToStr(GridView1.SelectedColumn));
      //ShowMessage(IntToStr(GridView1.Columns.Count));
      if (GridView1.SelectedColumn + 8) < GridView1.Columns.Count then begin
        GridView1.ScrollToColumn(GridView1.Columns[GridView1.SelectedColumn + 8]);
        GridView1.SelectCell(GridView1.SelectedColumn + 8, GridView1.SelectedRow);
      end
      else begin
        GridView1.ScrollToColumn(GridView1.Columns[GridView1.Columns.Count - 1]);
        GridView1.SelectCell(GridView1.Columns.Count - 1, GridView1.SelectedRow);
      end;
      Application.ProcessMessages;

      //ShowMessage(IntToStr(GridView1.SelectedColumn));
      //ShowMessage('Right');
  end;
end;

end.
