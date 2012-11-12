unit ReporteEntradasSalidasPlano;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, sndkey32,
  ExtCtrls, StdCtrls, CellEditors, ScrollView, ComCtrls,ComObj,DateUtils,
  CustomGridViewControl, CustomGridView, GridView, Menus,Clipbrd,LTCUtils,Larco_functions;

type
  TfrmReporteESPlano = class(TForm)
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    Button1: TButton;
    btnBuscar: TButton;
    Button3: TButton;
    txtPlano: TEdit;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    GroupBox5: TGroupBox;
    gvPlanos: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    Copiar1: TMenuItem;
    OpenDialog1: TOpenDialog;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BindPlanos();
    procedure Button3Click(Sender: TObject);
    procedure BindGrid();
    procedure txtPlanoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtPlanoKeyPress(Sender: TObject; var Key: Char);
    procedure btnBuscarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure Copiar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmReporteESPlano: TfrmReporteESPlano;

implementation

uses Main;

{$R *.dfm}

procedure TfrmReporteESPlano.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmReporteESPlano.CheckBox1Click(Sender: TObject);
begin
gvPlanos.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmReporteESPlano.btnOKClick(Sender: TObject);
var i: integer;
sDescs : String;
begin
  GroupBox5.Visible := False;
  if CheckBox1.Checked = True then begin
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

procedure TfrmReporteESPlano.btnTodosClick(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos.Caption) = UT('Seleccionar Todos') then begin
        btnTodos.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvPlanos.RowCount - 1 do
        begin
                gvPlanos.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos.Caption := 'Seleccionar Todos';
        for i:= 0 to gvPlanos.RowCount - 1 do
        begin
                gvPlanos.Cell[1,i].AsBoolean := False;
        end;
  end;
end;

procedure TfrmReporteESPlano.FormCreate(Sender: TObject);
begin
  deFrom.Date := StartOfAMonth(YearOf(Now), MonthOf(Now));
  deTo.Date :=  EndOfTheMonth(Now);
  BindPlanos();

  BindGrid();
end;

procedure TfrmReporteESPlano.BindGrid();
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
                'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) AS Entradas, ' +
                'SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Salidas, ' +
                'SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) - SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Cantidad ' +
                'FROM tblPlano P ' +
                'INNER JOIN tblStock S ON P.PN_Id = S.PN_Id ' +
                'WHERE S.ST_Fecha >= ' + QuotedStr(deFrom.Text) + ' AND S.ST_Fecha <= ' + QuotedStr(deTo.Text  + ' 23:59:59.99');

      if Pos('*', txtPlano.Text) <> 0 then begin
        SQLStr := SQLStr + ' AND P.PN_Numero LIKE ''' + StringReplace(txtPlano.Text, '*', '%', [rfReplaceAll, rfIgnoreCase]) + '''';
      end else if txtPlano.Text <> 'Todos' then begin
        SQLStr := SQLStr + ' AND P.PN_Numero IN (''' + StringReplace(txtPlano.Text, ',', ''',''', [rfReplaceAll, rfIgnoreCase]) + ''')';
      end;

      SQLStr := SQLStr + ' GROUP BY P.PN_Descripcion, P.PN_Numero';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      GridView1.ClearRows();
      while not Qry.Eof do begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['PN_Descripcion']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['PN_Numero']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Entradas']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Salidas']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);

        entradas := entradas + StrToInt(VarToStr(Qry['Entradas']));
        salidas := salidas + StrToInt(VarToStr(Qry['Salidas']));
        existencia := existencia + StrToInt(VarToStr(Qry['Cantidad']));

        Qry.Next;
      end;

        GridView1.AddRow(1);
        GridView1.Cells[1,GridView1.RowCount -1] := 'Totales :';
        GridView1.Cells[2,GridView1.RowCount -1] := IntToStr(entradas);
        GridView1.Cells[3,GridView1.RowCount -1] := IntToStr(salidas);
        GridView1.Cells[4,GridView1.RowCount -1] := IntToStr(existencia);
    end
    finally
      CloseConns(Qry, Conn);
    end;
end;

procedure TfrmReporteESPlano.BindPlanos();
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

procedure TfrmReporteESPlano.Button3Click(Sender: TObject);
begin
  if GroupBox5.Visible = True then
  begin
          GroupBox5.Visible := False;
  end
  else begin
      GroupBox5.Visible := True;
      if txtPlano.Text = 'Todos' then
      begin
              CheckBox1.Checked := True;
              gvPlanos.Enabled := False;
              btnTodos.Enabled := False;
      end
      else
      begin
              CheckBox1.Checked := False;
              gvPlanos.Enabled := True;
              btnTodos.Enabled := True;
      end;

      GroupBox5.Top := txtPlano.Top + txtPlano.Height + 5;
      GroupBox5.Left := txtPlano.Left + 5;
  end;
end;

procedure TfrmReporteESPlano.txtPlanoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      BindGrid();
  end
end;

procedure TfrmReporteESPlano.txtPlanoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmReporteESPlano.btnBuscarClick(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmReporteESPlano.Button1Click(Sender: TObject);
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

procedure TfrmReporteESPlano.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Entradas Salidas Por No.Plano';

      Sheet.Cells[1,1] := 'Entradas vs Salidas Por No.Plano';
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

procedure TfrmReporteESPlano.Copiar1Click(Sender: TObject);
begin
  Button1Click(nil);
end;

end.
