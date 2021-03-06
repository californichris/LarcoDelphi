unit ReportePromedioCump;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, CellEditors,ADODB,DB,IniFiles,ComObj,All_Functions,StrUtils,chris_Functions,
  TeEngine, Series, ExtCtrls, TeeProcs, Chart, ScrollView,LTCUtils,
  CustomGridViewControl, CustomGridView, GridView, Menus,Clipbrd;

type
  TfrmPromCumpli = class(TForm)
    GroupBox3: TGroupBox;
    lblLiberadas: TLabel;
    lblCLiberadas2: TLabel;
    lblCScrap2: TLabel;
    lblScrap: TLabel;
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label9: TLabel;
    Label3: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Button1: TButton;
    Detalle: TButton;
    Imprimir: TButton;
    cmbTipo: TComboBox;
    txtCliente: TEdit;
    Button2: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    PopupMenu1: TPopupMenu;
    Exportar1: TMenuItem;
    CopiarOrden1: TMenuItem;
    SaveDialog1: TSaveDialog;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindClientes();
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure btnTodosClick(Sender: TObject);
    procedure CopiarOrden1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPromCumpli: TfrmPromCumpli;
  giLiberadas:Integer;
  giScrap:Double;
implementation

uses Main;

{$R *.dfm}

procedure TfrmPromCumpli.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmPromCumpli.FormCreate(Sender: TObject);
begin
    deFrom.Date := Now;
    deTo.Date := Now;
    BindClientes();
    CheckBox1.Checked := False;
    btnOKClick(nil);

    Button1Click(nil);
end;

procedure TfrmPromCumpli.BindClientes();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';
    //Create Connection
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
        if (slClientes.IndexOf(VarToStr(Qry['Clave'])) = -1) then begin
                gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
        end;
        
        Qry.Next;
    End;
end;


procedure TfrmPromCumpli.Button1Click(Sender: TObject);
var  Qry : TADOQuery;
Conn : TADOConnection;
SQLStr : String;
begin
    //Create Connection
    giLiberadas := 0;
    giScrap := 0.00;
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'Cumplimiento ' + QuotedStr(deFrom.Text) + ',' +
               QuotedStr(deTo.Text  + ' 23:59:59.99') + ',';


    if txtCliente.Text <> 'Todos' then
    begin
        SQLStr := SQLStr + QuotedStr(txtCliente.Text) + ',';
    end
    else begin
        SQLStr := SQLStr + QuotedStr('') + ',';
    end;

    SQLStr := SQLStr + QuotedStr(cmbTipo.Text);

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;


    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := RightStr(VarToStr(Qry['Fecha']),10);;
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Liberadas']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['NoLiberadas']);
        giLiberadas := giLiberadas + StrToInt(GridView1.Cells[2,GridView1.RowCount -1]);
        GridView1.Cells[3,GridView1.RowCount -1] := IntToStr(StrToInt(GridView1.Cells[1,GridView1.RowCount -1])
                                                 +  StrToInt(GridView1.Cells[2,GridView1.RowCount -1]) );

        GridView1.Cells[4,GridView1.RowCount -1] := FormatFloat('########0.00',StrToInt(GridView1.Cells[2,GridView1.RowCount -1])
                                                 /  StrToInt(GridView1.Cells[3,GridView1.RowCount -1]) );
        giScrap := giScrap + StrToFloat(GridView1.Cells[4,GridView1.RowCount -1]);
        Qry.Next;
    End;

    lblLiberadas.Caption := '0.00';
    lblScrap.Caption := '0.00';
    if GridView1.RowCount > 0 then begin
        lblLiberadas.Caption :=  FormatFloat('########0.00',giLiberadas / GridView1.RowCount);
        lblScrap.Caption :=  FormatFloat('########0.00',giScrap / GridView1.RowCount);
    end;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmPromCumpli.Button2Click(Sender: TObject);
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

procedure TfrmPromCumpli.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmPromCumpli.btnOKClick(Sender: TObject);
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

procedure TfrmPromCumpli.Exportar1Click(Sender: TObject);
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

procedure TfrmPromCumpli.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Scrap';

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


procedure TfrmPromCumpli.btnTodosClick(Sender: TObject);
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

procedure TfrmPromCumpli.CopiarOrden1Click(Sender: TObject);
begin
      Clipboard.AsText := GridView1.Cells[0,GridView1.SelectedRow];
end;

end.
