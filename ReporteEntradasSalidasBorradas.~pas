unit ReporteEntradasSalidasBorradas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math;

type
  TfrmEntradasSalidasBorradas = class(TForm)
    GroupBox1: TGroupBox;
    GridView1: TGridView;
    Label1: TLabel;
    deFrom: TDateEditor;
    deTo: TDateEditor;
    Label2: TLabel;
    Button1: TButton;
    lblAnio: TLabel;
    lblCount: TLabel;
    GridView2: TGridView;
    lblCount2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEntradasSalidasBorradas: TfrmEntradasSalidasBorradas;
  Conn : TADOConnection;
  Qry : TADOQuery;

implementation

uses Main;

{$R *.dfm}

procedure TfrmEntradasSalidasBorradas.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);

  deFrom.Date := Now;
  deTo.Date := Now;

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  BindGrid();

end;

procedure TfrmEntradasSalidasBorradas.BindGrid();
begin
  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT E.*,P.PROV_Nombre, ' +
                  'CASE WHEN U.USE_Name IS NULL THEN '''' ELSE U.USE_Name END AS USE_Name ' +
                  'FROM tblEntradas_History E ' +
                  'LEFT OUTER JOIN tblProvedores P ON E.PROV_ID = P.PROV_ID ' +
                  'LEFT OUTER JOIN tblUsers U ON E.USE_ID = U.USE_Login ' +
                  'WHERE YEAR(ENT_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' AND (ENTH_Fecha >= ' +
                  QuotedStr(deFrom.Text) + ' AND ENTH_Fecha <= ' +
                  QuotedStr(deTo.Text + ' 23:59:59.999') + ') ORDER BY ENT_ID';
  Qry.Open;

  GridView1.ClearRows;
  While not Qry.Eof do
  Begin
      GridView1.AddRow(1);
      GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['ENT_ID']);
      GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['ENT_Pedimento']);
      GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['ENT_ClavePedimento']);
      GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['ENT_Fecha']);
      GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['ENT_PaisOrigen']);
      GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['ENT_Nacional']);
      GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['ENT_TipoImp']);
      GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['ENT_Factura']);
      GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['ENT_OrdenCompra']);
      GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['PROV_Nombre']);
      GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(Qry['USE_Name']);
      GridView1.Cells[11,GridView1.RowCount -1] := VarToStr(Qry['ENTH_Fecha']);
      Qry.Next;
  End;

  lblCount.Caption := 'Total de Entradas Borradas: ' + IntToStr(GridView1.RowCount);

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT S.*, ' +
                  'CASE WHEN U2.USE_Name IS NULL THEN '''' ELSE U2.USE_Name END AS SAL_Solicitado, ' +
                  'CASE WHEN U.USE_Name IS NULL THEN '''' ELSE U.USE_Name END AS USE_Name ' +
                  'FROM tblSalidas_History S ' +
                  'LEFT OUTER JOIN tblUsers U2 ON S.SAL_Solicitado = U2.USE_Login ' +
                  'LEFT OUTER JOIN tblUsers U ON S.USE_ID = U.USE_Login ' +
                  'WHERE YEAR(SAL_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' AND (SALH_Fecha >= ' +
                  QuotedStr(deFrom.Text) + ' AND SALH_Fecha <= ' +
                  QuotedStr(deTo.Text + ' 23:59:59.999') + ') ORDER BY SAL_ID';
  Qry.Open;

  GridView2.ClearRows;
  While not Qry.Eof do
  Begin
      GridView2.AddRow(1);
      GridView2.Cells[0,GridView2.RowCount -1] := VarToStr(Qry['SAL_ID']);
      GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry['SAL_Fecha']);
      GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['SAL_Orden']);
      GridView2.Cells[3,GridView2.RowCount -1] := VarToStr(Qry['SAL_Solicitado']);
      GridView2.Cells[4,GridView2.RowCount -1] := VarToStr(Qry['USE_Name']);
      GridView2.Cells[5,GridView2.RowCount -1] := VarToStr(Qry['SALH_Fecha']);
      Qry.Next;
  End;

  lblCount2.Caption := 'Total de Salidas Borradas : ' + IntToStr(GridView2.RowCount);


end;

procedure TfrmEntradasSalidasBorradas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmEntradasSalidasBorradas.Button1Click(Sender: TObject);
begin
  BindGrid();
end;

end.
