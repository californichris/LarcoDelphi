unit Scrap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions,
  ImpresionOrden,Larco_Functions;
type
  TfrmScrap = class(TForm)
    GridView1: TGridView;
    btnRefresh: TButton;
    btnCerrar: TButton;
    btnPrint: TButton;
    btnEditar: TButton;
    btnDetalle: TButton;
    btnBorrar: TButton;
    function FormIsRunning(FormName: String):Boolean;
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCerrarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure btnDetalleClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnEditarClick(Sender: TObject);
    procedure btnBorrarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScrap: TfrmScrap;
  gsYear : String;

implementation

uses Main, Ventas;

{$R *.dfm}

procedure TfrmScrap.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmScrap.btnCerrarClick(Sender: TObject);
begin
Close;
end;

procedure TfrmScrap.FormCreate(Sender: TObject);
begin
    gsYear := RightStr(getFormYear(frmMain.sConnString,Self.Name),2) + '-';

    BindGrid();
end;

procedure TfrmScrap.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'Select * from tblScrap S ' +
              'INNER JOIN tblOrdenes O ON S.SCR_NewItem = O.ITE_Nombre where S.SCR_Activo = 0';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    if Qry.RecordCount <= 0 then
        begin
                Exit;
        end;

    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['SCR_NewItem']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Requerida']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['SCR_Repro']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cell[5,GridView1.RowCount -1].AsBoolean := StrToBool(VarToStr(Qry['SCR_Impreso']));
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['SCR_Motivo']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['SCR_Tarea']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['SCR_EmpleadoRes']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;



procedure TfrmScrap.btnRefreshClick(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmScrap.btnDetalleClick(Sender: TObject);
begin
MessageDlg('Motivo : ' + GridView1.Cells[6,GridView1.SelectedRow] + #13 +
           'Area Reponsable : ' + GridView1.Cells[7,GridView1.SelectedRow] + #13 +
           'Empleado Responsable : ' + GridView1.Cells[8,GridView1.SelectedRow] , mtInformation,[mbOk], 0);
end;

procedure TfrmScrap.btnPrintClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    if GridView1.Cells[1,GridView1.SelectedRow] = '' then
        exit;

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'Update tblScrap SET SCR_Impreso = 1 where ITE_Nombre = ' + QuotedStr(GridView1.Cells[0,GridView1.SelectedRow]);

    conn.Execute(SQLStr);
    GridView1.Cell[5,GridView1.SelectedRow].AsBoolean := True;

    SQLStr := 'Select * from tblOrdenes where ITE_Nombre = ' + QuotedStr(GridView1.Cells[1,GridView1.SelectedRow]);

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    if Qry.RecordCount <= 0 then
                Exit;

    Application.Initialize;
    Application.CreateForm(TqrImpresionOrden,qrImpresionOrden);

    qrImpresionOrden.QROrden.Caption := VarToStr(Qry['ITE_Nombre']);
    qrImpresionOrden.QROrden2.Caption := VarToStr(Qry['ITE_Nombre']);

    qrImpresionOrden.QRCode1.Caption := '*' + VarToStr(Qry['ITE_Nombre']) + '*';
    qrImpresionOrden.QRCode2.Caption := '*' + VarToStr(Qry['ITE_Nombre']) + '*';
    qrImpresionOrden.QRSemana.Caption := '';
    qrImpresionOrden.QRNumero.Caption := VarToStr(Qry['Numero']);
    qrImpresionOrden.QRTerminal.Caption := VarToStr(Qry['Terminal']);
    qrImpresionOrden.QRRecibido.Caption := VarToStr(Qry['Recibido']);
    qrImpresionOrden.QREntrega.Caption := VarToStr(Qry['Interna']);
    qrImpresionOrden.QREntrega2.Caption := VarToStr(Qry['Interna']);
    qrImpresionOrden.QRNombre.Caption := VarToStr(Qry['Nombre']);
    qrImpresionOrden.QRFirma.Caption := '';
    qrImpresionOrden.QRObs.Caption := VarToStr(Qry['Observaciones']);
    qrImpresionOrden.QRDesc.Caption := VarToStr(Qry['Producto']);
    qrImpresionOrden.QRCompra.Caption := VarToStr(Qry['OrdenCompra']);
    qrImpresionOrden.QRProceso.Caption := VarToStr(Qry['TipoProceso']);
    qrImpresionOrden.QRCantidad.Caption := VarToStr(Qry['Ordenada']);

    qrImpresionOrden.QRMsg.Caption := 'Forma: Larco-015' + #13 +
                                      'Nivel de Revisión: D' + #13 +
                                      'Retención: 1 año+uso';
    //qrImpresionOrden.Print;
    qrImpresionOrden.Preview;
    qrImpresionOrden.Free;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmScrap.btnEditarClick(Sender: TObject);
begin
    if GridView1.Cells[1,GridView1.SelectedRow] = '' then
        exit;

if FormIsRunning('frmVentas') Then
  begin
        setActiveWindow(frmVentas.Handle);
        frmVentas.WindowState := wsNormal;
        frmVentas.BuscarClick(nil);
        frmVentas.txtOrden.Text := RightStr(GridView1.Cells[1,GridView1.SelectedRow],Length(GridView1.Cells[1,GridView1.SelectedRow]) - 3);
        frmVentas.btnAceptarClick(nil);
        frmVentas.txtOrden.SetFocus;
  end
else
  begin
        Application.CreateForm(TfrmVentas,frmVentas);
        frmVentas.Show;
        frmVentas.BuscarClick(nil);
        frmVentas.txtOrden.Text := RightStr(GridView1.Cells[1,GridView1.SelectedRow],Length(GridView1.Cells[1,GridView1.SelectedRow]) - 3);
        frmVentas.btnAceptarClick(nil);
        frmVentas.txtOrden.SetFocus;
  end;

end;

function TfrmScrap.FormIsRunning(FormName: String):Boolean;
var i:Integer;
begin
  Result := False;

  for  i := 0 to Screen.FormCount - 1 do
  begin
        if Screen.Forms[i].Name = FormName Then
          begin
                Result:= True;
                Break;
          end;
  end;

end;

procedure TfrmScrap.btnBorrarClick(Sender: TObject);
begin
    if GridView1.Cells[1,GridView1.SelectedRow] = '' then
        exit;

  if MessageDlg('Estas seguro que quieres borrar esta orden ' +
                GridView1.Cells[1,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  begin
      Exit;
  end;



end;

end.
