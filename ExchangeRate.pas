unit ExchangeRate;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd, CellEditors,Chris_Functions,Larco_Functions;

type
  TfrmExchangeRate = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label3: TLabel;
    txtAmount: TEdit;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    GridView1: TGridView;
    PopupMenu1: TPopupMenu;
    Copiarnombre1: TMenuItem;
    N1: TMenuItem;
    Refrescar1: TMenuItem;
    N2: TMenuItem;
    Editar1: TMenuItem;
    Borrar1: TMenuItem;
    deFecha: TDateEditor;
    Buscar: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindGrid();
    procedure FormCreate(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure txtAmountKeyPress(Sender: TObject; var Key: Char);
    procedure EditarClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure Copiarnombre1Click(Sender: TObject);
    procedure Refrescar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Borrar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmExchangeRate: TfrmExchangeRate;
  sPermits : String;
implementation

{$R *.dfm}

procedure TfrmExchangeRate.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmExchangeRate.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
//i: Integer;
begin

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT * FROM tblExchangeRate ORDER BY Rate_Date';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

{    for i:=0 to Qry.Fields.Count - 1 do
    begin
        ShowMessage(Qry.Fields.Fields[i].DisplayName);
    end;
}
    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Rate_ID']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Rate_Date']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Rate_Amount']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;


procedure TfrmExchangeRate.FormCreate(Sender: TObject);
begin
        BindGrid();
end;

procedure TfrmExchangeRate.NuevoClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    if txtAmount.Text = '' then
    begin
        MessageDlg('Por favor captura el tipo de cambio.', mtInformation,[mbOk], 0);
        Exit;
    end;

    if StrToFloat(txtAmount.Text) <= 0.00 then
    begin
        MessageDlg('Tipo de cambio tiene que ser mayor de 0.', mtInformation,[mbOk], 0);
        Exit;
    end;


    SQLStr := 'ExchangeRate 0,' + QuotedStr(deFecha.Text) + ',' + txtAmount.Text;

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();

    deFecha.Text := DateToStr(Now);
    txtAmount.Text := '';
    deFecha.SetFocus;

  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmExchangeRate.txtAmountKeyPress(Sender: TObject;
  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar(txtAmount.Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;

end;

procedure TfrmExchangeRate.EditarClick(Sender: TObject);
begin
deFecha.Text := GridView1.Cells[1,GridView1.SelectedRow];
txtAmount.Text := GridView1.Cells[2,GridView1.SelectedRow];

Nuevo.Visible := False;
Editar.Visible := False;
Borrar.Visible := False;

btnAceptar.Visible := True;
btnCancelar.Visible := True;

btnAceptar.Top := Nuevo.Top;
btnAceptar.Left := Nuevo.Left;

btnCancelar.Top := Editar.Top;
btnCancelar.Left := Editar.Left;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmExchangeRate.GridView1SelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
Editar.Enabled := False;
Borrar.Enabled := False;

if GridView1.Cells[ACol,ARow] <> '' then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
  end;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmExchangeRate.BorrarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if MessageDlg('Estas seguro que quieres borrar la pantalla : ' +
                  GridView1.Cells[1,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
            Exit;
    end;

    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'ExchangeRate 2,' + QuotedStr('') + ',0.00,' + GridView1.Cells[0,GridView1.SelectedRow];

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();

end;

procedure TfrmExchangeRate.btnCancelarClick(Sender: TObject);
begin
deFecha.Text := DateToStr(Now);
txtAmount.Text := '';

Nuevo.Visible := True;
Editar.Visible := True;
Borrar.Visible := True;

btnAceptar.Visible := False;
btnCancelar.Visible := False;

btnAceptar.Top := Nuevo.Top;
btnAceptar.Left := Nuevo.Left;

btnCancelar.Top := Editar.Top;
btnCancelar.Left := Editar.Left;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmExchangeRate.btnAceptarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if txtAmount.Text = '' then
    begin
        MessageDlg('Por favor captura el tipo de cambio.', mtInformation,[mbOk], 0);
        Exit;
    end;

    if StrToFloat(txtAmount.Text) <= 0.00 then
    begin
        MessageDlg('Tipo de cambio tiene que ser mayor de 0.', mtInformation,[mbOk], 0);
        Exit;
    end;

    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'ExchangeRate 1,' + QuotedStr(deFecha.Text) + ',' + txtAmount.Text +
              ',' + GridView1.Cells[0,GridView1.SelectedRow];

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();

    btnCancelarClick(nil);
end;

procedure TfrmExchangeRate.Copiarnombre1Click(Sender: TObject);
begin
Clipboard.AsText := GridView1.Cells[1,GridView1.SelectedRow];
end;

procedure TfrmExchangeRate.Refrescar1Click(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmExchangeRate.Editar1Click(Sender: TObject);
begin
EditarClick(nil);
end;

procedure TfrmExchangeRate.Borrar1Click(Sender: TObject);
begin
BorrarClick(nil);
end;

end.
