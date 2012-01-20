unit Screens;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd;

type
  TfrmScreens = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtNombre: TEdit;
    txtForma: TEdit;
    txtDescripcion: TEdit;
    btnAgregar: TButton;
    GridView1: TGridView;
    btnEditar: TButton;
    btnBorrar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    PopupMenu1: TPopupMenu;
    Copiarnombre1: TMenuItem;
    Refrescar1: TMenuItem;
    Editar1: TMenuItem;
    Borrar1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure BindGrid();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAgregarClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnBorrarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure btnEditarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure Refrescar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Borrar1Click(Sender: TObject);
    procedure Copiarnombre1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScreens: TfrmScreens;

implementation

{$R *.dfm}

procedure TfrmScreens.FormCreate(Sender: TObject);
begin
        BindGrid();
end;

procedure TfrmScreens.BindGrid();
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

    SQLStr := 'SELECT * FROM tblScreens ORDER BY SCR_ID';

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
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['SCR_ID']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['SCR_Name']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['SCR_FormName']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['SCR_Description']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;


procedure TfrmScreens.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmScreens.btnAgregarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    if txtNombre.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre de pantalla.', mtInformation,[mbOk], 0);
        Exit;
    end;

    if txtForma.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre de la forma.', mtInformation,[mbOk], 0);
        Exit;
    end;

    SQLStr := 'Screens 0,' + QuotedStr(txtNombre.Text) + ',' + QuotedStr(txtForma.Text) +
              ',' + QuotedStr(txtDescripcion.Text);

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();

    txtNombre.Text := '';
    txtForma.Text := '';
    txtDescripcion.Text := '';
    txtNombre.SetFocus;
end;

procedure TfrmScreens.GridView1SelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin

btnEditar.Enabled := False;
btnBorrar.Enabled := False;

if GridView1.Cells[ACol,ARow] <> '' then
  begin
        btnEditar.Enabled := True;
        btnBorrar.Enabled := True;
  end;
end;

procedure TfrmScreens.btnBorrarClick(Sender: TObject);
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

    SQLStr := 'Screens 2,' + QuotedStr('') + ',' + QuotedStr('') +
              ',' + QuotedStr('') + ',' + GridView1.Cells[0,GridView1.SelectedRow];

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();
end;

procedure TfrmScreens.btnAceptarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if txtNombre.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre de pantalla.', mtInformation,[mbOk], 0);
        Exit;
    end;

    if txtForma.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre de la forma.', mtInformation,[mbOk], 0);
        Exit;
    end;

    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'Screens 1,' + QuotedStr(txtNombre.Text) + ',' + QuotedStr(txtForma.Text) +
              ',' + QuotedStr(txtDescripcion.Text) + ',' + GridView1.Cells[0,GridView1.SelectedRow];

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrid();

    btnCancelarClick(nil);
end;

procedure TfrmScreens.btnEditarClick(Sender: TObject);
begin
txtNombre.Text := GridView1.Cells[1,GridView1.SelectedRow];
txtForma.Text := GridView1.Cells[2,GridView1.SelectedRow];
txtDescripcion.Text := GridView1.Cells[3,GridView1.SelectedRow];

btnAgregar.Visible := False;
btnEditar.Visible := False;
btnBorrar.Visible := False;

btnAceptar.Visible := True;
btnCancelar.Visible := True;

btnAceptar.Top := btnAgregar.Top;
btnAceptar.Left := btnAgregar.Left;

btnCancelar.Top := btnEditar.Top;
btnCancelar.Left := btnEditar.Left;
end;

procedure TfrmScreens.btnCancelarClick(Sender: TObject);
begin
txtNombre.Text := '';
txtForma.Text := '';
txtDescripcion.Text := '';

btnAgregar.Visible := True;
btnEditar.Visible := True;
btnBorrar.Visible := True;

btnAceptar.Visible := False;
btnCancelar.Visible := False;

btnAceptar.Top := btnAgregar.Top;
btnAceptar.Left := btnAceptar.Left;

btnCancelar.Top := btnEditar.Top;
btnCancelar.Left := btnEditar.Left;
end;

procedure TfrmScreens.Refrescar1Click(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmScreens.Editar1Click(Sender: TObject);
begin
btnEditarClick(nil);
end;

procedure TfrmScreens.Borrar1Click(Sender: TObject);
begin
btnBorrarClick(nil);
end;

procedure TfrmScreens.Copiarnombre1Click(Sender: TObject);
begin
Clipboard.AsText := GridView1.Cells[1,GridView1.SelectedRow];
end;

end.
