unit CatalogoScreens;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd,Larco_Functions;

type
  TfrmCatalogoScreens = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtNombre: TEdit;
    txtForma: TEdit;
    txtDescripcion: TEdit;
    Nuevo: TButton;
    GridView1: TGridView;
    Editar: TButton;
    Borrar: TButton;
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
    procedure NuevoClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure BorrarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
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
  frmCatalogoScreens: TfrmCatalogoScreens;
  sPermits : String;
implementation

{$R *.dfm}

procedure TfrmCatalogoScreens.FormCreate(Sender: TObject);
begin
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
  BindGrid();
end;

procedure TfrmCatalogoScreens.BindGrid();
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

    SQLStr := 'SELECT * FROM tblScreens WHERE SCR_FormName <> ' + QuotedStr('space') +
              ' ORDER BY SCR_ID';

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


procedure TfrmCatalogoScreens.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCatalogoScreens.NuevoClick(Sender: TObject);
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

procedure TfrmCatalogoScreens.GridView1SelectCell(Sender: TObject; ACol,
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

procedure TfrmCatalogoScreens.BorrarClick(Sender: TObject);
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

procedure TfrmCatalogoScreens.btnAceptarClick(Sender: TObject);
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

procedure TfrmCatalogoScreens.EditarClick(Sender: TObject);
begin
  txtNombre.Text := GridView1.Cells[1,GridView1.SelectedRow];
  txtForma.Text := GridView1.Cells[2,GridView1.SelectedRow];
  txtDescripcion.Text := GridView1.Cells[3,GridView1.SelectedRow];

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

procedure TfrmCatalogoScreens.btnCancelarClick(Sender: TObject);
begin
  txtNombre.Text := '';
  txtForma.Text := '';
  txtDescripcion.Text := '';

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

procedure TfrmCatalogoScreens.Refrescar1Click(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmCatalogoScreens.Editar1Click(Sender: TObject);
begin
  EditarClick(nil);
end;

procedure TfrmCatalogoScreens.Borrar1Click(Sender: TObject);
begin
  BorrarClick(nil);
end;

procedure TfrmCatalogoScreens.Copiarnombre1Click(Sender: TObject);
begin
  Clipboard.AsText := GridView1.Cells[1,GridView1.SelectedRow];
end;

end.
