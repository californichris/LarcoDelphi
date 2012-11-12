unit CatalogoCategories;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd,ComObj, Larco_Functions;

type
  TfrmCatalogoCategories = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    txtNombre: TEdit;
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
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindGrid();
    procedure FormCreate(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
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
  frmCatalogoCategories: TfrmCatalogoCategories;
  sPermits : String;
implementation

{$R *.dfm}

procedure TfrmCatalogoCategories.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmCatalogoCategories.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblCategories ORDER BY Category_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView1.ClearRows;
        While not Qry.Eof do
        begin
            GridView1.AddRow(1);
            GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Category_ID']);
            GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Category_Name']);
            Qry.Next;
        end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;


procedure TfrmCatalogoCategories.FormCreate(Sender: TObject);
begin
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
  BindGrid;
end;

procedure TfrmCatalogoCategories.NuevoClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    Conn := nil;
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        if txtNombre.Text = '' then
        begin
            MessageDlg('Por favor escriba un Nombre de Categoria.', mtInformation,[mbOk], 0);
            Exit;
        end;

        SQLStr := 'Categories 0,' + QuotedStr(txtNombre.Text);

        Conn.Execute(SQLStr);
    end
    finally
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    BindGrid();
    txtNombre.Text := '';
    txtNombre.SetFocus;
end;

procedure TfrmCatalogoCategories.EditarClick(Sender: TObject);
begin
  txtNombre.Text := GridView1.Cells[1,GridView1.SelectedRow];

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

procedure TfrmCatalogoCategories.BorrarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if MessageDlg('Estas seguro que quieres borrar la categoria : ' +
                  GridView1.Cells[1,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
            Exit;
    end;

   Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;

      SQLStr := 'Categories 2,' + QuotedStr('') + ',' + GridView1.Cells[0,GridView1.SelectedRow];

      Conn.Execute(SQLStr);
    end
    finally
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    BindGrid();
end;

procedure TfrmCatalogoCategories.btnCancelarClick(Sender: TObject);
begin
  txtNombre.Text := '';

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

procedure TfrmCatalogoCategories.btnAceptarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if txtNombre.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre de pantalla.', mtInformation,[mbOk], 0);
        Exit;
    end;

    Conn := nil;
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'Categories 1,' + QuotedStr(txtNombre.Text) + ',' + GridView1.Cells[0,GridView1.SelectedRow];

        Conn.Execute(SQLStr);
    end
    finally
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    BindGrid();
    btnCancelarClick(nil);
end;

procedure TfrmCatalogoCategories.GridView1SelectCell(Sender: TObject; ACol,
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

procedure TfrmCatalogoCategories.Copiarnombre1Click(Sender: TObject);
begin
  Clipboard.AsText := GridView1.Cells[1,GridView1.SelectedRow];
end;

procedure TfrmCatalogoCategories.Refrescar1Click(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmCatalogoCategories.Editar1Click(Sender: TObject);
begin
  EditarClick(nil);
end;

procedure TfrmCatalogoCategories.Borrar1Click(Sender: TObject);
begin
  BorrarClick(nil);
end;

end.
