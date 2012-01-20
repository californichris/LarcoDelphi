unit Grupos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView,ADODB,DB,IniFiles,All_Functions, Menus;

type
  TfrmGrupos = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    txtNombre: TEdit;
    Button1: TButton;
    btnCancelar: TButton;
    PopupMenu1: TPopupMenu;
    Borrar1: TMenuItem;
    Editar1: TMenuItem;
    gvGrupos: TGridView;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindGrupos();
    procedure Borrar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    function GrupoExists(Producto: String):Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGrupos: TfrmGrupos;
  gbEditar: Boolean;
  giRow : Integer;

implementation

uses Main;

{$R *.dfm}

procedure TfrmGrupos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmGrupos.FormCreate(Sender: TObject);
begin
    gbEditar := False;
    BindGrupos();
end;

procedure TfrmGrupos.BindGrupos();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT * FROM tblGrupos';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvGrupos.ClearRows;
    While not Qry.Eof do
    Begin
        gvGrupos.AddRow(1);
        gvGrupos.Cells[0,gvGrupos.RowCount -1] := VarToStr(Qry['Id']);
        gvGrupos.Cells[1,gvGrupos.RowCount -1] := VarToStr(Qry['Nombre']);
        Qry.Next;
    End;


    Qry.Close;
    Conn.Close;
end;

procedure TfrmGrupos.Borrar1Click(Sender: TObject);
var sId,sProducto : string;
Conn : TADOConnection;
SQLStr : String;
begin
  sProducto := gvGrupos.Cells[1,gvGrupos.SelectedRow];
  sId := gvGrupos.Cells[0,gvGrupos.SelectedRow];

  if MessageDlg('Estas seguro que quieres borrar el Grupo ' + sProducto + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  SQLStr := 'DELETE FROM tblGrupos WHERE Id = ' + sId;

  Conn.Execute(SQLStr);
  Conn.Close;

  BindGrupos();
end;

procedure TfrmGrupos.Editar1Click(Sender: TObject);
begin
        gbEditar := True;
        giRow := gvGrupos.SelectedRow;
        btnCancelar.Enabled := True;
        txtNombre.Text := gvGrupos.Cells[1,gvGrupos.SelectedRow];
        txtNombre.SetFocus;
end;

procedure TfrmGrupos.btnCancelarClick(Sender: TObject);
begin
        btnCancelar.Enabled := False;
        gbEditar := False;
        txtNombre.Text := '';
        txtNombre.SetFocus;
end;

procedure TfrmGrupos.Button1Click(Sender: TObject);
var Conn : TADOConnection;
SQLStr : String;
begin
    If txtNombre.Text = '' then
      begin
        MessageDlg('Por favor escriba un nombre de Producto.', mtInformation,[mbOk], 0);
        Exit;
      end;

    if GrupoExists(txtNombre.Text) then
      begin
        MessageDlg('Ya existe un producto con este nombre.', mtInformation,[mbOk], 0);
        Exit;
      end;

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    if gbEditar Then
        SQLStr := 'UPDATE tblGrupos SET Nombre = ' +  QuotedStr(txtNombre.Text) +
                  ' WHERE Nombre = ' +  QuotedStr(txtNombre.Text)
    else
        SQLStr := 'INSERT INTO tblGrupos(Nombre) ' +
                  'VALUES(' + QuotedStr(txtNombre.Text) + ')';

    Conn.Execute(SQLStr);
    Conn.Close;

    BindGrupos();
    txtNombre.Text := '';
    txtNombre.SetFocus;
    gbEditar := False;
    btnCancelar.Enabled := False;
end;

function TfrmGrupos.GrupoExists(Producto: String):Boolean;
var i:integer;
begin
        GrupoExists := False;
        for i:=0 to gvGrupos.RowCount -1 do
          begin
                if gbEditar and (i <> giRow ) then
                  if Producto = gvGrupos.Cells[1,i] then
                    GrupoExists := True;
          end;
end;

end.
