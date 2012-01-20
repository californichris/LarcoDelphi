unit CatalogoTipoMaterial;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd,ComObj,Larco_Functions;

type
  TfrmTipoMaterial = class(TForm)
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
    ddlTipo: TComboBox;
    Label2: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindGrid();
    procedure FormCreate(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
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
  frmTipoMaterial: TfrmTipoMaterial;
  sPermits : String;
implementation

{$R *.dfm}

procedure TfrmTipoMaterial.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmTipoMaterial.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblTiposMaterial ORDER BY TIP_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView1.ClearRows;
        While not Qry.Eof do
        begin
            GridView1.AddRow(1);
            GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['TIP_ID']);
            GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['TIP_Tipo']);
            GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['TIP_Descripcion']);
            Qry.Next;
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry.Close;
    Conn.Close;
end;


procedure TfrmTipoMaterial.FormCreate(Sender: TObject);
begin
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
  BindGrid();
end;

procedure TfrmTipoMaterial.btnAceptarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if txtNombre.Text = '' then
    begin
        MessageDlg('Por favor escriba un Nombre.', mtInformation,[mbOk], 0);
        Exit;
    end;

    if ddlTipo.Text = '' then
    begin
        MessageDlg('Por favor seleccione un Tipo.', mtInformation,[mbOk], 0);
        Exit;
    end;

    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'TiposMaterial 1,' + QuotedStr(ddlTipo.Text) + ',' + QuotedStr(txtNombre.Text) + ',' +
                  GridView1.Cells[0,GridView1.SelectedRow];

        Conn.Execute(SQLStr);

        BindGrid();
        btnCancelarClick(nil);
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Conn.Close;
end;

procedure TfrmTipoMaterial.btnCancelarClick(Sender: TObject);
begin
  txtNombre.Text := '';
  ddlTipo.Text := '';

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

procedure TfrmTipoMaterial.NuevoClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        if txtNombre.Text = '' then
        begin
            MessageDlg('Por favor escriba un Nombre.', mtInformation,[mbOk], 0);
            Exit;
        end;

        if ddlTipo.Text = '' then
        begin
            MessageDlg('Por favor seleccione un Tipo.', mtInformation,[mbOk], 0);
            Exit;
        end;

        SQLStr := 'TiposMaterial 0,' + QuotedStr(ddlTipo.Text) + ',' + QuotedStr(txtNombre.Text);

        Conn.Execute(SQLStr);

        BindGrid();
        txtNombre.Text := '';
        txtNombre.SetFocus;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Conn.Close;
end;

procedure TfrmTipoMaterial.EditarClick(Sender: TObject);
begin
  txtNombre.Text := GridView1.Cells[2,GridView1.SelectedRow];
  ddlTipo.ItemIndex := ddlTipo.Items.IndexOf(GridView1.Cells[1,GridView1.SelectedRow]);
  application.ProcessMessages;


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

procedure TfrmTipoMaterial.BorrarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin

    if MessageDlg('Estas seguro que quieres borrar la Unidad : ' +
                  GridView1.Cells[2,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
            Exit;
    end;

    Conn := nil;
    try

      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := frmMain.sConnString;
      Conn.LoginPrompt := False;

      SQLStr := 'TiposMaterial 2,'''','''',' + GridView1.Cells[0,GridView1.SelectedRow];

      Conn.Execute(SQLStr);
      BindGrid();

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Conn.Close;
end;

procedure TfrmTipoMaterial.GridView1SelectCell(Sender: TObject; ACol,
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

procedure TfrmTipoMaterial.Copiarnombre1Click(Sender: TObject);
begin
  Clipboard.AsText := GridView1.Cells[2,GridView1.SelectedRow];
end;

procedure TfrmTipoMaterial.Refrescar1Click(Sender: TObject);
begin
  BindGrid();
end;

procedure TfrmTipoMaterial.Editar1Click(Sender: TObject);
begin
  EditarClick(nil);
end;

procedure TfrmTipoMaterial.Borrar1Click(Sender: TObject);
begin
  BorrarClick(nil);
end;

end.
