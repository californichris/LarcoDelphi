unit CatalogoPermisos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd,ComObj;

type
  TfrmCatalogoPermisos = class(TForm)
    GroupBox1: TGroupBox;
    GridView1: TGridView;
    GridView2: TGridView;
    btnCerrar: TButton;
    btnActualizar: TButton;
    btnGrabar: TButton;
    btnNotAll: TButton;
    btnAll: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BindGrupos();
    procedure BindPantallas();
    procedure btnActualizarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnGrabarClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnCerrarClick(Sender: TObject);
    procedure btnAllClick(Sender: TObject);
    procedure CheckPermits(value: Boolean);
    procedure btnNotAllClick(Sender: TObject);
    function getBooleanValue(value: String):Boolean;
    function getStringValue(value: Boolean):String;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCatalogoPermisos: TfrmCatalogoPermisos;

implementation

{$R *.dfm}

procedure TfrmCatalogoPermisos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCatalogoPermisos.BindGrupos();
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

        SQLStr := 'SELECT * FROM tblGroups ORDER BY Group_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView1.ClearRows;
        While not Qry.Eof do
        begin
            GridView1.AddRow(1);
            GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Group_ID']);
            GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Group_Name']);
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

procedure TfrmCatalogoPermisos.BindPantallas();
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

        SQLStr := 'SELECT * FROM tblScreens WHERE SCR_FormName <> ' + QuotedStr('space') +
                  ' ORDER BY SCR_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView2.ClearRows;
        While not Qry.Eof do
        Begin
            GridView2.AddRow(1);
            GridView2.Cells[0,GridView2.RowCount -1] := VarToStr(Qry['SCR_ID']);
            GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry['SCR_Name']);
            Qry.Next;
        End;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry.Close;
    Conn.Close;
end;


procedure TfrmCatalogoPermisos.btnActualizarClick(Sender: TObject);
begin
  BindGrupos();
  BindPantallas();
  GridView1SelectCell(nil,0,GridView1.SelectedRow);
end;

procedure TfrmCatalogoPermisos.FormCreate(Sender: TObject);
begin
  btnActualizarClick(nil);
end;

procedure TfrmCatalogoPermisos.btnGrabarClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
i : Integer;
begin
    if GridView1.RowCount <= 0 then
        Exit;

    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'DELETE FROM tblGroup_Screens WHERE SCR_ID <> 24 AND Group_Id = ' +
                  GridView1.Cells[0,GridView1.SelectedRow];

        conn.Execute(SQLStr);

        for i := 0 to GridView2.RowCount - 1 do
        begin
            if GridView2.Cell[2,i].AsBoolean = True then begin
                SQLStr := 'INSERT INTO tblGroup_Screens(Group_Id, SCR_ID, Nuevo,Editar,Borrar,Buscar) ' +
                          'VALUES(' + GridView1.Cells[0,GridView1.SelectedRow] + ',' +
                          GridView2.Cells[0,i] + ',' + getStringValue(GridView2.Cell[3,i].AsBoolean) +
                          ',' + getStringValue(GridView2.Cell[4,i].AsBoolean) + ',' +
                          getStringValue(GridView2.Cell[5,i].AsBoolean) + ',' +
                          getStringValue(GridView2.Cell[6,i].AsBoolean) + ')';

                conn.Execute(SQLStr);
            end;
        end;

        GridView1SelectCell(nil,0,GridView1.SelectedRow);
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;
end;

procedure TfrmCatalogoPermisos.GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
i : Integer;
begin
    Conn := nil;
    Qry := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblGroup_Screens WHERE Group_Id = ' +
                  GridView1.Cells[0,GridView1.SelectedRow] +
                  ' ORDER BY SCR_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;


        CheckPermits(False);

        While not Qry.Eof do
        Begin
            for i := 0 to GridView2.RowCount - 1 do
            begin
                    if  GridView2.Cells[0,i] = VarToStr(Qry['SCR_ID']) then
                    begin
                            GridView2.Cell[2,i].AsBoolean := True;
                            GridView2.Cell[3,i].AsBoolean := getBooleanValue(VarToStr(Qry['Nuevo']));
                            GridView2.Cell[4,i].AsBoolean := getBooleanValue(VarToStr(Qry['Editar']));
                            GridView2.Cell[5,i].AsBoolean := getBooleanValue(VarToStr(Qry['Borrar']));
                            GridView2.Cell[6,i].AsBoolean := getBooleanValue(VarToStr(Qry['Buscar']));
                            break;
                    end;
            end;

            Qry.Next;
        End;
        Application.ProcessMessages;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry.Close;
    Conn.Close;

end;

procedure TfrmCatalogoPermisos.btnCerrarClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmCatalogoPermisos.btnAllClick(Sender: TObject);
begin
  CheckPermits(True);
end;

procedure TfrmCatalogoPermisos.CheckPermits(value: Boolean);
var i : Integer;
begin
  for i := 0 to GridView2.RowCount - 1 do
  begin
     GridView2.Cell[2,i].AsBoolean := value;
     GridView2.Cell[3,i].AsBoolean := value;
     GridView2.Cell[4,i].AsBoolean := value;
     GridView2.Cell[5,i].AsBoolean := value;
     GridView2.Cell[6,i].AsBoolean := value;
  end;
end;
procedure TfrmCatalogoPermisos.btnNotAllClick(Sender: TObject);
begin
  CheckPermits(False);
end;

function TfrmCatalogoPermisos.getBooleanValue(value: String):Boolean;
begin
        Result := True;
        if value = '0' then Result := False;
end;

function TfrmCatalogoPermisos.getStringValue(value: Boolean):String;
begin
        Result := '0';
        if value = True then Result := '1';
end;


end.
