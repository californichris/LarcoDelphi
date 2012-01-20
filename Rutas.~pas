unit Rutas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView,Routing,ADODB,DB,IniFiles,All_Functions,Larco_Functions;

type
  TfrmRutas = class(TForm)
    gbButtons: TGroupBox;
    gvRutas: TGridView;
    gvValidos: TGridView;
    Label4: TLabel;
    Label1: TLabel;
    Nuevo: TButton;
    Borrar: TButton;
    Editar: TButton;
    procedure NuevoClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindRoutes();
    procedure BindValidos(Task: String);
    procedure BorrarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure gvRutasSelectCell(Sender: TObject; ACol, ARow: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRutas: TfrmRutas;
  sPermits : String;
implementation

uses Main;

{$R *.dfm}

procedure TfrmRutas.NuevoClick(Sender: TObject);
begin
Borrar.Enabled := False;
Editar.Enabled := False;
Application.CreateForm(TfrmRouting,frmRouting);
frmRouting.lblRuta.Caption := '';
frmRouting.ShowModal;
BindRoutes();
end;

procedure TfrmRutas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmRutas.FormCreate(Sender: TObject);
begin
    BindRoutes();
    sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmRutas.BindRoutes();
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

    SQLStr := 'SELECT Rou_Code,T.Nombre As FromName,T2.Nombre As ToName,Rou_From,Rou_To ' +
              'FROM tblrouting R ' +
              'INNER JOIN tblTareas T ON Rou_From = T.id ' +
              'INNER JOIN tblTareas T2 ON Rou_To = T2.id ' +
              'GROUP BY Rou_Code,T.Nombre,T2.Nombre,Rou_From,Rou_To ' +
              'ORDER BY Rou_from ';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvRutas.ClearRows;
    While not Qry.Eof do
    Begin
        gvRutas.AddRow(1);
        gvRutas.Cells[0,gvRutas.RowCount -1] := VarToStr(Qry['Rou_Code']);
        gvRutas.Cells[1,gvRutas.RowCount -1] := VarToStr(Qry['FromName']);
        gvRutas.Cells[2,gvRutas.RowCount -1] := VarToStr(Qry['ToName']);
        gvRutas.Cells[3,gvRutas.RowCount -1] := VarToStr(Qry['Rou_From']);
        gvRutas.Cells[4,gvRutas.RowCount -1] := VarToStr(Qry['Rou_To']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
    gvValidos.ClearRows;
end;

procedure TfrmRutas.BindValidos(Task: String);
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

    SQLStr := 'SELECT CASE WHEN Nombre = ''*'' THEN ''Todos'' ELSE Nombre END AS Nombre ' +
              'FROM tblrouting WHERE Rou_From = ' + Task;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvValidos.ClearRows;
    While not Qry.Eof do
    Begin
        gvValidos.AddRow(1);
        gvValidos.Cells[0,gvValidos.RowCount -1] := VarToStr(Qry['Nombre']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmRutas.BorrarClick(Sender: TObject);
var Conn : TADOConnection;
SQLStr : String;
begin
  if MessageDlg('Estas seguro que quieres borrar la Ruta ' + gvRutas.Cells[0,gvRutas.SelectedRow] +
                ' que va de la tarea ' + gvRutas.Cells[1,gvRutas.SelectedRow] +
                ' a la tarea ' + gvRutas.Cells[2,gvRutas.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

            //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'DELETE FROM tblrouting WHERE Rou_From = ' + gvRutas.Cells[3,gvRutas.SelectedRow];

    Conn.Execute(SQLStr);
    Conn.Close;

    BindRoutes();
end;

procedure TfrmRutas.EditarClick(Sender: TObject);
begin
Borrar.Enabled := False;
Editar.Enabled := False;
Application.CreateForm(TfrmRouting,frmRouting);
frmRouting.lblRuta.Caption := gvRutas.Cells[3,gvRutas.SelectedRow];
frmRouting.lblTo.Caption := gvRutas.Cells[4,gvRutas.SelectedRow];
frmRouting.lblCode.Caption := gvRutas.Cells[0,gvRutas.SelectedRow];
frmRouting.ShowModal;
BindRoutes();
end;

procedure TfrmRutas.gvRutasSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
Borrar.Enabled := True;
Editar.Enabled := True;
BindValidos(gvRutas.Cells[3,gvRutas.SelectedRow]);
end;

end.
