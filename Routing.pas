unit Routing;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,ADODB,DB,IniFiles,All_Functions;

type
  TfrmRouting = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    cmbFrom: TComboBox;
    Label2: TLabel;
    cmbTo: TComboBox;
    Label3: TLabel;
    txtCodigo: TEdit;
    gvProductos: TGridView;
    Label4: TLabel;
    Button1: TButton;
    btnCancelar: TButton;
    lblRuta: TLabel;
    lblTo: TLabel;
    lblCode: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure BindProductos();
    procedure BindTasks();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCancelarClick(Sender: TObject);
    procedure BindValidos(Task: String);
    procedure FormActivate(Sender: TObject);
    function CheckForChanges():Boolean;
    procedure Button1Click(Sender: TObject);
    Procedure DeleteRoutes(Rou_From,Rou_To: String);
    procedure InsertRoutes(Rou_From,Rou_To,Rou_Code: String);
    function RouteExist(Rou_From,Rou_To: String):Boolean;
    function GetTaskId(TaskName: String):String;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRouting: TfrmRouting;
  gbFirst: Boolean;
implementation

uses Main;

{$R *.dfm}

procedure TfrmRouting.FormCreate(Sender: TObject);
begin
    BindTasks();
    BindProductos();
    gbFirst := False;
end;

procedure TfrmRouting.BindProductos();
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

    SQLStr := 'SELECT * FROM tblProductos';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvProductos.ClearRows;
    gvProductos.AddRow(1);
    gvProductos.Cells[1,gvProductos.RowCount -1] := 'Todos';
    While not Qry.Eof do
    Begin
        gvProductos.AddRow(1);
        //gvProductos.Cells[0,gvProductos.RowCount -1] := VarToStr(Qry['Id']);
        gvProductos.Cells[1,gvProductos.RowCount -1] := VarToStr(Qry['Nombre']);
        Qry.Next;
    End;


    Qry.Close;
    Conn.Close;
end;

procedure TfrmRouting.BindTasks();
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

    SQLStr := 'SELECT * FROM tblTareas';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    cmbFrom.Clear;
    cmbTo.Clear;
    While not Qry.Eof do
    Begin
        cmbFrom.Items.Add(VarToStr(Qry['Nombre']));
        cmbTo.Items.Add(VarToStr(Qry['Nombre']));
        Qry.Next;
    End;

    cmbFrom.Text := cmbFrom.Items[0];
    cmbTo.Text := cmbTo.Items[0];
    Qry.Close;
    Conn.Close;
end;


procedure TfrmRouting.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmRouting.btnCancelarClick(Sender: TObject);
begin
        Self.Close;
end;

procedure TfrmRouting.BindValidos(Task: String);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
i:Integer;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT CASE WHEN R.Nombre = ''*'' THEN ''Todos'' ELSE R.Nombre END AS Nombre,' +
              'Rou_Code,T.Nombre As FromName,T2.Nombre As ToName ' +
              'FROM tblrouting R ' +
              'INNER JOIN tblTareas T ON Rou_From = T.id ' +
              'INNER JOIN tblTareas T2 ON Rou_To = T2.id ' +
              'WHERE Rou_From = ' + Task +
              'GROUP BY R.Nombre,Rou_Code,T.Nombre,T2.Nombre ' +
              'ORDER BY R.Nombre ';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    cmbFrom.Text := VarToStr(Qry['FromName']);
    cmbTo.Text := VarToStr(Qry['ToName']);
    txtCodigo.Text := VarToStr(Qry['Rou_Code']);
    While not Qry.Eof do
    Begin
        for i:=0 to gvProductos.RowCount - 1 do
        begin
                if gvProductos.Cells[1,i] = VarToStr(Qry['Nombre']) then
                  begin
                     gvProductos.Cell[0,i].AsBoolean := True;
                     gvProductos.Cell[2,i].AsBoolean := True;
                     break;
                  end;
        end;

        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmRouting.FormActivate(Sender: TObject);
begin
    if gbFirst = True Then
      Exit;

    if lblRuta.Caption <> '' Then
    begin
        cmbFrom.Enabled := False;
        cmbTo.Enabled := False;

        BindValidos(lblRuta.Caption);
    end;
end;

function TfrmRouting.CheckForChanges():Boolean;
var i : Integer;
begin
  Result := False;
  for i:=0 to gvProductos.RowCount - 1 do
  begin
          if gvProductos.Cells[0,i] <> gvProductos.Cells[2,i] then
            begin
                Result := True;
                Break;
            end
          else
            begin
                if gvProductos.Cells[1,i] = 'Todos' Then
                   Break;
            end;
  end;

  if txtCodigo.Text <> lblCode.Caption Then
     Result := True;

end;

procedure TfrmRouting.Button1Click(Sender: TObject);
var sFrom,sTo :String;
begin

if txtcodigo.Text = '' Then
  begin
     ShowMessage('Por favor escriba un codigo para esta ruta.');
     Exit;
  end;

if lblRuta.Caption = '' Then //Si es agregar
  begin
        if cmbFrom.Text = cmbTo.Text Then
        begin
             ShowMessage('No puedes crear un ruta a la misma tarea.');
             Exit;
        end;

        if RouteExist(cmbFrom.Text,cmbTo.Text) then
        begin
             ShowMessage('Ya existe una ruta de la tarea ' + cmbFrom.Text + ' a la tarea ' + cmbTo.Text + '.');
             Exit;
        end;

        sFrom := '';
        sTo := '';
        sFrom := GetTaskId(cmbFrom.Text);
        sTo := GetTaskId(cmbTo.Text);

        if (sFrom = '') Or (sTo = '') then
        begin
             ShowMessage('Ocurrio un Error al leer informacion de la base de datos.' + #13 + 'Por favor intente nuevamente.' + #13 +
                         'Si el problema persiste por favor contacte al administrador');
             Exit;
        end;
        InsertRoutes(sFrom,sTo,txtCodigo.Text);
  end
Else  //Si es editar
  begin
      if CheckForChanges() Then
      begin
          //showMessage('Cambios');
          DeleteRoutes(lblRuta.Caption,lblTo.Caption);
          InsertRoutes(lblRuta.Caption,lblTo.Caption,txtCodigo.Text);
      end;
      //else
      //   showMessage('Nel');

  end;
end;

procedure TfrmRouting.DeleteRoutes(Rou_From,Rou_To: String);
var Conn : TADOConnection;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'DELETE FROM tblRouting WHERE Rou_From = ' + Rou_From + ' AND Rou_To = ' + Rou_To;

    Conn.Execute(SQLStr);
    Conn.Close;
end;

procedure TfrmRouting.InsertRoutes(Rou_From,Rou_To,Rou_Code: String);
var Conn : TADOConnection;
SQLStr : String;
i : Integer;
sNombre : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

        for i := 0 to gvProductos.RowCount - 1 do
        begin
                if gvProductos.Cell[0,i].AsBoolean = True then
                  begin
                      sNombre := gvProductos.Cells[1,i];
                      if sNombre = 'Todos' Then sNombre := '*';
                      SQLStr := 'INSERT INTO tblRouting(Nombre,Rou_From,Rou_Code,Rou_To) ' +
                                'VALUES(' + QuotedStr(sNombre) + ',' + Rou_From + ',' + QuotedStr(Rou_Code) +
                                ',' + Rou_To + ')';

                      Conn.Execute(SQLStr);
                  end;

                 // Si esta chekeado la opcion de todos no es necesairo agregar las demas rutas
                 //solo esta, ya que esta como su nombre lo dice incluye todos lo productos
                 if sNombre = '*' Then
                    Break;
        end;
    Conn.Close;
end;

function TfrmRouting.RouteExist(Rou_From,Rou_To: String):Boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := False;
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT TOP 1 R.Rou_From ' +
              'FROM tblrouting R ' +
              'INNER JOIN tblTareas T ON Rou_From = T.id ' +
              'INNER JOIN tblTareas T2 ON Rou_To = T2.id ' +
              'WHERE T.Nombre = ' + QuotedStr(Rou_From) + ' AND T2.Nombre = ' + QuotedStr(Rou_To);


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    If Qry.RecordCount > 0 Then
        Result := True;

    Qry.Close;
    Conn.Close;

end;

function TfrmRouting.GetTaskId(TaskName: String):String;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := '';
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT Id From tblTareas WHERE Nombre = ' + QuotedStr(TaskName);

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    If Qry.RecordCount > 0 Then
        Result := VarToStr(Qry['Id']);

    Qry.Close;
    Conn.Close;
end;

end.
