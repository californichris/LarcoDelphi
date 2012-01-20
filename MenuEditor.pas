unit MenuEditor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScrollView, CustomGridViewControl, CustomGridView, GridView,
  StdCtrls,Main,ADODB,DB, Menus,Clipbrd,ComObj;

type
  TfrmMenuEditor = class(TForm)
    GroupBox1: TGroupBox;
    GridView1: TGridView;
    GridView2: TGridView;
    GridView3: TGridView;
    btnDelete: TButton;
    btnAdd: TButton;
    btnCerrar: TButton;
    btnActualizar: TButton;
    Button1: TButton;
    btnUp: TButton;
    btnDown: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    procedure btnCerrarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BindCategorias();
    procedure BindPantallas();
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnActualizarClick(Sender: TObject);
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnAddClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnUpClick(Sender: TObject);
    procedure btnDownClick(Sender: TObject);
    procedure SaveCategoryOrder();
    procedure SaveScreenOrder();
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMenuEditor: TfrmMenuEditor;

implementation

{$R *.dfm}

procedure TfrmMenuEditor.btnCerrarClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmMenuEditor.FormCreate(Sender: TObject);
begin
        btnActualizarClick(nil);
        GridView1SelectCell(nil,0,GridView1.SelectedRow);        
end;

procedure TfrmMenuEditor.BindCategorias();
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

        SQLStr := 'SELECT * FROM tblCategories ORDER BY Category_Order,Category_ID';

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
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmMenuEditor.BindPantallas();
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

        SQLStr := 'SELECT * FROM tblScreens ' +
                  'WHERE SCR_ID NOT IN ( SELECT SCR_ID FROM tblCategory_Screens) AND ' +
                  'SCR_FormName <> ' + QuotedStr('space') + ' ORDER BY SCR_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView3.ClearRows;
        While not Qry.Eof do
        begin
            GridView3.AddRow(1);
            GridView3.Cells[0,GridView3.RowCount -1] := VarToStr(Qry['SCR_ID']);
            GridView3.Cells[1,GridView3.RowCount -1] := VarToStr(Qry['SCR_Name']);
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


procedure TfrmMenuEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmMenuEditor.btnActualizarClick(Sender: TObject);
begin
  BindCategorias();
  BindPantallas();
end;

procedure TfrmMenuEditor.GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
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

        SQLStr := 'SELECT C.Category_Id, C.SCR_ID, C.SCR_Order, S.SCR_Name FROM tblCategory_Screens C ' +
                  'INNER JOIN tblScreens S ON C.SCR_ID = S.SCR_ID WHERE C.Category_Id = ' +
                  GridView1.Cells[0,ARow]  + ' ORDER BY C.SCR_Order';


        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        GridView2.ClearRows;
        While not Qry.Eof do
        begin
            GridView2.AddRow(1);
            GridView2.Cells[0,GridView2.RowCount -1] := VarToStr(Qry['Category_ID']);
            GridView2.Cells[1,GridView2.RowCount -1] := VarToStr(Qry['SCR_ID']);
            GridView2.Cells[2,GridView2.RowCount -1] := VarToStr(Qry['SCR_Order']);
            GridView2.Cells[3,GridView2.RowCount -1] := VarToStr(Qry['SCR_Name']);
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

procedure TfrmMenuEditor.btnAddClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    if GridView3.RowCount <= 0 then begin
        Exit;
    end;

    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'INSERT INTO tblCategory_Screens(Category_Id, SCR_ID, SCR_Order) VALUES(' +
                  GridView1.Cells[0,GridView1.SelectedRow] + ',' +
                  GridView3.Cells[0,GridView3.SelectedRow] + ',' +
                  IntToStr(GridView2.RowCount + 1) + ')';

        conn.Execute(SQLStr);

        GridView1SelectCell(nil,0,GridView1.SelectedRow);
        BindPantallas();
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;
end;

procedure TfrmMenuEditor.btnDeleteClick(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    if GridView2.RowCount <= 0 then begin
        Exit;
    end;
{    else begin
       if GridView2.Cells[2,GridView2.SelectedRow] = '-' then begin
          GridView2.DeleteRow(GridView2.SelectedRow);
       end;
    end;
}
    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'DELETE FROM tblCategory_Screens WHERE Category_Id = ' +
                  GridView2.Cells[0,GridView2.SelectedRow] + ' AND SCR_ID = ' +
                  GridView2.Cells[1,GridView2.SelectedRow] + ' AND SCR_Order = ' +
                  GridView2.Cells[2,GridView2.SelectedRow] + ' ';

        conn.Execute(SQLStr);

        GridView1SelectCell(nil,0,GridView1.SelectedRow);
        BindPantallas();
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;


end;

procedure TfrmMenuEditor.Button1Click(Sender: TObject);
begin
frmMain.BindMenu(frmMain.sUserID);
end;

procedure TfrmMenuEditor.btnUpClick(Sender: TObject);
var iselectedRow : Integer;
begin
    if GridView1.SelectedRow = 0 then
        Exit;

    GridView1.MoveRow(GridView1.SelectedRow,GridView1.SelectedRow - 1);

    iselectedRow := GridView1.SelectedRow;
    SaveCategoryOrder();

    GridView1.SetFocus;
    GridView1.SelectedRow := iselectedRow - 1;
end;

procedure TfrmMenuEditor.btnDownClick(Sender: TObject);
var iselectedRow : Integer;
begin
    if GridView1.SelectedRow = GridView1.RowCount - 1 then
        Exit;

    GridView1.MoveRow(GridView1.SelectedRow,GridView1.SelectedRow + 1);
    iselectedRow := GridView1.SelectedRow;
    SaveCategoryOrder();

    GridView1.SetFocus;
    GridView1.SelectedRow := iselectedRow + 1;
end;

procedure TfrmMenuEditor.SaveCategoryOrder();
var SQLStr : String;
Conn : TADOConnection;
i: Integer;
begin
    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        for i := 0 to GridView1.RowCount - 1 do
        begin
            SQLStr := 'UPDATE tblCategories SET Category_Order = ' + IntToStr( i + 1) +
                      ' WHERE Category_Id = ' + GridView1.Cells[0,i];

            conn.Execute(SQLStr);
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;
end;

procedure TfrmMenuEditor.Button3Click(Sender: TObject);
var iselectedRow : Integer;
begin
    if GridView2.SelectedRow = 0 then
        Exit;

    GridView2.MoveRow(GridView2.SelectedRow,GridView2.SelectedRow - 1);

    iselectedRow := GridView2.SelectedRow;
    SaveScreenOrder();

    GridView1SelectCell(nil,0,GridView1.SelectedRow);

    GridView2.SetFocus;
    GridView2.SelectedRow := iselectedRow - 1;
end;

procedure TfrmMenuEditor.Button2Click(Sender: TObject);
var iselectedRow : Integer;
begin
    if GridView2.SelectedRow = GridView2.RowCount - 1 then
        Exit;

    GridView2.MoveRow(GridView2.SelectedRow,GridView2.SelectedRow + 1);
    iselectedRow := GridView2.SelectedRow;
    SaveScreenOrder();

    GridView1SelectCell(nil,0,GridView1.SelectedRow);    

    GridView2.SetFocus;
    GridView2.SelectedRow := iselectedRow + 1;
end;

procedure TfrmMenuEditor.SaveScreenOrder();
var SQLStr : String;
Conn : TADOConnection;
i: Integer;
begin
    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        for i := 0 to GridView2.RowCount - 1 do
        begin
            SQLStr := 'UPDATE tblCategory_Screens SET SCR_Order = ' + IntToStr( i + 1) +
                      ' WHERE Category_Id = ' + GridView2.Cells[0,i] +
                      ' AND SCR_ID = ' + GridView2.Cells[1,i] +
                      ' AND SCR_Order = ' + GridView2.Cells[2,i];

            conn.Execute(SQLStr);
        end;
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;
end;


procedure TfrmMenuEditor.Button4Click(Sender: TObject);
var SQLStr : String;
Conn : TADOConnection;
begin
    Conn := nil;
    try
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;

        SQLStr := 'INSERT INTO tblCategory_Screens(Category_Id, SCR_ID, SCR_Order) VALUES(' +
                  GridView1.Cells[0,GridView1.SelectedRow] + ',24,' +
                  IntToStr(GridView2.RowCount + 1) + ')';

        conn.Execute(SQLStr);

        GridView1SelectCell(nil,0,GridView1.SelectedRow);
        BindPantallas();
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;

end;

end.
