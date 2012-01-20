unit Year;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView, Columns,ColumnClasses,ExtCtrls, ComCtrls,IdTrivialFTPBase,Math,
  StrUtils,All_Functions,Chris_Functions,ADODB,DB,IniFiles,DateUtils,ComObj;

type
  TfrmYear = class(TForm)
    Label1: TLabel;
    btnGrabar: TButton;
    cmbYear: TComboBox;
    GridView1: TGridView;
    btnSave: TButton;
    btnCerrar: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnGrabarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCerrarClick(Sender: TObject);
    procedure BindGrid();
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmYear: TfrmYear;
  gsConnString,StartDDir: String;

implementation

uses Main;

{$R *.dfm}

procedure TfrmYear.FormCreate(Sender: TObject);
var sYear : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sYear := IniFile.ReadString('System','Year','');

    if sYear <> '' then
        cmbYear.Text := sYear
    else
        cmbYear.Text := IntToStr( YearOf(Date) );

    cmbYear.Items.Add(IntToStr(YearOf(Date) - 2));
    cmbYear.Items.Add(IntToStr(YearOf(Date) - 1));
    cmbYear.Items.Add(IntToStr(YearOf(Date)));
    cmbYear.Items.Add(IntToStr(YearOf(Date) + 1));
    cmbYear.Items.Add(IntToStr(YearOf(Date) + 2));

    BindGrid();
end;

procedure TfrmYear.btnGrabarClick(Sender: TObject);
var IniFile: TIniFile;
begin

    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    IniFile.WriteString('System','Year',cmbYear.Text);
    frmMain.StatusBar.Panels[3].Text := cmbYear.Text;

    ShowMessage('Se establecio el Añio ' + cmbYear.Text + ' satisfactoriamente.');
end;

procedure TfrmYear.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmYear.btnCerrarClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmYear.BindGrid();
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
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

    GridView1.ClearRows;
    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['SCR_ID']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['SCR_Name']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['SCR_FormName']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['SCR_Description']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['SCR_Year']);        
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmYear.btnSaveClick(Sender: TObject);
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

        for i := 0 to GridView1.RowCount - 1 do
        begin
            if IsNumeric(GridView1.Cells[4,i]) then begin
                SQLStr := 'UPDATE tblScreens SET SCR_Year = ' + QuotedStr(GridView1.Cells[4,i]) +
                          ' WHERE SCR_ID = ' + GridView1.Cells[0,i];

                conn.Execute(SQLStr);
            end
            else begin
                ShowMessage('El Añio capturado para la pantalla ' +
                            GridView1.Cells[1,GridView1.SelectedRow] + ' no es numerico. El registro ' +
                            'no sera actualizado.');
            end;
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : btnAddClick');
    end;

    Conn.Close;

end;

end.
