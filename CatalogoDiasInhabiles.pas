unit CatalogoDiasInhabiles;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, Grids, Calendar, ScrollView, CustomGridViewControl,
  CustomGridView, GridView, CellEditors, chris_Functions, StdCtrls, Larco_Functions,ADODB,DB,
  ImgList, StrUtils;

type
  TfrmDiasInhabiles = class(TForm)
    gvNonWorkingDays: TGridView;
    yearList: TComboBox;
    add: TButton;
    GroupBox1: TGroupBox;
    deStart: TDateEditor;
    grabar: TButton;
    cancelar: TButton;
    imlGrid: TImageList;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindNonWorkingDays(Year: String);
    function DeleteNonWorkingDay(nonWorkingDay: String):Boolean;
    function InsertNonWorkingDay(nonWorkingDay: String):Boolean;
    function NonWorkingDayExists(nonWorkingDay: String):Boolean;    
    procedure yearListChange(Sender: TObject);
    procedure gvNonWorkingDaysCellClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure grabarClick(Sender: TObject);
    procedure addClick(Sender: TObject);
    procedure cancelarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDiasInhabiles: TfrmDiasInhabiles;
  Conn : TADOConnection;
  Qry : TADOQuery;

implementation

uses Main;
{$R *.dfm}


procedure TfrmDiasInhabiles.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmDiasInhabiles.FormCreate(Sender: TObject);
var formYear, i: Integer;
begin
  cancelarClick(nil);
  deStart.Date := Now;
  formYear := StrToInt(getFormYear(frmMain.sConnString,Self.Name));

  for i:= formYear - 5 to formYear + 5 do
  begin
    yearList.Items.Add(IntToStr(i));
  end;

  yearList.Text := IntToStr(formYear);

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;
  Qry := TADOQuery.Create(nil);
  Qry.Connection := Conn;

  BindNonWorkingDays(IntToStr(formYear));
end;

procedure TfrmDiasInhabiles.BindNonWorkingDays(year: String);
begin
  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT CONVERT(VARCHAR, NonWorkingDay, 101) AS NonWorkingDay FROM tblNonWorkingDay WHERE YEAR(NonWorkingDay) = ' + year + ' ORDER BY NonWorkingDay';
  Qry.Open;

  gvNonWorkingDays.ClearRows;
  While not Qry.Eof do
  Begin
      gvNonWorkingDays.AddRow(1);
      gvNonWorkingDays.Cells[0,gvNonWorkingDays.RowCount -1] := Qry['NonWorkingDay'];
      gvNonWorkingDays.Cell[1,gvNonWorkingDays.RowCount -1].AsInteger := 0;
      Qry.Next;
  End;

  Qry.Close;
end;

function TfrmDiasInhabiles.DeleteNonWorkingDay(nonWorkingDay: String):Boolean;
begin

  result := True;
  Try
  Begin
    Qry.SQL.Clear;
    Qry.SQL.Text := 'DELETE FROM tblNonWorkingDay WHERE NonWorkingDay = CONVERT(DATETIME, ' + QuotedStr(nonWorkingDay) + ', 101)';
    Qry.ExecSQL;
  End;
  Except
    result := False;
  End;

  Qry.Close;
end;

function TfrmDiasInhabiles.InsertNonWorkingDay(nonWorkingDay: String):Boolean;
begin

  result := True;
  Try
  Begin
    Qry.SQL.Clear;
    Qry.SQL.Text := 'INSERT INTO tblNonWorkingDay(NonWorkingDay) VALUES(CONVERT(DATETIME, ' + QuotedStr(nonWorkingDay) + ', 101))';
    Qry.ExecSQL;
  End;
  Except
    result := False;
  End;

  Qry.Close;
end;

function TfrmDiasInhabiles.NonWorkingDayExists(nonWorkingDay: String):Boolean;
begin

  result := True;
  Try
  Begin
    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT NonWorkingDay FROM tblNonWorkingDay WHERE NonWorkingDay = CONVERT(DATETIME, ' + QuotedStr(nonWorkingDay) + ', 101)';
    Qry.Open;

    if Qry.RecordCount > 0 then begin
      result := True;
    end
    else begin
      result := False;
    end;

  End;                             
  Except
    result := False;
  End;

  Qry.Close;
end;

procedure TfrmDiasInhabiles.yearListChange(Sender: TObject);
begin
  BindNonWorkingDays(yearList.Text);
end;

procedure TfrmDiasInhabiles.gvNonWorkingDaysCellClick(Sender: TObject;
  ACol, ARow: Integer);
var nonWorkingDay : String;
begin
  if ACol = 1 then begin
        if MessageDlg('Estas seguro que quieres borrar este dia?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
                nonWorkingDay := gvNonWorkingDays.Cells[0, ARow];
                if DeleteNonWorkingDay(nonWorkingDay) then begin
                  gvNonWorkingDays.DeleteRow(ARow);
                end
                else begin
                  ShowMessage('No fue posible borrar el dia inhabil.');
                end;
        end;
  end;
end;

procedure TfrmDiasInhabiles.grabarClick(Sender: TObject);
var nonWorkingDay : String;
begin
  nonWorkingDay := deStart.Text;
  if NonWorkingDayExists(nonWorkingDay) then begin
    ShowMessage('El dia inhabil ya existe.');
  end
  else begin
    if InsertNonWorkingDay(nonWorkingDay) then begin
      BindNonWorkingDays(yearList.Text);
    end
    else begin
      ShowMessage('No fue posible agregar el dia inhabil.');
    end;
  end;
end;


procedure TfrmDiasInhabiles.addClick(Sender: TObject);
begin
  GroupBox1.Visible := True;
  add.Enabled := False;
  yearList.Enabled := False;

  gvNonWorkingDays.Top := 80;
  gvNonWorkingDays.Height := 393;
end;

procedure TfrmDiasInhabiles.cancelarClick(Sender: TObject);
begin
  GroupBox1.Visible := False;
  add.Enabled := True;
  yearList.Enabled := True;

  gvNonWorkingDays.Top := 33;
  gvNonWorkingDays.Height := 440;
end;

end.
