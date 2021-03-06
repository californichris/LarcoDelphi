unit Editor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls,Larco_Functions;

type
  TfrmEditor = class(TForm)
    Label1: TLabel;
    lblOrden: TLabel;
    GridView1: TGridView;
    btnCancelar: TButton;
    btnApplicar: TButton;
    txtOrden: TEdit;
    btnBuscar: TButton;
    lblAnio: TLabel;
    procedure BindGrid(Orden: String);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure GridView1AfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: String; var Accept: Boolean);
    procedure txtOrdenKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnApplicarClick(Sender: TObject);
  private
    { Private declarations }
    Orden : String;
  public
    { Public declarations }
     gsYear : String;
  end;

var
  frmEditor: TfrmEditor;

implementation
{$R *.dfm}

Uses Main;

procedure TfrmEditor.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmEditor.FormCreate(Sender: TObject);
begin
    lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);
    gsYear := RightStr(lblAnio.Caption,2) + '-';

    if self.Orden <> '' then
            lblOrden.Caption := Self.Orden
    else
    begin
            txtOrden.Visible := True;
            btnBuscar.Visible := True;
    end;

end;

procedure TfrmEditor.BindGrid(Orden : String);
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    if Orden = '' then
        exit;


    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;


    SQLStr := 'ItemTasks ' + QuotedStr(Orden);


    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    GridView1.ClearRows;
    if Qry.RecordCount <= 0 then
        begin
                ShowMessage('No se encontro ninguna orden con el numero ' + Orden );
                Exit;
        end;

    While not Qry.Eof do
    Begin
        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['TAS_ID']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Tarea']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['Status']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Start']);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Stop']);
        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['Login']);
        Qry.Next;
    End;
end;


procedure TfrmEditor.btnBuscarClick(Sender: TObject);
var orden : String;
begin
  orden := txtOrden.Text;
  if(Length(orden) <= 10) then
    orden := gsYear + orden;

BindGrid(orden);
end;

procedure TfrmEditor.GridView1AfterEdit(Sender: TObject; ACol,  ARow: Integer; Value: String; var Accept: Boolean);
var i : integer;
begin

if ACol = 2 then
begin
        if Value = 'Vacio' then
          begin
                GridView1.Cells[3,ARow] := '';
                GridView1.Cells[4,ARow] := '';
                GridView1.Cells[5,ARow] := '';
                GridView1.Refresh;
                if ARow <> GridView1.RowCount then
                        GridView1.SelectCell(3, ARow - 1);

                GridView1.SetFocus;
          end
        else if Value = 'Listo' then
          begin
                if MessageDlg('Quieres compleatar todas las tareas previas a esta?',
                  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                begin
                        for i := 0 to ARow - 1 do
                          begin
                                if GridView1.Cells[2,ARow] <> 'Terminado' then
                                begin
                                   GridView1.Cells[2,i] := 'Terminado';
                                   if GridView1.Cells[3,i] = '' then GridView1.Cells[3,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                   if GridView1.Cells[4,i] = '' then GridView1.Cells[4,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                end;
                          end;
                end;

                GridView1.Cells[3,ARow] := '';
                GridView1.Cells[4,ARow] := '';
                GridView1.Cells[5,ARow] := '';
                GridView1.Refresh;
                if ARow <> GridView1.RowCount then
                        GridView1.SelectCell(3, ARow - 1);

                GridView1.SetFocus;
          end
        else if Value = 'Activo' then
          begin
                if MessageDlg('Quieres compleatar todas las tareas previas a esta?',
                  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                begin
                        for i := 0 to ARow - 1 do
                          begin
                                if GridView1.Cells[2,ARow] <> 'Terminado' then
                                begin
                                   GridView1.Cells[2,i] := 'Terminado';
                                   if GridView1.Cells[3,i] = '' then GridView1.Cells[3,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                   if GridView1.Cells[4,i] = '' then GridView1.Cells[4,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                end;
                          end;
                end;

                GridView1.Cells[3,ARow] := DateToStr(Now) + ' ' + TimeToStr(Time);
                GridView1.Cells[4,ARow] := '';
                GridView1.Cells[5,ARow] := '';
                GridView1.Refresh;
                if ARow <> GridView1.RowCount then
                        GridView1.SelectCell(4, ARow - 1);

                GridView1.SetFocus;
          end
        else
          begin
                if MessageDlg('Quieres compleatar todas las tareas previas a esta?',
                  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                begin
                        for i := 0 to ARow - 1 do
                          begin
                                if GridView1.Cells[2,ARow] <> 'Terminado' then
                                begin
                                   GridView1.Cells[2,i] := 'Terminado';
                                   if GridView1.Cells[3,i] = '' then GridView1.Cells[3,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                   if GridView1.Cells[4,i] = '' then GridView1.Cells[4,i] := DateToStr(Now) + ' ' + TimeToStr(Time);
                                end;
                          end;
                end;

                GridView1.Cells[3,ARow] := DateToStr(Now) + ' ' + TimeToStr(Time);
                GridView1.Cells[4,ARow] := DateToStr(Now) + ' ' + TimeToStr(Time);
                GridView1.Cells[5,ARow] := '';
                GridView1.Refresh;
                if ARow <> GridView1.RowCount then
                        GridView1.SelectCell(5, ARow - 1);

                GridView1.SetFocus;
          end


end;

if ACol = 5 then
begin
    for i := 0 to ARow - 1 do
      begin
            if (Value <> '') and (GridView1.Cells[2,i] <> 'Vacio') and (GridView1.Cells[2,i] <> 'Listo') and (GridView1.Cells[5,i] = '')then
            begin

               GridView1.Cells[5,i] := Value;
            end;
      end;

    GridView1.Refresh;
end;


end;

procedure TfrmEditor.txtOrdenKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        btnBuscarClick(nil);
   end;

end;

procedure TfrmEditor.btnCancelarClick(Sender: TObject);
begin
Self.Close;
end;

procedure TfrmEditor.btnApplicarClick(Sender: TObject);
var i:integer;
SQLStr,sStatus : String;
Conn : TADOConnection;
begin
    if MessageDlg('Estas seguro que quieres actualizar esta orden?',
      mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
        Exit;
    end;

    // agregar validacion a los datos
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    for i := 0 to GridView1.RowCount - 1 do
      begin
         if GridView1.Cells[2,i] <> '' then
         Begin
            if GridView1.Cells[2,i] = 'Vacio' then
            begin
                sStatus := '0';

                SQLStr := 'UPDATE tblItemTasks SET ITS_Status = NULL, ITS_DTStart = NULL' +
                           ', ITS_DTStop = NULL , USE_Login = NULL, ITS_Machine = NULL '
                           + ' WHERE TAS_ID = ' + GridView1.Cells[0,i] +
                           ' AND ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
            end
            else if GridView1.Cells[2,i] = 'Listo' then
            begin
                sStatus := '0';


                SQLStr := 'UPDATE tblItemTasks SET ITS_Status = ' + sStatus + ', ITS_DTStart = GETDATE()' +
                           ', ITS_DTStop = NULL , USE_Login = ' +  QuotedStr(GridView1.Cells[5,i]) +
                           ', ITS_Machine = ' + QuotedStr(GetLocalIP) + ' WHERE TAS_ID = ' +
                           GridView1.Cells[0,i] + ' AND ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
            end
            else if GridView1.Cells[2,i] = 'Activo' then
            begin
                sStatus := '1';


                SQLStr := 'UPDATE tblItemTasks SET ITS_Status = ' + sStatus + ', ITS_DTStart = ' +
                           QuotedStr(GridView1.Cells[3,i]) + ', ITS_DTStop = NULL ' +
                           ', USE_Login = ' +  QuotedStr(GridView1.Cells[5,i]) + ', ITS_Machine = ' +
                           QuotedStr(GetLocalIP) + ' WHERE TAS_ID = ' + GridView1.Cells[0,i] +
                           ' AND ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
            end
            else if GridView1.Cells[2,i] = 'Terminado' then
            begin
                sStatus := '2';


                SQLStr := 'UPDATE tblItemTasks SET ITS_Status = ' + sStatus + ', ITS_DTStart = ' +
                           QuotedStr(GridView1.Cells[3,i]) + ', ITS_DTStop = ' + QuotedStr(GridView1.Cells[4,i]) +
                           ', USE_Login = ' +  QuotedStr(GridView1.Cells[5,i]) + ', ITS_Machine = ' +
                           QuotedStr(GetLocalIP) + ' WHERE TAS_ID = ' + GridView1.Cells[0,i] +
                           ' AND ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
            end
            else if GridView1.Cells[2,i] = 'Retrabajo' then
            begin
                sStatus := '3';

                SQLStr := '';
            end
            else if GridView1.Cells[2,i] = 'Scrap' then
            begin
                sStatus := '9';

                SQLStr := '';
            end;


            if SQLStr <> '' then
                    conn.Execute(SQLStr);
         end;
    end;
    Conn.Close;
    ShowMessage('La orden se actualizo exitosamente');
    BindGrid(gsYear + txtOrden.Text);
end;

end.
