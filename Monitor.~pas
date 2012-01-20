unit Monitor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView, Columns,ColumnClasses,ExtCtrls, ComCtrls,IdTrivialFTPBase,Math,
  StrUtils,All_Functions,Chris_Functions,ADODB,DB,IniFiles, Menus,Clipbrd;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    gvItems: TGridView;
    gvPropiedades: TGridView;
    PopupMenu1: TPopupMenu;
    Refresh1: TMenuItem;
    PopupMenu2: TPopupMenu;
    Copiar1: TMenuItem;
    Copiarcomo1: TMenuItem;
    Separadoporcomas1: TMenuItem;
    Encomillas1: TMenuItem;
    procedure ControlMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ControlMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure ControlMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure Panel1DragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure Panel1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure TaskKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DrawLine(APoint1, APoint2: TPoint);
    procedure GridView1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TaskMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SaveTaskLocation(TaskName:String; X,Y: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DrawRouts();
    procedure FormPaint(Sender: TObject);
    procedure TaskSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure BindItems(Task,Status: String);
    procedure BindItemDetail(Item: String);
    procedure gvItemsSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure Refresh();
    procedure RefreshData();
    procedure Refresh1Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure gvItemsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure gvPropiedadesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Copiar1Click(Sender: TObject);
    procedure Separadoporcomas1Click(Sender: TObject);
    procedure Encomillas1Click(Sender: TObject);
  private
  inReposition : boolean;
  oldPos : TPoint;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
{$R WinXP.res}
{$R *.dfm}

Uses Main;

procedure TForm1.FormCreate(Sender: TObject);
begin

    //Panel1.OnMouseDown := ControlMouseDown;
    //Panel1.OnMouseMove := ControlMouseMove;
    Panel1.OnMouseUp := ControlMouseUp;

    Refresh();
end;


procedure TForm1.Panel1DragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
  //Accept := Source is TListBox;
  Accept := Source is TGridView;
end;

procedure TForm1.Panel1DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  if (Source is TGridView) then
  begin
      (Source as TGridView).Left := X;
      (Source as TGridView).Top := Y;
  end;

  SaveTaskLocation((Source as TGridView).Name,X,Y);
  Panel1.Repaint;
  DrawRouts;
end;


procedure TForm1.TaskKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var i:Integer;
begin
  if Key = VK_UP Then
  begin
       i:= (Sender as TListBox).ItemIndex;
       if i= 1 then
       begin
           (Sender as TListBox).Selected[2]:= True;
       end;
  end;

end;

procedure TForm1.DrawLine(APoint1, APoint2: TPoint);
var C: TControlCanvas;
Bitmap : TBitMap;
begin
Bitmap := TBitmap.Create;

C := TControlCanvas.Create;
C.Control := Panel1;
C.MoveTo(APoint1.X + 65 ,Apoint1.Y + 40);
//gvTask.Width := 130;
//gvTask.Height := 80;
if  Apoint2.Y > (Apoint1.Y + 81) then
begin
        C.LineTo(APoint2.X + 65 ,Apoint2.Y - 9);
        Bitmap.LoadFromFile(StartDDir + 'DownArrow.bmp');
        C.Draw(APoint2.X + 65 - 5,Apoint2.Y - 9,Bitmap);
end
else if Apoint2.Y < (Apoint1.Y - 81) then
begin
        C.LineTo(APoint2.X + 65 ,Apoint2.Y + 80 + 8);
        Bitmap.LoadFromFile(StartDDir + 'UpArrow.bmp');
        C.Draw(APoint2.X + 65 - 5,Apoint2.Y + 80,Bitmap);
end
else if Apoint2.X > (Apoint1.X + 131) then
begin
        C.LineTo(APoint2.X - 9 ,Apoint2.Y + 40);
        Bitmap.LoadFromFile(StartDDir + 'LeftArrow.bmp');
        C.Draw(APoint2.X - 9,Apoint2.Y + 40 -5,Bitmap);
end
else if Apoint2.X < (Apoint1.X - 131) then
begin
        C.LineTo(APoint2.X + 130 + 8 ,Apoint2.Y + 40);
        Bitmap.LoadFromFile(StartDDir + 'RightArrow.bmp');
        C.Draw(APoint2.X + 130,Apoint2.Y + 40 -5 ,Bitmap);
end;



//C.LineTo(APoint2.X - 9 ,Apoint2.Y + 40);
//C.Draw(APoint2.X - 9,Apoint2.Y + 40 - 5,Bitmap);
end;

procedure TForm1.GridView1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
(Sender as TGridView).BeginDrag(True);
end;

procedure TForm1.TaskMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
(Sender as TGridView).BeginDrag(True);
{  if (Sender is TWinControl) then
  begin
    inReposition:=True;
    SetCapture(TWinControl(Sender).Handle);
    GetCursorPos(oldPos);
  end;
 }
end;

procedure TForm1.SaveTaskLocation(TaskName:String; X,Y: Integer);
var Conn : TADOConnection;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;

    SQLStr := 'UPDATE tblMonitor SET MValue = ' + QuotedStr(IntToStr(X) + ',' + IntToStr(Y)) +
              ' WHERE MTYPE = ''Task'' AND MName = ' + QuotedStr(TaskName);

    Conn.Execute(SQLStr);
    Conn.Close;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TForm1.DrawRouts;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
APoint1,Apoint2: TPoint;
Location : TStringList;
begin
    Location := TStringList.Create;
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection := Conn;

    SQLStr := 'SELECT Rou_From,Rou_Code,Rou_To,M.MValue,M2.MValue AS MValue2 ' +
              'FROM tblrouting R ' +
              'INNER JOIN tblTareas T ON Rou_From = T.id ' +
              'INNER JOIN tblTareas T2 ON Rou_To = T2.id ' +
              'INNER JOIN tblMonitor M ON T.Nombre = M.MName ' +
              'INNER JOIN tblMonitor M2 ON T2.Nombre = M2.MName ' +
              'GROUP BY Rou_from,Rou_code,Rou_to,M.MValue,M2.MValue ' +
              'ORDER BY Rou_from';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    While not Qry.Eof do
    Begin
        Location.Clear;
        Location.CommaText := Qry['MValue'];
        Apoint1.X := StrToInt(Location[0]);
        Apoint1.Y := StrToInt(Location[1]);
        Location.Clear;
        Location.CommaText := Qry['MValue2'];
        Apoint2.X := StrToInt(Location[0]);
        Apoint2.Y := StrToInt(Location[1]);
        DrawLine(Apoint1,Apoint2);
        Qry.Next;
    end;


    Qry.Close;
    Conn.Close;
end;

procedure TForm1.FormPaint(Sender: TObject);
begin
   Panel1.Repaint;
   DrawRouts;
end;

procedure TForm1.TaskSelectCell(Sender: TObject; ACol, ARow: Integer);
begin
        gvItems.ClearRows;
        gvPropiedades.ClearRows;
        if StrToInt((Sender as TGridView).Cells[1,ARow]) > 0 then
          begin
                BindItems((Sender as TGridView).Name,(Sender as TGridView).Cells[0,Arow]);
          end;
end;

procedure TForm1.BindItems(Task,Status: String);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection := Conn;

    if Status <> 'Ordenes' then
    begin
          SQLStr := 'SELECT I.ITE_Nombre,I.ITE_ID ' +
                    'FROM tblItemTasks I ' +
                    'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                    'INNER JOIN tblDescriptions D ON I.ITS_Status = D.DES_CODE ' +
                    'WHERE T.[Nombre] = ' + QuotedStr(Task) + ' AND D.DEC_NOTE = ' + QuotedStr(Status) +
                    ' ORDER BY I.ITE_Nombre';
    end
    else begin
          SQLStr := 'SELECT S.ITE_Nombre,I.ITE_ID ' +
                    'FROM tblScrap S ' +
                    'INNER JOIN tblItems I ON S.ITE_Nombre = I.ITE_Nombre ' +
                    'ORDER BY S.ITE_Nombre';
    end;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    While not Qry.Eof do
    Begin
        gvItems.AddRow(1);
        gvItems.Cells[0,gvItems.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        gvItems.Cells[1,gvItems.RowCount -1] := VarToStr(Qry['ITE_ID']);
        Qry.Next;
    end;


    Qry.Close;
    Conn.Close;
end;

procedure TForm1.BindItemDetail(Item: String);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection := Conn;

    SQLStr := 'SELECT * FROM tblOrdenes O ' +
              'INNER JOIN tblItems I ON O.ITE_ID = I.ITE_ID ' +
              'WHERE O.ITE_ID = ' + Item;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvPropiedades.ClearRows;
    if Qry.RecordCount > 0 then
    begin
        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Tipo Proceso';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['TipoProceso']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Cantidad Requerida';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Requerida']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Cantidad Ordenada';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Ordenada']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Descripcion';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Producto']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Numero';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Numero']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Terminal';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Terminal']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Fecha Recibido';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Recibido']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Fecha Entrega';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Entrega']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Nombre';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Nombre']);

        gvPropiedades.AddRow(1);
        gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Prioridad';
        gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['ITE_Priority']);
    end;


    Qry.Close;
    Conn.Close;
end;

procedure TForm1.gvItemsSelectCell(Sender: TObject; ACol, ARow: Integer);
begin
BindItemDetail(gvItems.Cells[1,ARow]);
end;

procedure TForm1.Refresh();
var gvTask:TGridView;
Conn : TADOConnection;
SQLStr,sStatus : String;
Qry,Qry2 : TADOQuery;
Location : TStringList;
begin
    try

    While Panel1.ControlCount > 0 do
    Begin
        if (Panel1.Controls[0] is TGridView) then
                (Panel1.Controls[0] as TGridView).Free;
     end;

    application.ProcessMessages;

    Location := TStringList.Create;
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;;
    Conn.LoginPrompt := False;

    //Create Query commands
    Qry := TADOQuery.Create(nil);
    Qry.Connection := Conn;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT Id,Nombre FROM tblTareas ORDER BY Id';
    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    if Qry.RecordCount <= 0 then
        Exit;

    While not Qry.Eof do
    Begin
        SQLStr := 'SELECT MValue FROM tblMonitor WHERE MName = ' + QuotedStr(VarToStr(Qry['Nombre']));

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        Location.CommaText := Qry2['MValue'];
        gvTask := TGridView.Create(nil);
        gvTask.Parent := Panel1;
        gvTask.Name := VarToStr(Qry['Nombre']);
        gvTask.Left := StrToInt(Location[0]);
        gvTask.Top := StrToInt(Location[1]);
        gvTask.Width := 130;
        gvTask.Height := 80;
        if gvTask.Name = 'Calidad' then
                gvTask.Height := 90;
        gvTask.Font.Name := 'Tahoma';
        //gvTask.OnMouseDown := TaskMouseDown;

        gvTask.OnMouseDown := ControlMouseDown;
        gvTask.OnMouseMove := ControlMouseMove;
        gvTask.OnMouseUp := ControlMouseUp;

        gvTask.OnSelectCell := TaskSelectCell;
        gvTask.OnKeyDown := gvPropiedadesKeyDown;
        gvTask.Options := gvTask.Options + [goDissableColumnMoving];
        gvTask.Options := gvTask.Options + [goSelectFullRow];


        gvTask.Columns.Clear;
        gvTask.Columns.Add(TTextualColumn);
        gvTask.Columns.Add(TTextualColumn);
        gvTask.Columns[0].Header.Caption := VarToStr(Qry['Nombre']);
        gvTask.Columns[0].Width := 80;
        gvTask.Columns[0].Options := gvTask.Columns[0].Options - [coCanSort];
        gvTask.Columns[1].Width := 45;
        gvTask.Columns[1].Options := gvTask.Columns[0].Options - [coCanSort];

        sStatus := '3';
        if Qry['Nombre'] = 'Calidad' Then
                sStatus := '4';


        if Qry['Nombre'] <> 'Scrap' Then
        begin
            SQLStr := 'SELECT DES_CODE,DEC_NOTE,SUM(CASE WHEN ITE_ID IS NULL THEN 0 ELSE 1 END) AS Total ' +
                      'FROM tblDescriptions D ' +
                      'LEFT OUTER JOIN tblItemTasks I ON I.ITS_STATUS = D.DES_CODE AND I.Tas_ID = ' + VarToStr(Qry['Id']) + ' ' +
                      'WHERE DES_CODE < ' + sStatus + ' AND RTRIM(DES_GROUP) = ' + QuotedStr('TASK STATUSES') + ' ' +
                      'GROUP BY D.DEC_NOTE,D.DES_CODE ' +
                      'ORDER BY D.DES_CODE';
        end
        else begin
            SQLStr := 'SELECT ''Ordenes'' AS DEC_NOTE,COUNT(*) AS Total ' +
                      'FROM tblScrap ';
        end;
        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        gvTask.ClearRows;
        While not Qry2.Eof do
        Begin
            gvTask.AddRow(1);
            gvTask.Cells[0,gvTask.RowCount -1] := VarToStr(Qry2['DEC_NOTE']);
            gvTask.Cells[1,gvTask.RowCount -1] := VarToStr(Qry2['Total']);
            //gvTask.Cells[2,gvTask.RowCount -1] := VarToStr(Qry2['DES_CODE']);
            Qry2.Next;
        End;

        Qry.Next;
    end;

    Qry2.Close;
    Qry.Close;
    Conn.Close;
    application.ProcessMessages;

    DrawRouts;

    application.ProcessMessages;

    except

    end;
end;

procedure TForm1.RefreshData();
var Conn : TADOConnection;
SQLStr,sStatus : String;
Qry : TADOQuery;
i : integer;
begin
   // try

    frmMain.ProgressBar1.Visible := True;
    application.ProcessMessages;
    frmMain.ProgressBar1.Max := Panel1.ControlCount - 1;
    frmMain.ProgressBar1.Position := 0;
    frmMain.ProgressBar1.Step := 1;

    for i:=0 to Panel1.ControlCount - 1 do
    Begin
        if (Panel1.Controls[i] is TGridView) then
        begin
              //Create Connection
              Conn := TADOConnection.Create(nil);
              Conn.ConnectionString := frmMain.sConnString;;
              Conn.LoginPrompt := False;

              //Create Query commands
              Qry := TADOQuery.Create(nil);
              Qry.Connection := Conn;

              sStatus := '3';
              if (Panel1.Controls[i] as TGridView).Name = 'Calidad' Then
                      sStatus := '4';

              if (Panel1.Controls[i] as TGridView).Name <> 'Scrap' Then
              begin
                    SQLStr := 'SELECT DES_CODE,DEC_NOTE,SUM(CASE WHEN ITE_ID IS NULL THEN 0 ELSE 1 END) AS Total ' +
                              'FROM tblDescriptions D ' +
                              'LEFT OUTER JOIN tblItemTasks I ON I.ITS_STATUS = D.DES_CODE ' +
                              'AND I.TAS_ID = (SELECT ID FROM tblTareas WHERE Nombre = ' + QuotedStr((Panel1.Controls[i] as TGridView).Name) + ') ' +
                              'WHERE DES_CODE < ' + sStatus + ' AND RTRIM(DES_GROUP) = ' + QuotedStr('TASK STATUSES') + ' ' +
                              'GROUP BY D.DEC_NOTE,D.DES_CODE ' +
                              'ORDER BY D.DES_CODE ';
              end
              else begin
                  SQLStr := 'SELECT ''Ordenes'' AS DEC_NOTE,COUNT(*) AS Total ' +
                            'FROM tblScrap ';
              end;



              Qry.SQL.Clear;
              Qry.SQL.Text := SQLStr;
              Qry.Open;

              (Panel1.Controls[i] as TGridView).ClearRows;
              While not Qry.Eof do
              Begin
                  (Panel1.Controls[i] as TGridView).AddRow(1);
                  (Panel1.Controls[i] as TGridView).Cells[0,(Panel1.Controls[i] as TGridView).RowCount -1] := VarToStr(Qry['DEC_NOTE']);
                  (Panel1.Controls[i] as TGridView).Cells[1,(Panel1.Controls[i] as TGridView).RowCount -1] := VarToStr(Qry['Total']);
                  Qry.Next;
              End;

              frmMain.ProgressBar1.StepIt;
              //ShowMessage((Panel1.Controls[i] as TGridView).Name);
        end;
     end;
     frmMain.ProgressBar1.Visible := False;
     application.ProcessMessages;
end;



procedure TForm1.Refresh1Click(Sender: TObject);
begin
RefreshData;
end;

procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
                RefreshData;

end;

procedure TForm1.gvItemsKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
                RefreshData;
end;

procedure TForm1.gvPropiedadesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
                RefreshData;
end;

procedure TForm1.ControlMouseDown(
  Sender: TObject;
  Button: TMouseButton;
  Shift: TShiftState;
  X, Y: Integer);
begin
(Sender as TGridView).BeginDrag(True);
{  if (Sender is TWinControl) then
  begin
    inReposition:=True;
    SetCapture(TWinControl(Sender).Handle);
    GetCursorPos(oldPos);
  end;}
end; (*ControlMouseDown*)

procedure TForm1.ControlMouseMove(
  Sender: TObject;
  Shift: TShiftState;
  X, Y: Integer);
const
  minWidth = 20;
  minHeight = 20;
var
  newPos: TPoint;
begin
  if inReposition then
  begin
    with TWinControl(Sender) do
    begin
      GetCursorPos(newPos);

        Screen.Cursor := crHandPoint;//crSize;

        Left := Left - oldPos.X + newPos.X;
        Top := Top - oldPos.Y + newPos.Y;
        oldPos := newPos;

        if Left < 0 then Left := 0;
        if Left > Panel1.Width - Width then Left := Panel1.Width - Width;
        if Top < 0 then Top := 0;
        if Top > Panel1.Height - Height then Top := Panel1.Height - Height;
    end;
  end;
end; (*ControlMouseMove*)

procedure TForm1.ControlMouseUp(
  Sender: TObject;
  Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if inReposition then
  begin
    Screen.Cursor := crDefault;
    ReleaseCapture;
    inReposition := False;
    //Panel1.Repaint;
    //DrawRouts;
    //SaveTaskLocation((Sender as TGridView).Name,(Sender as TGridView).Left,(Sender as TGridView).Top);
  end;
end; (*ControlMouseUp*)


procedure TForm1.Copiar1Click(Sender: TObject);
begin
        if PopupMenu2.PopupComponent = gvItems then
           Clipboard.AsText := gvItems.Cells[0,gvItems.SelectedRow]

end;

procedure TForm1.Separadoporcomas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
        sText := '';
        if PopupMenu2.PopupComponent = gvItems then
        begin
           for i:= 0 to gvItems.RowCount - 1 do
                   sText := sText + gvItems.Cells[0,i] + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end;
end;

procedure TForm1.Encomillas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
        sText := '';
        if PopupMenu2.PopupComponent = gvItems then
        begin
           for i:= 0 to gvItems.RowCount - 1 do
                   sText := sText + QuotedStr(gvItems.Cells[0,i]) + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end;


end;

end.
