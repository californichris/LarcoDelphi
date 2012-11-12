unit Tareas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls,Larco_Functions;

type
  TfrmTareas = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    txtNombre: TEdit;
    btnAceptar: TButton;
    btnCancelar: TButton;
    PopupMenu1: TPopupMenu;
    Borrar1: TMenuItem;
    Editar1: TMenuItem;
    gvTareas: TGridView;
    Label3: TLabel;
    txtTiempo: TEdit;
    Label4: TLabel;
    Label2: TLabel;
    txtOrden: TEdit;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    btnNext: TButton;
    btnBack: TButton;
    Nuevo1: TMenuItem;
    lblID: TLabel;
    Image1: TImage;
    btnUp: TButton;
    btnDown: TButton;
    Image2: TImage;
    chkPrimera: TCheckBox;
    chkUltima: TCheckBox;
    txtInterno: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindTareas();
    procedure Borrar1Click(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function TareaExists(Tarea: String):Boolean;
    procedure gvTareasSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure btnNextClick(Sender: TObject);
    procedure btnBackClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure btnCancelarClick(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    procedure btnUpClick(Sender: TObject);
    procedure btnDownClick(Sender: TObject);
    procedure ReEnumerar();
    procedure SaveOrden();
    procedure Nuevo1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure BorrarTarea();
    function BoolToStrInt(Value:Boolean):String;
    procedure txtNombreKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTareas: TfrmTareas;
  giOpcion : Integer;
  giRow : Integer;
  gbOrden : Boolean;
  sPermits : String;  
implementation

uses Main;

{$R *.dfm}

procedure TfrmTareas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmTareas.FormCreate(Sender: TObject);
begin
    giOpcion := 0;
    gbOrden := False;
    BindTareas();
    txtOrden.Text := gvTareas.Cells[1,gvTareas.SelectedRow];
    txtNombre.Text := gvTareas.Cells[2,gvTareas.SelectedRow];
    txtTiempo.Text := gvTareas.Cells[3,gvTareas.SelectedRow];
    lblID.Caption := gvTareas.Cells[0,gvTareas.SelectedRow];
    chkPrimera.Checked  := gvTareas.Cell[4,gvTareas.SelectedRow].AsBoolean;
    chkUltima.Checked  := gvTareas.Cell[5,gvTareas.SelectedRow].AsBoolean;
    txtInterno.Text := gvTareas.Cells[6,gvTareas.SelectedRow];

    image1.Parent := btnUp;
    image1.Top := 0;
    image1.Left := 0;

    image2.Parent := btnDown;
    image2.Top := 0;
    image2.Left := 0;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
    
end;

procedure TfrmTareas.BindTareas();
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

    SQLStr := 'SELECT * FROM tblTareas Order By TAS_Order';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;


    gvTareas.ClearRows;
    While not Qry.Eof do
    Begin
        gvTareas.AddRow(1);
        gvTareas.Cells[0,gvTareas.RowCount -1] := VarToStr(Qry['Id']);
        gvTareas.Cells[1,gvTareas.RowCount -1] := VarToStr(Qry['TAS_Order']);
        gvTareas.Cells[2,gvTareas.RowCount -1] := VarToStr(Qry['Nombre']);
        gvTareas.Cells[3,gvTareas.RowCount -1] := VarToStr(Qry['Tiempo']);
        gvTareas.Cell[4,gvTareas.RowCount -1].AsBoolean  := StrToBool(VarToStr(Qry['IsPutOnly']));
        gvTareas.Cell[5,gvTareas.RowCount -1].AsBoolean  := StrToBool(VarToStr(Qry['IsLast']));
        gvTareas.Cells[6,gvTareas.RowCount -1] := VarToStr(VarToStr(Qry['Interno']));
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmTareas.Borrar1Click(Sender: TObject);
begin
  Nuevo.Enabled := False;
  Editar.Enabled := False;
  giOpcion := 3;
  EnableControls(False);
end;

procedure TfrmTareas.btnAceptarClick(Sender: TObject);
var Conn : TADOConnection;
SQLStr,sTarea : String;
begin
    txtNombre.Text := Trim(txtNombre.Text);
    txtTiempo.Text := Trim(txtTiempo.Text);
    txtInterno.Text := Trim(txtInterno.Text);

    if pos(' ', txtNombre.Text) <> 0 then begin
        MessageDlg('Espacios no son validos en el Nombre de Tarea.', mtInformation,[mbOk], 0);
        Exit;
    end;

    If txtNombre.Text = '' then
      begin
        MessageDlg('Por favor escriba un nombre de Tarea.', mtInformation,[mbOk], 0);
        Exit;
      end;

    if TareaExists(txtNombre.Text) then
      begin
        MessageDlg('Ya existe una Tarea con este nombre.', mtInformation,[mbOk], 0);
        Exit;
      end;

    if txtTiempo.Text = '' Then
      Begin
        MessageDlg('Por favor establesca el tiempo maximo de esta tarea.', mtInformation,[mbOk], 0);
        Exit;
      end;

    if not isNumeric(txtTiempo.Text) Then
      begin
        MessageDlg('El Tiempo debe de ser numerico.', mtInformation,[mbOk], 0);
        Exit;
      end;

    if txtInterno.Text = '' Then txtInterno.Text := '0';

    if not isNumeric(txtInterno.Text) Then
      begin
        MessageDlg('El Tiempo antes de fecha interna debe de ser numerico.', mtInformation,[mbOk], 0);
        Exit;
      end;


    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    if giOpcion = 2 Then
      Begin            //IsConjuncion,ConjuncionGrupo
        sTarea := gvTareas.Cells[2,gvTareas.SelectedRow];
        SQLStr := 'UPDATE tblTareas SET Nombre = ' +  QuotedStr(txtNombre.Text) +
                  ', Tiempo = ' + txtTiempo.Text + ', TAS_Order = ' + txtOrden.Text +
                  ', Interno = ' + txtInterno.Text +
                  ', IsPutOnly = ' + BoolToStrInt(chkPrimera.Checked)  +
                  ', IsLast = ' + BoolToStrInt(chkUltima.Checked)  +
                  ' WHERE Id = ' +  QuotedStr(lblID.Caption );

        Conn.Execute(SQLStr);

        SQLStr := 'UPDATE tblMonitor SET MName = ' + QuotedStr(txtNombre.Text) +
                  ' WHERE MType = ''Task'' AND MName = ' + QuotedStr(sTarea);

        Conn.Execute(SQLStr);

        if gbOrden Then //Si hubo algun cambio en el orden de las tareas
                SaveOrden();
      end
    else if giOpcion = 1 Then
      Begin
        SQLStr := 'INSERT INTO tblTareas(TAS_Order,Nombre,Tiempo,Interno,IsPutOnly,IsLast) ' +
                  'VALUES(' + txtOrden.Text  +',' + QuotedStr(txtNombre.Text) +
                  ',' + txtTiempo.Text + ',' + txtInterno.Text + ',' +
                  BoolToStrInt(chkPrimera.Checked) + ',' + BoolToStrInt(chkUltima.Checked) + ')';

        Conn.Execute(SQLStr);

        SQLStr := 'INSERT INTO tblMonitor(MType,MName,MValue) ' +
                  'VALUES(' + QuotedStr('Task') + ',' + QuotedStr(txtNombre.Text) +
                  ',' + QuotedStr('50,50') + ')';

        Conn.Execute(SQLStr);

        //insert records on tblitemTasks for existing items
      end
    else if giOpcion = 3 Then
      Begin
        BorrarTarea();
        if gvTareas.RowCount >= 1 then
                gvTareas.SelectedRow := 0;
      end;




    Conn.Close;

    BindTareas();
    txtNombre.Text := '';
    txtTiempo.Text := '';
    txtNombre.SetFocus;
    EnableControls(True);
    giOpcion := 0;
    gbOrden := False;
    gvTareas.SelectCell(2,gvTareas.SelectedRow);
    gvTareas.SetFocus;
  EnableFormButtons(gbButtons, sPermits);
end;

function TfrmTareas.TareaExists(Tarea: String):Boolean;
var i:integer;
begin
        TareaExists := False;
        for i:=0 to gvTareas.RowCount -1 do
          begin
                if (giOpcion = 2) and (i <> giRow ) then
                  if UT(Tarea) = UT(gvTareas.Cells[2,i]) then
                    TareaExists := True;
          end;
end;


procedure TfrmTareas.gvTareasSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
txtOrden.Text := gvTareas.Cells[1,gvTareas.SelectedRow];
txtNombre.Text := gvTareas.Cells[2,gvTareas.SelectedRow];
txtTiempo.Text := gvTareas.Cells[3,gvTareas.SelectedRow];
txtInterno.Text := gvTareas.Cells[6,gvTareas.SelectedRow];
lblID.Caption := gvTareas.Cells[0,gvTareas.SelectedRow];
chkPrimera.Checked  := gvTareas.Cell[4,gvTareas.SelectedRow].AsBoolean;
chkUltima.Checked  := gvTareas.Cell[5,gvTareas.SelectedRow].AsBoolean;
giRow := gvTareas.SelectedRow;
end;

procedure TfrmTareas.btnNextClick(Sender: TObject);
begin
gvTareas.SelectCell(2,gvTareas.SelectedRow + 1);
gvTareas.SetFocus;
end;

procedure TfrmTareas.btnBackClick(Sender: TObject);
begin
gvTareas.SelectCell(2,gvTareas.SelectedRow - 1);
gvTareas.SetFocus;
end;

procedure TfrmTareas.NuevoClick(Sender: TObject);
begin
  txtNombre.Text := '';
  txtTiempo.Text := '';
  txtInterno.Text := '';
  chkPrimera.Checked := False;
  chkUltima.Checked := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  EnableControls(False);
  giOpcion := 1;
  txtOrden.Text := IntToStr(StrToInt(gvTareas.Cells[1,gvTareas.RowCount - 1]) + 1);
  txtNombre.SetFocus;
end;

procedure TfrmTareas.EditarClick(Sender: TObject);
begin
  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  giOpcion := 2;
  EnableControls(False);
  txtNombre.SetFocus;
end;

procedure TfrmTareas.BorrarClick(Sender: TObject);
begin
  Nuevo.Enabled := False;
  Editar.Enabled := False;
  giOpcion := 3;

  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

end;

procedure TfrmTareas.EnableControls(Value:Boolean);
begin
        txtNombre.ReadOnly := Value;
        txtTiempo.ReadOnly := Value;
        txtInterno.ReadOnly := Value;
        chkUltima.Enabled  := not Value;
        chkPrimera.Enabled := not Value;
        //txtOrden.ReadOnly := Value;

        btnBack.Enabled  := Value;
        btnNext.Enabled := Value;

        btnAceptar.Enabled := not Value;
        btnCancelar.Enabled := not Value;

        if giOpcion = 2 then
        begin
          btnUp.Enabled := not Value;
          btnDown.Enabled := not Value;
          btnUp.Repaint;
          btnDown.Repaint;
        end;

        if Value then
        begin
          Nuevo.Enabled := Value;
          Borrar.Enabled := Value;
          Editar.Enabled := Value;
        end;

  EnableFormButtons(gbButtons, sPermits);        
end;

procedure TfrmTareas.btnCancelarClick(Sender: TObject);
begin
    EnableControls(True);
    BindTareas();
    giOpcion := 0;
    gbOrden := False;
    gvTareas.SelectCell(2,gvTareas.SelectedRow);
    gvTareas.SetFocus;
end;

procedure TfrmTareas.Image1Click(Sender: TObject);
var SelectedRow:Integer;
begin
if gvTareas.SelectedRow = 0 then
        Exit;

gbOrden := True;
gvTareas.MoveRow(gvTareas.SelectedRow,gvTareas.SelectedRow - 1);
SelectedRow := gvTareas.SelectedRow;
ReEnumerar();
gvTareas.SetFocus;
gvTareas.SelectedRow := SelectedRow - 1;
end;

procedure TfrmTareas.Image2Click(Sender: TObject);
var SelectedRow:Integer;
begin
if gvTareas.SelectedRow = gvTareas.RowCount -1 then
        Exit;

gbOrden := True;
gvTareas.MoveRow(gvTareas.SelectedRow,gvTareas.SelectedRow + 1);
SelectedRow := gvTareas.SelectedRow;
ReEnumerar();
gvTareas.SetFocus;
gvTareas.SelectedRow := SelectedRow + 1;
end;

procedure TfrmTareas.btnUpClick(Sender: TObject);
begin
Image1Click(nil);
end;

procedure TfrmTareas.btnDownClick(Sender: TObject);
begin
Image2Click(nil);
end;


procedure TfrmTareas.ReEnumerar();
var i : Integer;
begin
  for i:=0 to gvTareas.RowCount - 1 do
  begin
      gvTareas.Cells[1,i] := IntToStr(i + 1);
  end;

end;

procedure TfrmTareas.SaveOrden();
var Conn : TADOConnection;
SQLStr : String;
i: Integer;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

  for i:=0 to gvTareas.RowCount - 1 do
  begin
         SQLStr := 'UPDATE tblTareas SET TAS_Order = ' + gvTareas.Cells[1,i] +
                  ' WHERE Id = ' +  gvTareas.Cells[0,i];

        Conn.Execute(SQLStr);
  end;

  Conn.Close;
end;
procedure TfrmTareas.Nuevo1Click(Sender: TObject);
begin
    NuevoClick(nil);
end;

procedure TfrmTareas.Editar1Click(Sender: TObject);
begin
    EditarClick(nil);
end;

procedure TfrmTareas.BorrarTarea();
var sId,sTarea : string;
Conn : TADOConnection;
SQLStr : String;
begin
  sTarea := gvTareas.Cells[2,gvTareas.SelectedRow];
  sId := gvTareas.Cells[0,gvTareas.SelectedRow];

  if MessageDlg('Estas seguro que quieres borrar la Tarea ' + sTarea + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  SQLStr := 'DELETE FROM tblTareas WHERE Id = ' + sId;

  Conn.Execute(SQLStr);

  SQLStr := 'DELETE FROM tblMonitor WHERE MName = ' + QuotedStr(sTarea);

  Conn.Execute(SQLStr);

  SQLStr := 'DELETE FROM tblRouting WHERE Rou_From = ' + sId + ' or Rou_To = ' + sId;

  Conn.Execute(SQLStr);

  Conn.Close;

end;

function TfrmTareas.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;

procedure TfrmTareas.txtNombreKeyPress(Sender: TObject; var Key: Char);
begin
        if Key = Chr(vk_Space) then
            begin
              Key := #0;
            end

end;

end.
