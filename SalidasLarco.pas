unit SalidasLarco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math, Mask;

type
  TfrmSalidasLarco = class(TForm)
    gbButtons: TGroupBox;
    lblId: TLabel;
    lblAnio: TLabel;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    Panel1: TPanel;
    Primero: TButton;
    Anterior: TButton;
    Ultimo: TButton;
    Siguiente: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    GroupBox2: TGroupBox;
    gvSalidas: TGridView;
    txtOrden: TMaskEdit;
    Label6: TLabel;
    Label3: TLabel;
    deFecha: TDateEditor;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure EnableButtons();
    Procedure BindData();
    Procedure ClearData();
    procedure BindDetalle(EntradaID: String);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure ActualizarDetalle(SalidaID: String);
    procedure PrimeroClick(Sender: TObject);
    procedure txtOrdenExit(Sender: TObject);
    procedure txtOrdenKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure gvSalidasAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: String; var Accept: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSalidasLarco: TfrmSalidasLarco;
  Conn : TADOConnection;
  Qry : TADOQuery;
  giOpcion : Integer;
  gsYear : String;
  sPermits : String;    
implementation

uses Main, Login;

{$R *.dfm}

procedure TfrmSalidasLarco.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmSalidasLarco.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);
  gsYear := RightStr(lblAnio.Caption,2) + '-';

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblSalidasLarco WHERE YEAR(SL_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' ORDER BY SL_ID';
  Qry.Open;

  ClearData();

  EnableButtons();
  BindData();
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmSalidasLarco.EnableControls(Value:Boolean);
begin
  txtOrden.Enabled := not Value;
  gvSalidas.Enabled := not Value;
  deFecha.Enabled := not Value;
end;

procedure TfrmSalidasLarco.EnableButtons();
begin
  Nuevo.Enabled := True;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  //Imprimir.Enabled := False;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
        //Imprimir.Enabled := True;
  end;
  EnableFormButtons(gbButtons, sPermits);  
end;


procedure TfrmSalidasLarco.ClearData();
begin
  txtOrden.Text := '';
  deFecha.Text := DateToStr(Now);
  gvSalidas.ClearRows;
end;

procedure TfrmSalidasLarco.BindData();
begin
  if Qry.RecordCount = 0 then
          Exit;

  lblId.Caption  := VarToStr(Qry['SL_ID']);
  txtOrden.Text := RightStr(VarToStr(Qry['SL_Orden']),10);
  deFecha.Text := VarToStr(Qry['SL_Fecha']);

  BindDetalle(lblId.Caption);
end;

procedure TfrmSalidasLarco.BindDetalle(EntradaID: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
  SQLStr := 'SELECT SD.SD_ID, M.MAT_Descripcion, SD.SD_Cantidad, ' +
            'SD.SL_Cantidad, SD.SL_Pedimento, (SD.SD_Cantidad - SD.SL_Cantidad) AS [Desperdicio] ' +
            'FROM tblSalidasDetalle SD ' +
            'INNER JOIN tblSalidas S ON S.SAL_ID = SD.SAL_ID ' +
            'INNER JOIN tblMateriales M ON SD.MAT_ID = M.MAT_ID ' +
            'WHERE S.SAL_Orden = ' + QuotedStr(gsYear + txtOrden.Text);

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  gvSalidas.ClearRows;
  While not Qry2.Eof do
  Begin
      gvSalidas.AddRow(1);
      gvSalidas.Cells[0,gvSalidas.RowCount -1] := VarToStr(Qry2['SD_ID']);
      gvSalidas.Cells[1,gvSalidas.RowCount -1] := VarToStr(Qry2['MAT_Descripcion']);
      gvSalidas.Cells[2,gvSalidas.RowCount -1] := VarToStr(Qry2['SD_Cantidad']);
      gvSalidas.Cells[3,gvSalidas.RowCount -1] := VarToStr(Qry2['SL_Cantidad']);
      gvSalidas.Cells[4,gvSalidas.RowCount -1] := VarToStr(Qry2['SL_Pedimento']);
      gvSalidas.Cells[5,gvSalidas.RowCount -1] := VarToStr(Qry2['Desperdicio']);
      Qry2.Next;
  End;

  Qry2.Close;
end;

procedure TfrmSalidasLarco.NuevoClick(Sender: TObject);
begin
  ClearData();
  EnableControls(False);
  txtOrden.SetFocus;
  giOpcion := 1;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  //Imprimir.Enabled := False;
end;

procedure TfrmSalidasLarco.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
//  Imprimir.Enabled := False;
  txtOrden.Enabled := False;
  deFecha.SetFocus;
end;

procedure TfrmSalidasLarco.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
//  Imprimir.Enabled := False;
end;

procedure TfrmSalidasLarco.btnCancelarClick(Sender: TObject);
begin
  ClearData();
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  EnableButtons();
  BindData();
end;

procedure TfrmSalidasLarco.btnAceptarClick(Sender: TObject);
var user: String;
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['SL_Orden'] := gsYear + txtOrden.Text;
        Qry['SL_Fecha'] := deFecha.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry.Post;

        ActualizarDetalle(Qry['SL_ID']);
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['SL_Orden'] := gsYear + txtOrden.Text;
        Qry['SL_Fecha'] := deFecha.Text;
        Qry['USE_ID'] := frmMain.sUserID;
        Qry.Post;

        ActualizarDetalle(Qry['SL_ID']);
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar esta Salida?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro esta Salida?',
                        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                      Application.CreateForm(TfrmLogin, frmLogin);
                      frmLogin.lblValidate.Caption := 'true';
                      if frmLogin.ShowModal <> mrOK then begin
                            ShowMessage('No tienes permiso para scrapear esta orden.');
                      end
                      else begin
                          user := frmLogin.txtUser.Text;

                          Qry.Edit;
                          Qry['USE_ID'] := user;
                          Qry.Post;

                          Qry.Delete;
                          ActualizarDetalle(lblId.Caption);
                          gvSalidas.ClearRows;
                      end;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtOrden.Text <> '' then
        begin
              if not Qry.Locate('SL_Orden',txtOrden.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Salida para esta Orden de Trabajo : ' + txtOrden.Text + '.', mtInformation,[mbOk], 0);
                    txtOrden.SetFocus;
                    Exit;
                end;
        end
        else if deFecha.Text <> '' then
        begin
              if not Qry.Locate('SL_Fecha',deFecha.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ninguna Salida con Fecha : ' + deFecha.Text + '.', mtInformation,[mbOk], 0);
                    deFecha.SetFocus;
                    Exit;
                end;
        end;
  end;

  ClearData();
  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  EnableButtons();
  BindData();
end;

function TfrmSalidasLarco.ValidateData():Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
  result := True;
  if txtOrden.Text = '   -   -  ' Then
    begin
      MessageDlg('Por favor ingrese el numero de Orden de Trabajo.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;

  SQLStr := 'SELECT ITE_Nombre FROM tblOrdenes WHERE ITE_Nombre = ' +
            QuotedStr(RightStr(lblAnio.Caption,2) + '-' + txtOrden.Text);

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount <= 0 then
    begin
      MessageDlg('El numero de Orden de Trabajo es incorrecto.', mtInformation,[mbOk], 0);
      result :=  False;
    end;

  SQLStr := 'SELECT SL_Orden FROM tblSalidasLarco WHERE SL_Orden = ' +
            QuotedStr(RightStr(lblAnio.Caption,2) + '-' + txtOrden.Text);

  if giOpcion = 2 then begin
        SQLStr := ' AND SL_ID <> ' + lblId.Caption;
  end;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then
    begin
      MessageDlg('El numero de Orden de Trabajo es incorrecto.', mtInformation,[mbOk], 0);
      result :=  False;
    end;


end;

procedure TfrmSalidasLarco.ActualizarDetalle(SalidaID: String);
var i : Integer;
SQLStr : String;
begin
  {SQLStr := 'UPDATE tblSalidasDetalle SET SL_Cantidad = NULL, SL_Pedimento = NULL, ED_ID = NULL ' +
            'FROM tblSalidas S ' +
            'INNER JOIN tblSalidasDetalle SD ON S.SAL_ID = SD.SAL_ID ' +
            'WHERE SAL_ORDEN = ' + QuotedStr(gsYear + txtOrden.Text);
  conn.Execute(SQLStr);
  }

  if (giOpcion = 1) then begin

      for i:= 0 to gvSalidas.RowCount - 1 do
      begin
          SQLStr := 'SalidasLarco ' + gvSalidas.Cells[0,i] + ',' + gvSalidas.Cells[3,i] +
                    ',' +  QuotedStr(gvSalidas.Cells[4,i]) + ',' + gvSalidas.Cells[5,i];
          conn.Execute(SQLStr);

      end;

  end
  else if (giOpcion = 3) then begin

      for i:= 0 to gvSalidas.RowCount - 1 do
      begin
          SQLStr := 'SalidasLarcoBorrar ' + gvSalidas.Cells[0,i];
          conn.Execute(SQLStr);

      end;

  end
  else if (giOpcion = 2) then begin

      for i:= 0 to gvSalidas.RowCount - 1 do
      begin
          SQLStr := 'SalidasLarcoEditar ' + gvSalidas.Cells[0,i] + ',' + gvSalidas.Cells[3,i] +
                    ',' +  QuotedStr(gvSalidas.Cells[4,i]) + ',' + gvSalidas.Cells[5,i];
          conn.Execute(SQLStr);

      end;

  end;

end;

procedure TfrmSalidasLarco.PrimeroClick(Sender: TObject);
begin
  if Qry.RecordCount = 0 then
          Exit;

  if (Sender as TButton).Caption = '| <' then
    Qry.First
  else if (Sender as TButton).Caption = '<' then
    Qry.Prior
  else if (Sender as TButton).Caption = '>' then
    Qry.Next
  else if (Sender as TButton).Caption = '> |' then
    Qry.Last;


  BindData();
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);  
//  Imprimir.Enabled := True;
end;


procedure TfrmSalidasLarco.txtOrdenExit(Sender: TObject);
var Qry2 : TADOQuery;
SQLStr : String;
begin
  SQLStr := 'SELECT SD.SD_ID, M.MAT_Descripcion, SD.SD_Cantidad ' +
            //'SD.SL_Cantidad, SD.SL_Pedimento, (SD.SD_Cantidad - SD.SL_Cantidad) AS [Desperdicio] ' +
            'FROM tblSalidasDetalle SD ' +
            'INNER JOIN tblSalidas S ON S.SAL_ID = SD.SAL_ID ' +
            'INNER JOIN tblMateriales M ON SD.MAT_ID = M.MAT_ID ' +
            'WHERE S.SAL_Orden = ' + QuotedStr(gsYear + txtOrden.Text);

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount <= 0 then
        ShowMessage('No se encontraron salidas para esta orden de trabajo.');

  gvSalidas.ClearRows;
  While not Qry2.Eof do
  Begin
      gvSalidas.AddRow(1);
      gvSalidas.Cells[0,gvSalidas.RowCount -1] := VarToStr(Qry2['SD_ID']);
      gvSalidas.Cells[1,gvSalidas.RowCount -1] := VarToStr(Qry2['MAT_Descripcion']);
      gvSalidas.Cells[2,gvSalidas.RowCount -1] := VarToStr(Qry2['SD_Cantidad']);
      gvSalidas.Cells[3,gvSalidas.RowCount -1] := VarToStr(Qry2['SD_Cantidad']);
      gvSalidas.Cells[4,gvSalidas.RowCount -1] := '';
      gvSalidas.Cells[5,gvSalidas.RowCount -1] := '0';
      Qry2.Next;
  End;

  Qry2.Close;
end;

procedure TfrmSalidasLarco.txtOrdenKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = vk_Return then
  begin
    txtOrdenExit(nil);
  end;
end;

procedure TfrmSalidasLarco.gvSalidasAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: String; var Accept: Boolean);
var icantidad, iValue: Double;
begin
  if ACol <> 3 then Exit;

  if ((ACol = 3) and (not IsNumeric(Value)) )then
  begin
      ShowMessage('La cantidad a descargar debe de ser numerica.');
      Accept := False;
      Exit;
  end;

  icantidad := StrToFloat(gvSalidas.Cells[2,ARow]);
  iValue := StrToFloat(Value);

  if iValue > icantidad then begin
      ShowMessage('La cantidad a descargar debe de ser menor o igual que la cantidad.');
      Accept := False;
      Exit;
  end;

  gvSalidas.Cells[5,ARow] := FloatToStr(icantidad - iValue);

end;

end.
