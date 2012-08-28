unit CatalogoPlanos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math;

type
  TfrmCatalogoPlanos = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtPlano: TEdit;
    txtDescripcion: TEdit;
    txtCantidad: TEdit;
    lblId: TLabel;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    lblAnio: TLabel;
    btnAceptar: TButton;
    btnCancelar: TButton;
    Panel1: TPanel;
    Primero: TButton;
    Anterior: TButton;
    Ultimo: TButton;
    Siguiente: TButton;
    Label4: TLabel;
    Label5: TLabel;
    gvInternos: TGridView;
    gvAliases: TGridView;
    txtInterno: TEdit;
    AddInternos: TButton;
    DeleteInternos: TButton;
    txtAlias: TEdit;
    AddAlias: TButton;
    DeleteAlias: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    Procedure ClearData();
    procedure FormCreate(Sender: TObject);
    procedure SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
    procedure ButtonKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
    procedure btnCancelarClick(Sender: TObject);
    procedure BindData();
    procedure EnableControls(Value:Boolean);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure EnableButtons();
    procedure PrimeroClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure txtInternoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure AddInternosClick(Sender: TObject);
    procedure AddAliasClick(Sender: TObject);
    procedure DeleteInternosClick(Sender: TObject);
    procedure DeleteAliasClick(Sender: TObject);
    procedure ActualizarDetalle(plano: String);
    function ValidatePlanoToBeDeleted(plano: String):Boolean;
    procedure EditarPlano(PlanoId: String);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCatalogoPlanos: TfrmCatalogoPlanos;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
  giOpcion : Integer;  //0= nada, 1 Nuevo, 2, editar, 3 borrar, 4 buscar
implementation

uses Main, Login;

{$R *.dfm}

procedure TfrmCatalogoPlanos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmCatalogoPlanos.ClearData();
begin
  txtPlano.Text := '';
  txtDescripcion.Text := '';
  txtCantidad.Text := '';

  txtInterno.Text := '';
  txtAlias.Text := '';
  
  gvInternos.ClearRows;
  gvAliases.ClearRows;
end;

procedure TfrmCatalogoPlanos.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;
  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblPlano ORDER BY PN_Numero';
  Qry.Open;

  if Qry.RecordCount > 0 then begin
    BindData();
  end;

  giOpcion := 0;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);

  EnableControls(True);
  EnableButtons();
end;

procedure TfrmCatalogoPlanos.BindData();
var SQLStr, tipo : String;
Qry2 : TADOQuery;
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['PN_Id']);
    txtPlano.Text := VarToStr(Qry['PN_Numero']);
    txtDescripcion.Text := VarToStr(Qry['PN_Descripcion']);

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT PA_Alias,PA_Tipo ' +
              'FROM tblPlano P ' +
              'INNER JOIN tblPlanoAlias PA ON P.PN_ID = PA.PN_ID ' +
              'WHERE P.PN_ID = ' + lblId.Caption + 'ORDER BY PA_Tipo, PA_Alias';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    While not Qry2.Eof do
    Begin
        tipo := VarToStr(Qry2['PA_Tipo']);

        if 'Interno' = tipo then begin
          gvInternos.AddRow(1);
          gvInternos.Cells[0,gvInternos.RowCount -1] := VarToStr(Qry2['PA_Alias']);
        end
        else begin
          gvAliases.AddRow(1);
          gvAliases.Cells[0,gvAliases.RowCount -1] := VarToStr(Qry2['PA_Alias']);
        end;

        Qry2.Next;
    End;

    Qry2.Close;
    Qry2.Free;
end;

procedure TfrmCatalogoPlanos.EnableControls(Value:Boolean);
begin
    txtPlano.ReadOnly := Value;
    txtDescripcion.ReadOnly := Value;
    txtCantidad.ReadOnly := Value;
    txtInterno.ReadOnly :=Value;
    txtAlias.ReadOnly := Value;
    AddInternos.Enabled := not Value;
    DeleteInternos.Enabled := not Value;
    AddAlias.Enabled := not Value;
    DeleteAlias.Enabled := not Value;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;

// Start of Action Buttons
procedure TfrmCatalogoPlanos.NuevoClick(Sender: TObject);
begin
  giOpcion := 1;
  ClearData();
  EnableControls(False);
  EnableButtons();

  txtPlano.SetFocus;
end;

procedure TfrmCatalogoPlanos.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  EnableButtons();

  txtPlano.SetFocus;
end;

procedure TfrmCatalogoPlanos.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  EnableButtons();

end;

procedure TfrmCatalogoPlanos.BuscarClick(Sender: TObject);
begin
  ShowMessage('No esta implementado todavia!!');
  Exit;

  
  ClearData();
  giOpcion := 4;

  EnableButtons();
end;

procedure TfrmCatalogoPlanos.btnCancelarClick(Sender: TObject);
begin
  giOpcion := 0;

  ClearData();
  BindData();

  EnableControls(True);
  EnableButtons();
end;

// End Of Action Buttons


procedure TfrmCatalogoPlanos.EnableButtons();
begin
  if giOpcion = 0 then begin
    Nuevo.Enabled := True;
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
    if Qry.RecordCount > 0 Then
    begin
          Editar.Enabled := True;
          Borrar.Enabled := True;
          Buscar.Enabled := True;

          EnableFormButtons(gbButtons, sPermits);          
    end;
  end
  else if giOpcion = 1 then begin
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
  end
  else if giOpcion = 2 then begin
    Nuevo.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;
  end
  else if giOpcion = 3 then begin
    Nuevo.Enabled := False;
    Editar.Enabled := False;
    Buscar.Enabled := False;
  end
  else if giOpcion = 4 then begin
    Nuevo.Enabled := False;
    Editar.Enabled := False;
    Borrar.Enabled := False;
  end;

  if giOpcion = 0 then begin
      btnAceptar.Enabled := False;
      btnCancelar.Enabled := False;
  end else begin
    btnAceptar.Enabled := True;
    btnCancelar.Enabled := True;
  end;

end;

procedure TfrmCatalogoPlanos.PrimeroClick(Sender: TObject);
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

  ClearData();
  BindData();
  EnableButtons();
end;

procedure TfrmCatalogoPlanos.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin

        if not ValidateData() then
          Exit;

        try
          Qry.Insert;
          Qry['PN_Numero'] := txtPlano.Text;
          Qry['PN_Descripcion'] := txtDescripcion.Text;
          Qry['Update_Date'] := DateTimeToStr(Now);
          Qry['Update_User'] := frmMain.sUserLogin;
          Qry.Post;

          ActualizarDetalle(Qry['PN_ID']);
        except
            on E : EDatabaseError do begin
              if Pos('PRIMARY KEY', E.Message) > 0 then begin
                ShowMessage('Ya existe un plano con este nombre.');
              end
              else begin
                ShowMessage(E.ClassName+' error raised, with message : '+E.Message);
              end;
              //Reloading recorset.
              Qry.Close;
              Qry.Open;

              Qry.Locate('PN_ID',lblId.Caption ,[loPartialKey] );
              Exit;
            end;
        end;

  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        try
          Qry.Edit;
          Qry['PN_Numero'] := txtPlano.Text;
          Qry['PN_Descripcion'] := txtDescripcion.Text;
          Qry['Update_Date'] := DateTimeToStr(Now);
          Qry['Update_User'] := frmMain.sUserLogin;
          Qry.Post;

          ActualizarDetalle(Qry['PN_ID']);          
        except
            on E : EDatabaseError do begin
              if Pos('PRIMARY KEY', E.Message) > 0 then begin
                ShowMessage('Ya existe un plano con este nombre.');
              end
              else begin
                ShowMessage(E.ClassName+' error raised, with message : '+E.Message);
              end;
              //Reloading recorset.
              Qry.Close;
              Qry.Open;

              Qry.Locate('PN_ID',lblId.Caption ,[loPartialKey] );
              Exit;
            end;
        end;

  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar este Plano?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro este Plano?',
                        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                if not ValidatePlanoToBeDeleted(lblId.Caption) then
                  Exit;

                //TODO: Validate againts tblOrdenes

                Qry.Edit;
                Qry['Update_Date'] := DateTimeToStr(Now);
                Qry['Update_User'] := frmMain.sUserLogin;
                Qry.Post;

                //Delete Childs first
                ActualizarDetalle(lblId.Caption);
                Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
    ShowMessage('No esta implementado todavia!!');
  end;

  giOpcion := 0;
  ClearData();
  EnableControls(True);
  EnableButtons();
  BindData();
end;

function TfrmCatalogoPlanos.ValidateData():Boolean;
begin
  ValidateData := True;
  if txtPlano.Text = '' Then
  begin
    MessageDlg('Por favor capture el Numero de Plano.', mtInformation,[mbOk], 0);
    result :=  False;
  end;

end;

procedure TfrmCatalogoPlanos.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      AppActivate(Application.Handle);
      SendKeys('{TAB}',False);
  end
  else if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
  begin
      btnCancelarClick(nil);
  end
  else if (Key = 83) and (ssCtrl in Shift)and (btnAceptar.Enabled = True)  then
  begin
      btnAceptarClick(nil);
  end;

  if giOpcion = 0 then begin
    if (Key = 78) and (ssCtrl in Shift)then
    begin
        NuevoClick(nil);
    end
    else if (Key = 69) and (ssCtrl in Shift)then
    begin
        EditarClick(nil);
    end
    else if (Key = 68) and (ssCtrl in Shift)then
    begin
        BorrarClick(nil);
    end
    else if (Key = 66) and (ssCtrl in Shift)then
    begin
        BuscarClick(nil);
    end
    else if (Key = 88) and (ssCtrl in Shift)then
    begin
        Self.Close;
    end;
  end;

end;

procedure TfrmCatalogoPlanos.ButtonKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
  if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
  begin
      btnCancelarClick(nil);
  end
  else if (Key = 83) and (ssCtrl in Shift)and (btnAceptar.Enabled = True)  then
  begin
      btnAceptarClick(nil);
  end;

  if giOpcion = 0 then begin
    if (Key = 78) and (ssCtrl in Shift)then
    begin
        NuevoClick(nil);
    end
    else if (Key = 69) and (ssCtrl in Shift)then
    begin
        EditarClick(nil);
    end
    else if (Key = 68) and (ssCtrl in Shift)then
    begin
        BorrarClick(nil);
    end
    else if (Key = 66) and (ssCtrl in Shift)then
    begin
        BuscarClick(nil);
    end
    else if (Key = 88) and (ssCtrl in Shift)then
    begin
        Self.Close;
    end;
  end;
end;

procedure TfrmCatalogoPlanos.txtInternoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If Key = vk_return then
  begin
    if (Sender as TEdit).Name = 'txtInterno' then begin
        AddInternosClick(AddInternos);
    end
    else begin
        AddAliasClick(AddAlias);
    end;

  end
  else if (Key = vk_Escape) and (btnCancelar.Enabled = True)  then
  begin
      btnCancelarClick(nil);
  end
  else if (Key = 83) and (ssCtrl in Shift)and (btnAceptar.Enabled = True)  then
  begin
      btnAceptarClick(nil);
  end;

  if giOpcion = 0 then begin
    if (Key = 78) and (ssCtrl in Shift)then
    begin
        NuevoClick(nil);
    end
    else if (Key = 69) and (ssCtrl in Shift)then
    begin
        EditarClick(nil);
    end
    else if (Key = 68) and (ssCtrl in Shift)then
    begin
        BorrarClick(nil);
    end
    else if (Key = 66) and (ssCtrl in Shift)then
    begin
        BuscarClick(nil);
    end
    else if (Key = 88) and (ssCtrl in Shift)then
    begin
        Self.Close;
    end;
  end;
end;

procedure TfrmCatalogoPlanos.AddInternosClick(Sender: TObject);
var i : Integer;
SQLStr: String;
Qry2 : TADOQuery;
begin
  if txtInterno.Text = '' then begin
      ShowMessage('Nombre Interno es requerido.');
      Exit;
  end;

  if txtInterno.Text = txtPlano.Text then begin
      ShowMessage('El Nombre Interno es igual que el Nombre del Plano.');
      Exit;
  end;

  for i:= 0 to gvInternos.RowCount - 1 do
  begin
     if gvInternos.Cells[0,i] = txtInterno.Text then begin
      ShowMessage('Ya existe este Nombre Interno.');
      Exit;
     end;
  end;

  for i:= 0 to gvAliases.RowCount - 1 do
  begin
     if gvAliases.Cells[0,i] = txtInterno.Text then begin
      ShowMessage('Ya existe un Alias con este Nombre.');
      Exit;
     end;
  end;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PA_ID FROM tblPlanoAlias WHERE PA_Alias = ' + QuotedStr(txtInterno.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    ShowMessage('Ya existe este Nombre Interno para otro Numero de Plano.');
    Exit;
  end;

  Qry2.Close;
  Qry2.Free;

  gvInternos.AddRow(1);
  gvInternos.Cells[0,gvInternos.RowCount -1] := txtInterno.Text;

  txtInterno.Text := '';
  txtInterno.SetFocus;

end;

procedure TfrmCatalogoPlanos.AddAliasClick(Sender: TObject);
var i : Integer;
SQLStr: String;
Qry2 : TADOQuery;
begin
  if txtAlias.Text = '' then begin
      ShowMessage('Alias es requerido.');
      Exit;
  end;

  if txtAlias.Text = txtPlano.Text then begin
      ShowMessage('El Alias capturado es igual que el Nombre del Plano.');
      Exit;
  end;

  for i:= 0 to gvInternos.RowCount - 1 do
  begin
     if gvInternos.Cells[0,i] = txtAlias.Text then begin
      ShowMessage('Ya existe un Nombre Interno con este Alias.');
      Exit;
     end;
  end;

  for i:= 0 to gvAliases.RowCount - 1 do
  begin
     if gvAliases.Cells[0,i] = txtAlias.Text then begin
      ShowMessage('Ya existe este Alias.');
      Exit;
     end;
  end;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PA_ID FROM tblPlanoAlias WHERE PA_Alias = ' + QuotedStr(txtAlias.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    ShowMessage('Ya existe este Alias para otro Numero de Plano.');
    Exit;
  end;

  Qry2.Close;
  Qry2.Free;

  gvAliases.AddRow(1);
  gvAliases.Cells[0,gvAliases.RowCount -1] := txtAlias.Text;

  txtAlias.Text := '';
  txtAlias.SetFocus;

end;

procedure TfrmCatalogoPlanos.DeleteInternosClick(Sender: TObject);
begin
  gvInternos.DeleteRow(gvInternos.SelectedRow);
end;

procedure TfrmCatalogoPlanos.DeleteAliasClick(Sender: TObject);
begin
  gvAliases.DeleteRow(gvAliases.SelectedRow);
end;

procedure TfrmCatalogoPlanos.ActualizarDetalle(plano: String);
var i : Integer;
SQLStr : String;
sDate: String;
begin
  if (giOpcion = 2) or (giOpcion = 3) then begin
      SQLStr := 'DELETE FROM tblPlanoAlias WHERE PN_Id = ' + plano;
      conn.Execute(SQLStr);
  end;

  if (giOpcion = 1) or (giOpcion = 2) then begin
      sDate := DateTimeToStr(Now);
      for i:= 0 to gvInternos.RowCount - 1 do
      begin
            SQLStr := 'INSERT INTO tblPlanoAlias(PN_Id, PA_Alias, PA_Tipo, Update_Date, Update_User) ' +
                      'VALUES(' + plano + ',' + QuotedStr(gvInternos.Cells[0,i]) +
                      ',''Interno'',' + QuotedStr(sDate) + ',' + frmMain.sUserLogin +')';

            conn.Execute(SQLStr);
      end;

      for i:= 0 to gvAliases.RowCount - 1 do
      begin
            SQLStr := 'INSERT INTO tblPlanoAlias(PN_ID, PA_Alias, PA_Tipo, Update_Date, Update_User) ' +
                      'VALUES(' + plano + ',' + QuotedStr(gvAliases.Cells[0,i]) +
                      ',''Alias'',' + QuotedStr(sDate) + ',' + frmMain.sUserLogin +')';

            conn.Execute(SQLStr);
      end;
  end;

end;

function TfrmCatalogoPlanos.ValidatePlanoToBeDeleted(plano: String):Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  ValidatePlanoToBeDeleted := True;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PN_ID FROM tblStock WHERE PN_ID = ' + plano;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    ShowMessage('No puedes borrar este Plano, por que hay piezas en Stock.');
    result :=  False;
  end;

  Qry2.Close;
  Qry2.Free;

end;

procedure TfrmCatalogoPlanos.EditarPlano(PlanoId: String);
begin
    if Qry.Locate('PN_ID',PlanoId ,[loPartialKey] ) then begin
      EditarClick(nil);
    end;
end;

end.