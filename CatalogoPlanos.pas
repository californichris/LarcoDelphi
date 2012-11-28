unit CatalogoPlanos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math, Menus, ComObj,Clipbrd;

type
  TfrmCatalogoPlanos = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    txtPlano: TEdit;
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
    gbBuscar: TGroupBox;
    Label8: TLabel;
    lblTotal: TLabel;
    txtBuscarPlano: TEdit;
    Button1: TButton;
    gvResults: TGridView;
    Button2: TButton;
    SaveDialog1: TSaveDialog;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    CopiarOrden1: TMenuItem;
    cmbProductos: TComboBox;
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
    procedure txtPlanoExit(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PrimeroKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure txtBuscarPlanoExit(Sender: TObject);
    procedure txtBuscarPlanoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure gvResultsDblClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure txtPlanoKeyPress(Sender: TObject; var Key: Char);
    procedure txtInternoKeyPress(Sender: TObject; var Key: Char);
    procedure txtAliasKeyPress(Sender: TObject; var Key: Char);
    procedure BindProductos();
    procedure cmbProductosKeyPress(Sender: TObject; var Key: Char);
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

procedure TfrmCatalogoPlanos.BindProductos();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT UPPER(Nombre) AS Nombre FROM tblProductos ORDER BY Nombre';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbProductos.Items.Clear;
    While not Qry2.Eof do
    Begin
        cmbProductos.Items.Add(Qry2['Nombre']);
        Qry2.Next;
    End;

    cmbProductos.Text := '';
    Qry2.Close;
    Qry2.Free;
end;

procedure TfrmCatalogoPlanos.ClearData();
begin
  txtPlano.Text := '';
  cmbProductos.Text := '';
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

  BindProductos();

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
    cmbProductos.Text := VarToStr(Qry['PN_Descripcion']);

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
    cmbProductos.Enabled := not Value;
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
  gbButtons.Height := 0;
  gbBuscar.Height := 465;

  txtBuscarPlano.Text := '';
  gvResults.ClearRows;
  txtBuscarPlano.SetFocus();

  giOpcion := 4;
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
          Qry['PN_Descripcion'] := cmbProductos.Text;
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
          Qry['PN_Descripcion'] := cmbProductos.Text;
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

  txtPlano.Text := Trim(txtPlano.Text);
  if txtPlano.Text = '' Then
  begin
    MessageDlg('Por favor capture el Numero de Plano.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  cmbProductos.Text := Trim(cmbProductos.Text);
  if cmbProductos.Text = '' Then
  begin
    MessageDlg('Por favor seleccione la descripcion.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  if (cmbProductos.Items.IndexOf(cmbProductos.Text) = -1) then
  begin
      MessageDlg('Descripcion incorrecta seleccionelo de la lista.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
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
  txtInterno.Text := UpperCase(txtInterno.Text);
  if Trim(txtInterno.Text) = '' then begin
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
  txtAlias.Text := UpperCase(txtAlias.Text);
  if Trim(txtAlias.Text) = '' then begin
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
      BindData();
      EditarClick(nil);
    end;
end;

procedure TfrmCatalogoPlanos.txtPlanoExit(Sender: TObject);
begin
txtPlano.Text := UpperCase(txtPlano.Text);
end;

procedure TfrmCatalogoPlanos.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
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
    end;
  end;
end;

procedure TfrmCatalogoPlanos.PrimeroKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
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
    end;
  end;
end;

procedure TfrmCatalogoPlanos.Button2Click(Sender: TObject);
begin
  gbButtons.Height := 465;
  gbBuscar.Height := 0;
end;

procedure TfrmCatalogoPlanos.Button1Click(Sender: TObject);
var SQLStr, SQLWhere: String;
Qry2 : TADOQuery;
begin
  gvResults.ClearRows;
  lblTotal.Caption := '';
  if Trim(txtBuscarPlano.Text) = '' then begin
    ShowMessage('Numero de Plano es requerido.');
    Exit;
  end;

  txtBuscarPlano.Text := UpperCase(Trim(txtBuscarPlano.Text));
  SQLWhere := txtBuscarPlano.Text;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT P.PN_Id, P.PN_Numero,P.PN_Descripcion, SUM(CASE WHEN ST_TIPO = ''Entrada'' Then ST_Cantidad ELSE 0 END) - ' +
            'SUM(CASE WHEN ST_TIPO = ''Salida'' Then ST_Cantidad ELSE 0 END) AS Cantidad ' +
            'FROM tblPlano P LEFT OUTER JOIN tblStock S ON S.PN_Id = P.PN_Id WHERE P.PN_Numero ';
  if Pos('*', txtBuscarPlano.Text) <> 0 then begin
    SQLWhere := ' LIKE ' + QuotedStr(StringReplace(SQLWhere, '*', '%', [rfReplaceAll, rfIgnoreCase]));
  end
  else begin
    SQLWhere := ' = ' + QuotedStr(SQLWhere);
  end;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr + SQLWhere + ' GROUP BY P.PN_Id, P.PN_Numero, P.PN_Descripcion';
  Qry2.Open;

  while not Qry2.Eof do begin
      gvResults.AddRow(1);
      gvResults.Cells[0,gvResults.RowCount -1] := VarToStr(Qry2['PN_Id']);
      gvResults.Cells[1,gvResults.RowCount -1] := VarToStr(Qry2['PN_Numero']);
      gvResults.Cells[2,gvResults.RowCount -1] := VarToStr(Qry2['PN_Descripcion']);
      gvResults.Cells[3,gvResults.RowCount -1] := VarToStr(Qry2['Cantidad']);
      Qry2.Next;
  end;

end;

procedure TfrmCatalogoPlanos.txtBuscarPlanoExit(Sender: TObject);
begin
  txtBuscarPlano.Text := UpperCase(Trim(txtBuscarPlano.Text));
end;

procedure TfrmCatalogoPlanos.txtBuscarPlanoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if Key = vk_return then
  begin
    Button1Click(nil);
  end

end;

procedure TfrmCatalogoPlanos.gvResultsDblClick(Sender: TObject);
var id : String;
begin
  id := gvResults.Cell[0,gvResults.SelectedRow].AsString;
  Qry.Locate('PN_Id', id, [loPartialKey] );

  giOpcion := 0;
  ClearData();
  BindData();

  gbButtons.Height := 465;
  gbBuscar.Height := 0;
end;

procedure TfrmCatalogoPlanos.MenuItem1Click(Sender: TObject);
var sFileName: String;
begin
  if gvResults.RowCount = 0 then
  begin
          ShowMessage('No hay informacion que exportar.');
          Exit;
  end;

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(gvResults,sFileName);

  end;
end;

procedure TfrmCatalogoPlanos.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
var XApp : Variant;
Sheet : Variant;
Row,col :Integer;
begin
      Try //Create the excel object
      Begin
            XApp:= CreateOleObject('Excel.Application');
            //XApp.Visible := True;
            XApp.Visible := False;
            XApp.DisplayAlerts := False;
      end;
      except
       showmessage('No se pudo abrir Microsoft Excel,  parece que no esta instalado en el sistema.');
       exit;
      end;

      XApp.Workbooks.Add(xlWorkSheet);
      Sheet := XApp.Workbooks[1].WorkSheets[1];
      Sheet.Name := 'Scrap';

      for Col := 1 to Grid.Columns.Count do
              Sheet.Cells[1,Col] := Grid.Columns[Col - 1].Header.Caption;

      for Row := 1 to Grid.RowCount do
                for Col := 1 to Grid.Columns.Count do
                        Sheet.Cells[Row + 1,Col] := Grid.Cells[Col - 1,Row - 1];


      Sheet.Cells.Select;
      Sheet.Cells.EntireColumn.AutoFit;

      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

       showmessage('El archivo se creo exitosamente.');
end;

procedure TfrmCatalogoPlanos.CopiarOrden1Click(Sender: TObject);
begin
  Clipboard.AsText := gvResults.Cells[1,gvResults.SelectedRow]
end;

procedure TfrmCatalogoPlanos.txtPlanoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmCatalogoPlanos.txtInternoKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmCatalogoPlanos.txtAliasKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmCatalogoPlanos.cmbProductosKeyPress(Sender: TObject;
  var Key: Char);
begin
   Key := upcase(Key);
end;

end.
