unit EntradasSalidasStock;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math, Mask;

type
  TfrmESStock = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lblId: TLabel;
    lblAnio: TLabel;
    txtPlano: TEdit;
    txtCantidad: TEdit;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    Panel1: TPanel;
    Primero: TButton;
    Anterior: TButton;
    Ultimo: TButton;
    Siguiente: TButton;
    Label4: TLabel;
    deFecha: TDateEditor;
    Label5: TLabel;
    cmbTipo: TComboBox;
    AddPlano: TButton;
    txtOrden: TMaskEdit;
    lblPNId: TLabel;
    lblValidOrden: TLabel;
    ddlAnio: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure BindData();
    Procedure ClearData();
    procedure EnableControls(Value:Boolean);
    procedure EnableButtons();
    procedure PrimeroClick(Sender: TObject);
    procedure txtPlanoExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure txtOrdenExit(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    function GetNumeroDePlano(PlanoId: String):String;
    procedure SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
    procedure AddPlanoClick(Sender: TObject);
    function ValidatePlano():Boolean;
    function ValidateOrden():Boolean;
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmESStock: TfrmESStock;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
  giOpcion : Integer;  //0= nada, 1 Nuevo, 2, editar, 3 borrar, 4 buscar
  gsYear,gsOYear : String;

implementation

uses Main, CatalogoPlanos, CatalogoPlanosModal;

{$R *.dfm}

procedure TfrmESStock.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action := caFree;
end;

// Start of Action Buttons
procedure TfrmESStock.NuevoClick(Sender: TObject);
begin
  giOpcion := 1;
  ClearData();
  EnableControls(False);
  EnableButtons();

  txtPlano.SetFocus;
end;

procedure TfrmESStock.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  EnableButtons();

  txtPlano.SetFocus;
end;

procedure TfrmESStock.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  EnableButtons();
end;

procedure TfrmESStock.BuscarClick(Sender: TObject);
begin
  ShowMessage('No esta implementado todavia!!');
  Exit;

  ClearData();
  giOpcion := 4;

  EnableButtons();
end;

procedure TfrmESStock.btnCancelarClick(Sender: TObject);
begin
  giOpcion := 0;

  ClearData();
  BindData();

  EnableControls(True);
  EnableButtons();
end;

// End Of Action Buttons

procedure TfrmESStock.ClearData();
begin
  txtPlano.Text := '';
  txtOrden.Text := '';
  txtCantidad.Text := '';
  deFecha.Text := '';
  cmbTipo.Text := '';
  lblValidOrden.Caption := '';
  lblPNId.Caption := '';
end;

procedure TfrmESStock.EnableControls(Value:Boolean);
begin
    txtPlano.ReadOnly := Value;
    txtOrden.ReadOnly := Value;
    txtCantidad.ReadOnly := Value;
    deFecha.Enabled := not Value;
    cmbTipo.Enabled := not Value;
    AddPlano.Enabled := not Value;
    ddlAnio.Enabled := not Value;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;

procedure TfrmESStock.EnableButtons();
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

procedure TfrmESStock.BindData();
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['ST_Id']);
    lblPNId.Caption := VarToStr(Qry['PN_Id']);
    txtPlano.Text := GetNumeroDePlano(lblPNId.Caption);
    txtOrden.Text := RightStr( VarToStr(Qry['ITE_Nombre']), Length(VarToStr(Qry['ITE_Nombre']))-3 );
    deFecha.Text := VarToStr(Qry['ST_Fecha']);
    txtCantidad.Text := VarToStr(Qry['ST_Cantidad']);
    cmbTipo.Text := VarToStr(Qry['ST_Tipo']);
end;

procedure TfrmESStock.PrimeroClick(Sender: TObject);
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

procedure TfrmESStock.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString, Self.Name);
  gsOYear := RightStr(lblAnio.Caption,2);
  gsYear := gsOYear + '-';

  ddlAnio.Clear;
  ddlAnio.Items.Add('2006');
  ddlAnio.Items.Add('2007');
  ddlAnio.Items.Add('2008');
  ddlAnio.Items.Add('2009');
  ddlAnio.Items.Add('2010');
  ddlAnio.Items.Add('2011');
  ddlAnio.Items.Add('2012');
  ddlAnio.ItemIndex := 6;

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;
  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblStock S WHERE YEAR(ST_Fecha) = ' +
                  QuotedStr(lblAnio.Caption) + ' ORDER BY ST_Fecha Desc, ST_ID Desc';
  Qry.Open;

  if Qry.RecordCount > 0 then begin
    BindData();
  end;

  giOpcion := 0;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);

  EnableControls(True);
  EnableButtons();
end;

procedure TfrmESStock.txtPlanoExit(Sender: TObject);
var found : boolean;
begin
  if (giOpcion = 0) or (Trim(txtPlano.Text) = '') then
    Exit;

  found := ValidatePlano();

  if not found then begin
    if MessageDlg('El Numero de Plano no existe, deseas agregarlo?',
              mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Application.CreateForm(TfrmCatalogoPlanos,frmCatalogoPlanos);
      frmCatalogoPlanos.NuevoClick(nil);
      frmCatalogoPlanos.txtPlano.Text := txtPlano.Text;
      frmCatalogoPlanos.Show();
    end
  end;
end;

procedure TfrmESStock.txtOrdenExit(Sender: TObject);
var SQLStr, sOrden, sNoParte, sAlias, sPlano: String;
Qry2 : TADOQuery;
begin
  sOrden := TrimRight( StringReplace(txtOrden.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
  if (giOpcion = 0) or (Trim(sOrden) = '') then
    Exit;

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  //SQLStr := 'SELECT ITE_Id,Numero FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
  SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblOrdenes O ' +
            'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
            'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
            'WHERE ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  lblValidOrden.Caption := '';

  if Qry2.RecordCount > 0 then begin
    lblValidOrden.Caption := VarToStr(Qry2['ITE_Id']);
    sNoParte := VarToStr(Qry2['Numero']);
    sAlias := VarToStr(Qry2['PA_Alias']);
    sPlano := VarToStr(Qry2['PN_Numero']);

    if (Trim(sNoParte) <> Trim(txtPlano.Text)) and (sAlias = '') then begin
      if MessageDlg('El Numero de Parte [' + sNoParte + '] de esta orden no es un Nombre Interno o Alias de este numero de Plano, deseas agregarlo?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        Application.CreateForm(TfrmCatalogoPlanos,frmCatalogoPlanos);
        if lblPNId.Caption <> '' then begin
          frmCatalogoPlanos.EditarPlano(lblPNId.Caption);
          frmCatalogoPlanos.txtInterno.Text := sNoParte;
          frmCatalogoPlanos.txtAlias.Text := sNoParte;
        end;

        frmCatalogoPlanos.Show();
      end;
    end;

  end
  else begin
    ShowMessage('La Orden de Trabajo no es valida.');
  end;

  Qry2.Close;
  Qry2.Free;
end;

procedure TfrmESStock.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['PN_Id'] := lblPNId.Caption;
        Qry['ITE_Nombre'] := gsYear + txtOrden.Text;
        Qry['ST_Cantidad'] := txtCantidad.Text;
        Qry['ST_Fecha'] := deFecha.Text;
        Qry['ST_Tipo'] := 'Entrada';
        Qry['Update_Date'] := DateTimeToStr(Now);
        Qry['Update_User'] := frmMain.sUserLogin;
        Qry.Post;

  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['PN_Id'] := lblPNId.Caption;
        Qry['ITE_Nombre'] := gsYear + txtOrden.Text;
        Qry['ST_Cantidad'] := txtCantidad.Text;
        Qry['ST_Fecha'] := deFecha.Text;
        Qry['ST_Tipo'] := 'Entrada';
        Qry['Update_Date'] := DateTimeToStr(Now);
        Qry['Update_User'] := frmMain.sUserLogin;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar este registro?',
                mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro este registro?',
                        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin

                Qry.Edit;
                Qry['Update_Date'] := DateTimeToStr(Now);
                Qry['Update_User'] := frmMain.sUserLogin;
                Qry.Post;

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

function TfrmESStock.ValidateData():Boolean;
var sOrden: String;
begin
  ValidateData := True;
  sOrden := TrimRight( StringReplace(txtOrden.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
  if Trim(txtPlano.Text) = '' Then
  begin
    MessageDlg('Por favor capture el Numero de Plano.', mtInformation,[mbOk], 0);
    result :=  False;
  end;

  if Trim(lblPNId.Caption) = '' Then
  begin
    if not ValidatePlano() then
    begin
      MessageDlg('El Numero de Plano no es valido.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;
  end;

  if Trim(sOrden) = '' Then
  begin
    MessageDlg('Por favor capture la Orden de Trabajo.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  if Trim(lblValidOrden.Caption) =  '' Then
  begin
    if not ValidateOrden() then
    begin
      MessageDlg('La Orden de Trabajo no es valida.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;
  end;

  if Trim(txtCantidad.Text) = '' Then
  begin
    MessageDlg('Por favor capture la Cantidad.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  if not IsNumeric(Trim(txtCantidad.Text)) then
  begin
    MessageDlg('La Cantidad debe ser un valor numerico.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  if Trim(deFecha.Text) = '' then
  begin
    MessageDlg('La Fecha es requerida.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;


  {
  if Trim(cmbTipo.Text) = '' Then
  begin
    MessageDlg('Por favor seleccione el Tipo.', mtInformation,[mbOk], 0);
    result :=  False;
  end;
  }

end;

function TfrmESStock.ValidatePlano():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := False;
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PN_Id FROM tblPlano WHERE PN_Numero = ' + QuotedStr(txtPlano.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  lblPNId.Caption := '';
  if Qry2.RecordCount > 0 then begin
    lblPNId.Caption := VarToStr(Qry2['PN_Id']);
    result := True;
  end;

  Qry2.Close;
  Qry2.Free;
end;

function TfrmESStock.ValidateOrden():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := False;
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  //SQLStr := 'SELECT ITE_Id,Numero FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);
  SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblOrdenes O ' +
            'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
            'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
            'WHERE ITE_Nombre = ' + QuotedStr(gsYear + txtOrden.Text);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    result := True;
  end;

  Qry2.Close;
  Qry2.Free;

end;


function TfrmESStock.GetNumeroDePlano(PlanoId: String):String;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := '';
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT PN_Numero FROM tblPlano WHERE PN_Id = ' + PlanoId;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    result := VarToStr(Qry2['PN_Numero']);
  end;

  Qry2.Close;
  Qry2.Free;
end;

procedure TfrmESStock.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
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
    end;
  end;

end;

procedure TfrmESStock.AddPlanoClick(Sender: TObject);
begin
  Application.CreateForm(TfrmCatalogoPlanos,frmCatalogoPlanos);
  //frmCatalogoPlanos.NuevoClick(nil);
  //frmCatalogoPlanos.txtPlano.Text := txtPlano.Text;
  frmCatalogoPlanos.Show();
end;

procedure TfrmESStock.txtCantidadKeyPress(Sender: TObject; var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
       else
                Key := #0;
end;

end.
