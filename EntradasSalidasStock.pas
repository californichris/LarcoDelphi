unit EntradasSalidasStock;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,
  CellEditors, Larco_functions, Math, Mask, Menus, ComObj,Clipbrd;

type
  TfrmESStock = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lblId: TLabel;
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
    txtNumero: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    gbBuscar: TGroupBox;
    Label8: TLabel;
    txtBuscarPlano: TEdit;
    Button1: TButton;
    gvResults: TGridView;
    Button2: TButton;
    lblTotal: TLabel;
    PopupMenu2: TPopupMenu;
    MenuItem1: TMenuItem;
    CopiarOrden1: TMenuItem;
    SaveDialog1: TSaveDialog;
    PopupMenu1: TPopupMenu;
    Copiar1: TMenuItem;
    Pegar1: TMenuItem;
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
    function ValidateOrdenInStock():Boolean;
    function ValidateExistencia():Boolean;
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PrimeroKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure txtBuscarPlanoExit(Sender: TObject);
    procedure txtBuscarPlanoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button2Click(Sender: TObject);
    procedure gvResultsDblClick(Sender: TObject);
    function FormIsRunning(FormName: String):Boolean;
    procedure MenuItem1Click(Sender: TObject);
    procedure ExportGrid(Grid: TGridView;sFileName: String);
    procedure CopiarOrden1Click(Sender: TObject);
    procedure txtPlanoKeyPress(Sender: TObject; var Key: Char);
    procedure Copiar1Click(Sender: TObject);
    procedure Pegar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  type
    TStock = class(TObject)
    id: String;
    plano: String;
    anio: String;
    orden: String;
    fecha: String;
    cantidad: String;
    noParte: String;
    tipo: String;
  end;

var
  frmESStock: TfrmESStock;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
  giOpcion : Integer;  //0= nada, 1 Nuevo, 2 editar, 3 borrar, 4 buscar
  giStock : TStock;

implementation

uses Main, CatalogoPlanos, CatalogoPlanosModal;

{$R *.dfm}

procedure TfrmESStock.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  CloseConns(Qry, Conn);
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
  gbButtons.Height := 0;
  gbBuscar.Height := 302;

  txtBuscarPlano.Text := '';
  gvResults.ClearRows;
  txtBuscarPlano.SetFocus();

  giOpcion := 4;
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
  cmbTipo.Text := 'Entrada';
  lblValidOrden.Caption := '';
  lblPNId.Caption := '';
  txtNumero.Text := '';
  ddlAnio.Text := '';
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
var year : String;
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

    year := '20' +  LeftStr(VarToStr(Qry['ITE_Nombre']), 2);
    ddlAnio.ItemIndex := ddlAnio.Items.IndexOf(year);
    ddlAnio.Text := year;
    ValidateOrden();

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
var year: String;
iYear, i : integer;
begin
  ddlAnio.Clear;
  year := FormatDateTime( 'yyyy', Now);
  iYear := StrToInt(year);
  for i := 2000 to iYear do begin
    ddlAnio.Items.Add(IntToStr(i));
  end;

  ddlAnio.ItemIndex := ddlAnio.Items.IndexOf(getFormYear(frmMain.sConnString, Self.Name));

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;
  Qry := TADOQuery.Create(nil);
  Qry.Connection := Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := 'SELECT * FROM tblStock ORDER BY ST_Fecha Desc, ST_ID Desc';
  Qry.Open;

  if Qry.RecordCount > 0 then begin
    BindData();
  end;

  giOpcion := 0;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);

  EnableControls(True);
  EnableButtons();
  lblTotal.Caption := '';

  giStock := TStock.Create;
end;

procedure TfrmESStock.txtPlanoExit(Sender: TObject);
var found : boolean;
begin
  if (giOpcion = 0) or (Trim(txtPlano.Text) = '') then
    Exit;

  txtPlano.Text := UpperCase(Trim(txtPlano.Text));
  found := ValidatePlano();

  if not found then begin
    if MessageDlg('El Numero de Plano no existe, deseas agregarlo?',
              mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if FormIsRunning('frmCatalogoPlanos') Then
      begin
        setActiveWindow(frmCatalogoPlanos.Handle);
      end
      else begin
        Application.CreateForm(TfrmCatalogoPlanos, frmCatalogoPlanos);
      end;
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

  Qry2 := nil;
  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblOrdenes O ' +
              'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
              'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
              'WHERE ITE_Nombre = ' + QuotedStr(RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text);


    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    lblValidOrden.Caption := '';

    if Qry2.RecordCount > 0 then begin
      lblValidOrden.Caption := VarToStr(Qry2['ITE_Id']);
      sNoParte := VarToStr(Qry2['Numero']);
      sAlias := VarToStr(Qry2['PA_Alias']);
      sPlano := VarToStr(Qry2['PN_Numero']);
      txtNumero.Text := sNoParte;

      //Removed validation ask by Edgar. The Part Number does not need to be an alias of a blueprint we already
      //have the link using the ITE_Nombre.

      {if (Trim(sNoParte) <> Trim(txtPlano.Text)) and (sAlias = '') then begin
        if MessageDlg('El Numero de Parte [' + sNoParte + '] de esta orden no es un Nombre Interno o Alias de este numero de Plano, deseas agregarlo?',
                  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          if FormIsRunning('frmCatalogoPlanos') Then
          begin
            setActiveWindow(frmCatalogoPlanos.Handle);
          end
          else begin
            Application.CreateForm(TfrmCatalogoPlanos, frmCatalogoPlanos);
          end;
          if lblPNId.Caption <> '' then begin
            frmCatalogoPlanos.EditarPlano(lblPNId.Caption);
            frmCatalogoPlanos.txtInterno.Text := sNoParte;
            frmCatalogoPlanos.txtAlias.Text := sNoParte;
          end;

          frmCatalogoPlanos.Show();

        end;
      end;}

    end
    else begin
      if  cmbTipo.Text = 'Salida' then begin // checar en stock
        SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblStockOrdenes O ' +
                  'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
                  'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
                  'WHERE ITE_Nombre = ' + QuotedStr(RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text);


        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        lblValidOrden.Caption := '';

        if Qry2.RecordCount > 0 then begin
          lblValidOrden.Caption := VarToStr(Qry2['ITE_Id']);
          sNoParte := VarToStr(Qry2['Numero']);
          sAlias := VarToStr(Qry2['PA_Alias']);
          sPlano := VarToStr(Qry2['PN_Numero']);
          txtNumero.Text := sNoParte;
        end
        else begin
          ShowMessage('La Orden de Trabajo no es valida.');
        end;

      end
      else begin
        ShowMessage('La Orden de Trabajo no es valida.');
      end;
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;

end;

procedure TfrmESStock.btnAceptarClick(Sender: TObject);
begin
  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['PN_Id'] := lblPNId.Caption;
        Qry['ITE_Nombre'] := RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text;
        Qry['ST_Cantidad'] := txtCantidad.Text;
        Qry['ST_Fecha'] := deFecha.Text;
        Qry['ST_Tipo'] := cmbTipo.Text;
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
        Qry['ITE_Nombre'] := RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text;
        Qry['ST_Cantidad'] := txtCantidad.Text;
        Qry['ST_Fecha'] := deFecha.Text;
        Qry['ST_Tipo'] := cmbTipo.Text;
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
  txtPlano.Text := UpperCase(Trim(txtPlano.Text));

  ValidateData := True;
  sOrden := Trim( StringReplace(txtOrden.Text,'-','',[rfReplaceAll, rfIgnoreCase]) );
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

  if not ValidateOrdenInStock() then
  begin
    MessageDlg('Ya existe una ' +  cmbTipo.Text + ' para esta Orden de Trabajo.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
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

  if (cmbTipo.Text = 'Salida') then begin
    if not ValidateExistencia() then begin
      MessageDlg('No hay piezas suficientes en existencia para registrar esta Salida.', mtInformation,[mbOk], 0);
      result :=  False;
      Exit;
    end;
  end;

  if Trim(deFecha.Text) = '' then
  begin
    MessageDlg('La Fecha es requerida.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;

  if deFecha.Date > Now then
  begin
    MessageDlg('La Fecha no puede ser mayor que el dia de hoy.', mtInformation,[mbOk], 0);
    result :=  False;
    Exit;
  end;


  if Trim(cmbTipo.Text) = '' Then
  begin
    MessageDlg('Por favor seleccione el Tipo.', mtInformation,[mbOk], 0);
    result :=  False;
  end;

  if cmbTipo.Items.IndexOf(cmbTipo.Text) = -1 then
  begin
      MessageDlg('Tipo incorrecto, por favor seleccionelo de la lista.', mtInformation,[mbOk], 0);
      result :=  False;
  end;

end;

function TfrmESStock.ValidateExistencia():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
exist : Integer;
sExist : String;
begin
  result := False;
  Qry2 := nil;

  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT SUM(CASE WHEN S.ST_Tipo = ''Entrada'' THEN S.ST_Cantidad ELSE 0 END) - ' +
              'SUM(CASE WHEN S.ST_Tipo = ''Salida'' THEN S.ST_Cantidad ELSE 0 END) AS Existencia  ' +
              'FROM tblStock S WHERE S.PN_Id = ' + lblPNId.Caption;

    if giOpcion = 2 then begin
      SQLStr := SQLStr + ' AND S.ST_ID <> ' + lblId.Caption;
    end;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    exist := 0;
    if Qry2.RecordCount > 0 then begin
      sExist := VarToStr(Qry2['Existencia']);
      if '' <> sExist then begin
        exist := StrToInt(sExist);
      end;
    end;

    if StrToInt(txtCantidad.Text) <= exist then begin
      result := True;
    end;

  end
  finally
    CloseConns(Qry2, nil);
  end;

end;


function TfrmESStock.ValidatePlano():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := False;
  Qry2 := nil;

  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT PN_Id FROM tblPlano WHERE PN_Numero = ' + QuotedStr(txtPlano.Text);

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    lblPNId.Caption := '';
    if Qry2.RecordCount > 0 then begin
      lblPNId.Caption := VarToStr(Qry2['PN_Id']);
      result := True;
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;
  
end;

function TfrmESStock.ValidateOrdenInStock():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := True;
  Qry2 := nil;
  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT ST_ID FROM tblStock WHERE ITE_Nombre = ' + QuotedStr(RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text) +
              ' AND ST_Tipo = ' + QuotedStr(cmbTipo.Text);

    if giOpcion = 2 then begin
     SQLStr := SQLStr + ' AND ST_ID <> ' + lblId.Caption;
    end;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then begin
      result := False;
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;
end;


function TfrmESStock.ValidateOrden():Boolean;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := False;

  Qry2 := nil;
  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblOrdenes O ' +
              'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
              'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
              'WHERE ITE_Nombre = ' + QuotedStr(RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text);

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then begin
      txtNumero.Text := VarToStr(Qry2['Numero']);
      lblValidOrden.Caption := VarToStr(Qry2['ITE_Id']);
      result := True;
    end
    else begin
      if  cmbTipo.Text = 'Salida' then begin // checar en stock
        SQLStr := 'SELECT ITE_ID,ITE_Nombre,O.Numero,PA.*,P.* FROM tblStockOrdenes O ' +
                  'LEFT OUTER JOIN tblPlanoAlias PA ON O.Numero = PA.PA_Alias ' +
                  'LEFT OUTER JOIN tblPlano P ON PA.PN_Id = P.PN_Id ' +
                  'WHERE ITE_Nombre = ' + QuotedStr(RightStr(ddlAnio.Text, 2) + '-' + txtOrden.Text);


        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        lblValidOrden.Caption := '';

        if Qry2.RecordCount > 0 then begin
          txtNumero.Text := VarToStr(Qry2['Numero']);
          lblValidOrden.Caption := VarToStr(Qry2['ITE_Id']);
          result := True;
        end;
      end;
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;
end;


function TfrmESStock.GetNumeroDePlano(PlanoId: String):String;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := '';
  Qry2 := nil;
  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT PN_Numero FROM tblPlano WHERE PN_Id = ' + PlanoId;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then begin
      result := VarToStr(Qry2['PN_Numero']);
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;
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
  if FormIsRunning('frmCatalogoPlanos') Then
  begin
    setActiveWindow(frmCatalogoPlanos.Handle);
  end
  else begin
    Application.CreateForm(TfrmCatalogoPlanos, frmCatalogoPlanos);
  end;

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

procedure TfrmESStock.FormKeyDown(Sender: TObject; var Key: Word;
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

procedure TfrmESStock.PrimeroKeyDown(Sender: TObject; var Key: Word;
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

procedure TfrmESStock.Button1Click(Sender: TObject);
var SQLStr, SQLWhere, year: String;
Qry2 : TADOQuery;
entradas, salidas : integer;
doCount : boolean;
begin
  gvResults.ClearRows;
  lblTotal.Caption := '';
  if Trim(txtBuscarPlano.Text) = '' then begin
    ShowMessage('Numero de Plano es requerido.');
    Exit;
  end;

  txtBuscarPlano.Text := UpperCase(Trim(txtBuscarPlano.Text));
  SQLWhere := txtBuscarPlano.Text;

  Qry2 := nil;
  try
  begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection := Conn;

    SQLStr := 'SELECT * FROM tblStock S INNER JOIN tblPlano P ON S.PN_Id = P.PN_Id WHERE P.PN_Numero ';
    if Pos('*', txtBuscarPlano.Text) <> 0 then begin
      SQLWhere := ' LIKE ' + QuotedStr(StringReplace(SQLWhere, '*', '%', [rfReplaceAll, rfIgnoreCase]));
      doCount := false;
    end
    else begin
      SQLWhere := ' = ' + QuotedStr(SQLWhere);
      doCount := true;
    end;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr + SQLWhere;
    Qry2.Open;

    entradas := 0;
    salidas := 0;
    While not Qry2.Eof do
    Begin
        gvResults.AddRow(1);
        gvResults.Cells[0,gvResults.RowCount -1] := VarToStr(Qry2['ST_Id']);
        gvResults.Cells[1,gvResults.RowCount -1] := VarToStr(Qry2['PN_Numero']);
        year := '20' +  LeftStr(VarToStr(Qry2['ITE_Nombre']), 2);
        gvResults.Cells[2,gvResults.RowCount -1] := year;
        gvResults.Cells[3,gvResults.RowCount -1] := VarToStr(Qry2['ST_Fecha']);
        gvResults.Cells[4,gvResults.RowCount -1] := RightStr( VarToStr(Qry2['ITE_Nombre']), Length(VarToStr(Qry2['ITE_Nombre']))-3 );
        gvResults.Cells[5,gvResults.RowCount -1] := VarToStr(Qry2['ST_Tipo']);
        gvResults.Cells[6,gvResults.RowCount -1] := VarToStr(Qry2['ST_Cantidad']);

        if doCount then begin
          if 'Entrada' = VarToStr(Qry2['ST_Tipo']) then begin
            entradas := entradas + StrToInt(VarToStr(Qry2['ST_Cantidad']));
          end
          else begin
            salidas := salidas + StrToInt(VarToStr(Qry2['ST_Cantidad']));
          end;
          lblTotal.Caption := 'En Stock : '+ IntToStr(entradas - salidas);
        end;

        Qry2.Next;
    end;
  end
  finally
    CloseConns(Qry2, nil);
  end;

end;

procedure TfrmESStock.txtBuscarPlanoExit(Sender: TObject);
begin
  txtBuscarPlano.Text := UpperCase(Trim(txtBuscarPlano.Text));
end;

procedure TfrmESStock.txtBuscarPlanoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = vk_return then
  begin
    Button1Click(nil);
  end

end;

procedure TfrmESStock.Button2Click(Sender: TObject);
begin
  gbButtons.Height := 302;
  gbBuscar.Height := 0;
end;

procedure TfrmESStock.gvResultsDblClick(Sender: TObject);
var id : String;
begin
  id := gvResults.Cell[0,gvResults.SelectedRow].AsString;
  Qry.Locate('ST_Id', id, [loPartialKey] );

  giOpcion := 0;
  ClearData();
  BindData();
  //EnableButtons();

  gbButtons.Height := 302;
  gbBuscar.Height := 0;
end;

function TfrmESStock.FormIsRunning(FormName: String):Boolean;
var i:Integer;
begin
  Result := False;

  for  i := 0 to Screen.FormCount - 1 do
  begin
        if Screen.Forms[i].Name = FormName Then
          begin
                Result:= True;
                Break;
          end;
  end;

end;

procedure TfrmESStock.MenuItem1Click(Sender: TObject);
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

procedure TfrmESStock.ExportGrid(Grid: TGridView;sFileName: String);
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

procedure TfrmESStock.CopiarOrden1Click(Sender: TObject);
begin
  Clipboard.AsText := gvResults.Cells[1,gvResults.SelectedRow];
end;

procedure TfrmESStock.txtPlanoKeyPress(Sender: TObject; var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmESStock.Copiar1Click(Sender: TObject);
begin
  giStock.id := '';
  giStock.plano := lblPNId.Caption;
  giStock.anio := ddlAnio.Text;
  giStock.orden := txtOrden.Text;
  giStock.fecha := deFecha.Text;
  giStock.cantidad := txtCantidad.Text;
  giStock.noParte := txtNumero.Text;
  giStock.tipo := cmbTipo.Text;
end;

procedure TfrmESStock.Pegar1Click(Sender: TObject);
begin
  if giOpcion <> 1 then begin
    ShowMessage('Solo se puede pegar cuando se esta creando una nueva entrada/Salida');
    Exit;
  end;

  lblPNId.Caption := giStock.plano;
  txtPlano.Text := GetNumeroDePlano(lblPNId.Caption);
  txtOrden.Text := giStock.orden;
  deFecha.Text := giStock.fecha;
  txtCantidad.Text := giStock.cantidad;
  cmbTipo.Text := giStock.tipo;

  ddlAnio.ItemIndex := ddlAnio.Items.IndexOf(giStock.anio);
  ddlAnio.Text := giStock.anio;
  ValidateOrden();
end;

end.
