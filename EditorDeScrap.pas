unit EditorDeScrap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask ,Main,ADODB,DB, Menus,Clipbrd,ComObj,Larco_Functions,StrUtils,
  Chris_Functions, CellEditors,All_Functions,LTCUtils;
type
  TfrmScrapEditor = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    txtOrden: TMaskEdit;
    Label2: TLabel;
    txtMotivo: TEdit;
    Label3: TLabel;
    cmbTareas: TComboBox;
    Label4: TLabel;
    cmbEmpleados: TComboBox;
    Label5: TLabel;
    cmbDetectado: TComboBox;
    lblCantidad: TLabel;
    txtCantidad: TEdit;
    chkParcial: TCheckBox;
    lblRepro: TLabel;
    txtRepro: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    txtNuevaOrden: TMaskEdit;
    Label8: TLabel;
    cmbUsuario: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    lblAnio: TLabel;
    lblId: TLabel;
    deFecha: TDateEditor;
    Label9: TLabel;
    cmbEmpleadoDetecto: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindEmpleados();
    procedure BindTareas();
    procedure BindScrap();
    Procedure ClearData();
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure NuevoClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure txtReproKeyPress(Sender: TObject; var Key: Char);
    function ValidateData():Boolean;
    function ValidateOrden(Orden: String):Boolean;
    function ValidateScrap(Orden: String):Boolean;
    procedure btnAceptarClick(Sender: TObject);
    function BoolToStrInt(Value:Boolean):String;
    procedure LoadScrap();
    function ValidarCantidad(Item:String;Cantidad:Integer):Boolean;
    function FormIsRunning(FormName: String):Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScrapEditor: TfrmScrapEditor;
  giOpcion : Integer;
  gsYear : String;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Editor;

{$R *.dfm}

procedure TfrmScrapEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmScrapEditor.BindEmpleados();
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT ID,Nombre FROM tblEmpleados Order By Nombre';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        cmbEmpleados.Items.Clear;
        cmbUsuario.Items.Clear;
        cmbEmpleados.Items.Add('000 - Desconocido');
        cmbEmpleadoDetecto.Items.Add('000 - Desconocido');
        While not Qry2.Eof do
        Begin
            cmbEmpleados.Items.Add(FormatFloat('000',Qry2['ID']) + ' - ' + Qry2['Nombre']);
            cmbEmpleadoDetecto.Items.Add(FormatFloat('000',Qry2['ID']) + ' - ' + Qry2['Nombre']);
            cmbUsuario.Items.Add(FormatFloat('000',Qry2['ID']) + ' - ' + Qry2['Nombre']);
            Qry2.Next;
        End;

        cmbUsuario.Text := '';
        cmbEmpleados.Text := '';
        cmbEmpleadoDetecto.Text := '';
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindEmpleados');
    end;

    Qry2.Close;
end;

procedure TfrmScrapEditor.BindTareas();
var Qry2 : TADOQuery;
SQLStr : String;
begin

    Qry2 := nil;
    try
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        cmbTareas.Items.Clear;
        cmbDetectado.Items.Clear;
        While not Qry2.Eof do
        Begin
            cmbTareas.Items.Add(Qry2['Nombre']);
            cmbDetectado.Items.Add(Qry2['Nombre']);
            Qry2.Next;
        End;

        cmbTareas.Text := '';
        cmbDetectado.Text := '';

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindTareas');
    end;

    Qry2.Close;
end;



procedure TfrmScrapEditor.FormCreate(Sender: TObject);
begin
  lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);
  gsYear := RightStr(lblAnio.Caption,2);

  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  LoadScrap();

  BindEmpleados();
  BindTareas();
  if Qry.RecordCount > 0 then
      BindScrap()
  else
  begin
      Editar.Enabled := False;
      Borrar.Enabled := False;
      Buscar.Enabled := False;
  end;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmScrapEditor.BindScrap();
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['SCR_Id']);
    txtOrden.Text := RightStr( VarToStr(Qry['ITE_Nombre']), Length(VarToStr(Qry['ITE_Nombre']))-3 );
    txtMotivo.Text := VarToStr(Qry['SCR_Motivo']);
    cmbTareas.Text := VarToStr(Qry['SCR_Tarea']);
    cmbEmpleados.Text := VarToStr(Qry['Responsable']);
    cmbDetectado.Text := VarToStr(Qry['SCR_Detectado']);
    cmbUsuario.Text := VarToStr(Qry['Usuario']);
    cmbEmpleadoDetecto.Text := VarToStr(Qry['Detectado']);
    txtCantidad.Text := VarToStr(Qry['SCR_Cantidad']);
    txtRepro.Text := VarToStr(Qry['SCR_Repro']);
    deFecha.Text := VarToStr(Qry['Fecha']);
    txtNuevaOrden.Text := RightStr( VarToStr(Qry['SCR_NewItem']), Length(VarToStr(Qry['SCR_NewItem']))-3 );
    chkParcial.Checked := StrToBool(VarToStr(Qry['SCR_Parcial']));
end;

procedure TfrmScrapEditor.ClearData();
begin
    txtOrden.Text := '';
    txtMotivo.Text := '';
    cmbTareas.Text := '';
    cmbEmpleados.Text := '';
    cmbEmpleadoDetecto.Text := '';
    cmbDetectado.Text := '';
    cmbUsuario.Text := '';
    txtCantidad.Text := '';
    txtRepro.Text := '';
    deFecha.Text := DateToStr(Now);
    txtNuevaOrden.Text := '';
    chkParcial.Checked := false;
end;

procedure TfrmScrapEditor.Button1Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.First;
BindScrap;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmScrapEditor.Button2Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Prior;
BindScrap;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmScrapEditor.Button3Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Next;
BindScrap;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmScrapEditor.Button4Click(Sender: TObject);
begin
if Qry.RecordCount = 0 then
        Exit;

Qry.Last;
BindScrap;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmScrapEditor.EnableControls(Value:Boolean);
begin
    txtOrden.ReadOnly := Value;
    txtMotivo.ReadOnly := Value;
    txtCantidad.ReadOnly := Value;
    txtRepro.ReadOnly := Value;
    txtNuevaOrden.ReadOnly := Value;

    deFecha.Enabled := not Value;
    cmbTareas.Enabled := not Value;
    cmbEmpleados.Enabled := not Value;
    cmbEmpleadoDetecto.Enabled := not Value;
    cmbDetectado.Enabled := not Value;
    cmbUsuario.Enabled := not Value;
    chkParcial.Enabled := not Value;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;

procedure TfrmScrapEditor.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtOrden.SetFocus;
giOpcion := 1;
txtNuevaOrden.ReadOnly := True;
deFecha.Text := DateToStr(Now);

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmScrapEditor.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);
txtNuevaOrden.ReadOnly := True;
txtOrden.ReadOnly := True;

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
txtOrden.SetFocus;
end;

procedure TfrmScrapEditor.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmScrapEditor.BuscarClick(Sender: TObject);
begin
ClearData();
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;
txtOrden.ReadOnly := False;
txtOrden.SetFocus;
giOpcion := 4;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
end;

procedure TfrmScrapEditor.btnCancelarClick(Sender: TObject);
begin
ClearData();
EnableControls(True);

Nuevo.Enabled := True;
if Qry.RecordCount > 0 Then
begin
      Editar.Enabled := True;
      Borrar.Enabled := True;
      Buscar.Enabled := True;
end;
BindScrap();
giOpcion := 0;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmScrapEditor.txtReproKeyPress(Sender: TObject; var Key: Char);
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

function TfrmScrapEditor.ValidateData():Boolean;
begin
        result := True;

        if UT(txtMotivo.Text) = '' then
          begin
            MessageDlg('Por favor escriba un motivo.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if UT(cmbTareas.Text) = '' then
          begin
            MessageDlg('Por favor Seleccione una Area responsable.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if UT(cmbEmpleados.Text) = '' then
          begin
            MessageDlg('Por favor Seleccione un Empleado responsable.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if UT(cmbDetectado.Text) = '' then
          begin
            MessageDlg('Por favor Seleccione una Area de la lista detectado por.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if UT(cmbEmpleadoDetecto.Text) = '' then
          begin
            MessageDlg('Por favor Seleccione el Empleado que lo detecto.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if UT(cmbUsuario.Text) = '' then
          begin
            MessageDlg('Por favor Seleccione un Empleado de la lista scrapeado por.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if (UT(txtCantidad.Text) = '') or (UT(txtCantidad.Text) = '0') then
          begin
            MessageDlg('La cantidad no puede estar vacia o igual a cero.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if (cmbTareas.Items.IndexOf(cmbTareas.Text) = -1) then begin
            MessageDlg('Area Responsable Incorrecta : ' + cmbTareas.Text +
                       '. Seleccionelo de la lista.', mtInformation,[mbOk], 0);
            result :=  False;
        end;

        if (cmbEmpleados.Items.IndexOf(cmbEmpleados.Text) = -1) then begin
            MessageDlg('Empleado Responsable Incorrecto : ' + cmbEmpleados.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
        end;

        if (cmbDetectado.Items.IndexOf(cmbDetectado.Text) = -1) then begin
            MessageDlg('Area Detectado Incorrecto : ' + cmbDetectado.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
        end;

        if (cmbEmpleadoDetecto.Items.IndexOf(cmbEmpleadoDetecto.Text) = -1) then begin
            MessageDlg('Empleado que lo detecto Incorrecto : ' + cmbEmpleadoDetecto.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
        end;

        if (cmbUsuario.Items.IndexOf(cmbUsuario.Text) = -1) then begin
            MessageDlg('Scrapeado por Incorrecto : ' + cmbUsuario.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
        end;

        if not IsDate(deFecha.Text) then
          begin
            MessageDlg('Fecha Incorrecta: ' + deFecha.Text, mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtCantidad.Text) Then
          begin
            MessageDlg('El cantidad debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        txtRepro.Text := '0';
        if chkParcial.Checked then
        begin
              if txtRepro.Text = '' then
              begin
                MessageDlg('Por favor capture una cantidad a reprogramar.', mtInformation,[mbOk], 0);
                Exit;
              end;

              if not IsNumeric(txtRepro.Text) Then
                begin
                  MessageDlg('La cantidad a reprogramar debe de ser un valor numerico.', mtInformation,[mbOk], 0);
                  result :=  False;
                end;
        end;

end;

procedure TfrmScrapEditor.btnAceptarClick(Sender: TObject);
var SQLStr,sID,sOrden : String;
Qry2 : TADOQuery;
bFound : boolean;
begin
  sOrden := '';
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        if not ValidateOrden(gsYear + '-' + txtOrden.Text) then
          begin
            MessageDlg('La Orden de Trabajo : ' + txtOrden.Text + ' no existe.', mtInformation,[mbOk], 0);
            Exit;
          end;

        if not ValidateScrap(gsYear + '-' + txtOrden.Text) then
          begin
            MessageDlg('La Orden de Trabajo : ' + txtOrden.Text + ' ya fue declara como scrap.', mtInformation,[mbOk], 0);
            Exit;
          end;

        if not ValidarCantidad(gsYear + '-' + txtOrden.Text, StrToInt(txtCantidad.Text) ) then
        begin
          MessageDlg('La cantidad scrapeada es mayor que la cantidad de la orden.', mtInformation,[mbOk], 0);
          Exit;
        end;

        SQLStr := 'INSERT INTO tblScrap(ITE_Nombre,SCR_Motivo,SCR_Tarea,SCR_EmpleadoRes,SCR_Cantidad,' +
                  'SCR_Parcial,SCR_Repro,USE_Login,SCR_Fecha,SCR_NewItem,SCR_Impreso,SCR_Activo,SCR_Detectado, ' +
                  'SCR_EmpleadoDetectado,Update_Date,Update_User) ' +
                  'VALUES(' + QuotedStr(gsYear + '-' + txtOrden.Text) + ',' +
                  QuotedStr(txtMotivo.Text) + ',' + QuotedStr(cmbTareas.Text) + ',' +
                  QuotedStr(LeftStr(cmbEmpleados.Text,3)) + ',' + txtCantidad.Text + ',' +
                  BoolToStrInt(chkParcial.Checked) + ',' + txtRepro.Text +
                  ',' + QuotedStr(LeftStr(cmbUsuario.Text,3)) + ',' +
                  QuotedStr(deFecha.Text) + ',NULL,0,0,' +
                  QuotedStr(cmbDetectado.Text) + ',' + LeftStr(cmbEmpleadoDetecto.Text,3) + ',GETDATE(),' + frmMain.sUserLogin + ')';

        conn.Execute(SQLStr);

        LoadScrap();

        Qry.Last;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        if not ValidarCantidad(gsYear + '-' + txtOrden.Text, StrToInt(txtCantidad.Text) ) then
        begin
          MessageDlg('La cantidad scrapeada es mayor que la cantidad de la orden.', mtInformation,[mbOk], 0);
          Exit;
        end;

        SQLStr := 'UPDATE tblScrap SET SCR_Motivo = ' +  QuotedStr(txtMotivo.Text) +
                  ',SCR_Tarea = ' + QuotedStr(cmbTareas.Text) +
                  ',SCR_EmpleadoRes = ' + QuotedStr(LeftStr(cmbEmpleados.Text,3)) +
                  ',SCR_Cantidad = ' + txtCantidad.Text +
                  ',SCR_Parcial = ' + BoolToStrInt(chkParcial.Checked) +
                  ',SCR_Repro = ' + txtRepro.Text +
                  ',USE_Login = ' + QuotedStr(LeftStr(cmbUsuario.Text,3)) +
                  ',SCR_Fecha = ' + QuotedStr(deFecha.Text) +
                  ',SCR_Detectado = ' + QuotedStr(cmbDetectado.Text) +
                  ',SCR_EmpleadoDetectado = ' + LeftStr(cmbEmpleadoDetecto.Text, 3) +
                  ',Update_Date = GETDATE() ' +
                  ',Update_User = ' + frmMain.sUserLogin +
                  ' WHERE SCR_ID = ' + lblID.Caption;

        conn.Execute(SQLStr);

        LoadScrap();
        Qry.Locate('SCR_ID',lblId.Caption,[loPartialKey] )

  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar la orden : ' +
                      txtOrden.Text + ' que fue marcada como scrap?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar la orden : ' +
                            txtOrden.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        sOrden := txtOrden.Text;
                        Qry.Delete;
              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        sID := lblId.Caption;
        Qry.First;
        bFound := False;
        while not Qry.Eof do
        begin
                if VarToStr(Qry['ITE_Nombre']) = gsYear + '-' + txtOrden.Text then
                begin
                        BindScrap;
                        bFound := True;
                        Break;
                end;
                Qry.Next;
        end;

        if bFound = False then
          begin
              MessageDlg('No se encontro ningun Orden con estos datos.', mtInformation,[mbOk], 0);
              txtOrden.SetFocus;

              Qry.First;
              while not Qry.Eof do
              begin
                      if VarToStr(Qry['SCR_ID']) = sID then
                      begin
                              Break;
                      end;
                      Qry.Next;
              end;
              
              Exit;
          end;



  end;

EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;
Nuevo.Enabled := True;
if Qry.RecordCount > 0 Then
begin
      Editar.Enabled := True;
      Borrar.Enabled := True;
      Buscar.Enabled := True;

end;
BindScrap;
EnableFormButtons(gbButtons, sPermits);
giOpcion := 0;

if sOrden <> '' then
begin
        if MessageDlg('La orden : ' +
                      sOrden + ' se borro exitosamente deseas moverla de tarea?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then

        begin
            if FormIsRunning('frmEditor') Then
              begin
                    setActiveWindow(frmEditor.Handle);
                    frmEditor.WindowState := wsNormal;
                    frmEditor.lblAnio.Caption := lblAnio.Caption;
                    frmEditor.gsYear := RightStr(lblAnio.Caption,2) + '-';
                    frmEditor.txtOrden.Text := sOrden;
                    frmEditor.btnBuscarClick(nil);
              end
            else
              begin
                    Application.CreateForm(TfrmEditor,frmEditor);
                    frmEditor.Show;
                    frmEditor.lblAnio.Caption := lblAnio.Caption;
                    frmEditor.gsYear := RightStr(lblAnio.Caption,2) + '-';
                    frmEditor.txtOrden.Text := sOrden;
                    frmEditor.btnBuscarClick(nil);
              end;

        end;
end;

end;

function TfrmScrapEditor.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;


procedure TfrmScrapEditor.LoadScrap();
begin
    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT S.*,Convert(varchar(10),SCR_Fecha,101) AS Fecha, ' +
                    'CASE WHEN E.Nombre IS NULL THEN ''000 - Desconocido'' ELSE ' +
                    'Right(''000'' + S.SCR_EmpleadoRes, 3) + '' - '' + E.Nombre END AS Responsable, ' +
                    'CASE WHEN U.Nombre IS NULL THEN ''000 - Desconocido'' ELSE ' +
                    'Right(''000'' + S.USE_Login, 3) + '' - '' + U.Nombre END AS Usuario, ' +
                    'CASE WHEN D.Nombre IS NULL THEN ''000 - Desconocido'' ELSE Right(''000'' + CAST(S.SCR_EmpleadoDetectado AS Varchar(3)), 3) + '' - '' + D.Nombre END AS Detectado ' +
                    'FROM tblScrap S ' +
                    'LEFT OUTER JOIN tblEmpleados E ON S.SCR_EmpleadoRes = E.ID ' +
                    'LEFT OUTER JOIN tblEmpleados U ON S.USE_Login = U.ID ' +
                    'LEFT OUTER JOIN tblEmpleados D ON S.SCR_EmpleadoDetectado = D.ID ' +
                    'WHERE Left(ITE_Nombre,2) = ' + QuotedStr(gsYear) + ' ' +
                    'ORDER BY S.SCR_ID ';
    Qry.Open;

end;

function TfrmScrapEditor.ValidateOrden(Orden: String):Boolean;
var Qry2 : TADOQuery;
begin
    Result := False;
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := 'SELECT ITE_Nombre FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(Orden);
    Qry2.Open;

    If Qry2.RecordCount > 0 Then
        Result := True;

    Qry.Close;
end;

function TfrmScrapEditor.ValidateScrap(Orden: String):Boolean;
var Qry2 : TADOQuery;
begin
    Result := True;
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := 'SELECT ITE_Nombre FROM tblScrap WHERE ITE_Nombre = ' + QuotedStr(Orden);
    Qry2.Open;

    If Qry2.RecordCount > 0 Then
        Result := False;

    Qry.Close;
end;

function TfrmScrapEditor.ValidarCantidad(Item:String;Cantidad:Integer):Boolean;
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Result := True;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Top 1 Ordenada FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(Item);

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    if Qry2.RecordCount > 0 then
        if Cantidad > StrToInt(VarToStr(Qry2['Ordenada'])) then
                Result := False;

    Qry2.Close;
end;

function TfrmScrapEditor.FormIsRunning(FormName: String):Boolean;
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
end.
