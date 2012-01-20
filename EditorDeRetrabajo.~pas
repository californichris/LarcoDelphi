unit EditorDeRetrabajo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask ,Main,ADODB,DB, Menus,Clipbrd,ComObj,Larco_Functions,StrUtils,
  Chris_Functions, CellEditors,All_Functions,LTCUtils;

type
  TfrmEditorRetrabajo = class(TForm)
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    lblAnio: TLabel;
    lblId: TLabel;
    txtOrden: TMaskEdit;
    txtMotivo: TEdit;
    cmbTareas: TComboBox;
    cmbEmpleados: TComboBox;
    cmbDetectado: TComboBox;
    Primero: TButton;
    Anterior: TButton;
    Siguiente: TButton;
    Ultimo: TButton;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BindEmpleados();
    procedure BindTareas();
    procedure BindData();
    Procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure PrimeroClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure LoadData();
    procedure validarOrden(order: String);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEditorRetrabajo: TfrmEditorRetrabajo;
  giOpcion : Integer;
  gsYear : String;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;
implementation

uses Editor;

{$R *.dfm}

procedure TfrmEditorRetrabajo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmEditorRetrabajo.FormCreate(Sender: TObject);
begin
    lblAnio.Caption := getFormYear(frmMain.sConnString,Self.Name);
    gsYear := RightStr(lblAnio.Caption,2);

    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    LoadData();

    BindEmpleados();
    BindTareas();

    Nuevo.Enabled := True;
    Editar.Enabled := False;
    Borrar.Enabled := False;
    Buscar.Enabled := False;

    if Qry.RecordCount > 0 Then
    begin
          Editar.Enabled := True;
          Borrar.Enabled := True;
          Buscar.Enabled := True;
          BindData();
    end;
  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEditorRetrabajo.BindEmpleados();
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
        cmbEmpleados.Items.Add('000 - Desconocido');
        While not Qry2.Eof do
        Begin
            cmbEmpleados.Items.Add(FormatFloat('000',Qry2['ID']) + ' - ' + Qry2['Nombre']);
            Qry2.Next;
        End;

        cmbEmpleados.Text := '';
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindEmpleados');
    end;

    Qry2.Close;
end;

procedure TfrmEditorRetrabajo.BindTareas();
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

procedure TfrmEditorRetrabajo.BindData();
begin
    if Qry.RecordCount <= 0 Then
    begin
        ClearData();
        Exit;
    end;

    lblId.Caption := VarToStr(Qry['RET_Id']);
    txtOrden.Text := RightStr( VarToStr(Qry['ITE_Nombre']), Length(VarToStr(Qry['ITE_Nombre']))-3 );
    txtMotivo.Text := VarToStr(Qry['RET_Motivo']);
    cmbTareas.Text := VarToStr(Qry['RET_Area']);
    cmbEmpleados.Text := VarToStr(Qry['Responsable']);
    cmbDetectado.Text := VarToStr(Qry['RET_Detectado']);
end;

procedure TfrmEditorRetrabajo.ClearData();
begin
    txtOrden.Text := '';
    txtMotivo.Text := '';
    cmbTareas.Text := '';
    cmbEmpleados.Text := '';
    cmbDetectado.Text := '';
end;

procedure TfrmEditorRetrabajo.EnableControls(Value:Boolean);
begin
    txtOrden.ReadOnly := Value;
    txtMotivo.ReadOnly := Value;

    cmbTareas.Enabled := not Value;
    cmbEmpleados.Enabled := not Value;
    cmbDetectado.Enabled := not Value;

    btnAceptar.Enabled := not Value;
    btnCancelar.Enabled := not Value;
end;

procedure TfrmEditorRetrabajo.PrimeroClick(Sender: TObject);
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


  BindData;
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;

  Nuevo.Enabled := True;
  Editar.Enabled := True;
  Borrar.Enabled := True;
  Buscar.Enabled := True;
  EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEditorRetrabajo.btnCancelarClick(Sender: TObject);
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
BindData();
giOpcion := 0;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEditorRetrabajo.EditarClick(Sender: TObject);
begin
  giOpcion := 2;
  EnableControls(False);
  txtOrden.ReadOnly := True;

  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  txtMotivo.SetFocus;
end;

procedure TfrmEditorRetrabajo.BorrarClick(Sender: TObject);
begin
  giOpcion := 3;
  btnAceptar.Enabled := True;
  btnCancelar.Enabled := True;

  Nuevo.Enabled := False;
  Editar.Enabled := False;
  Buscar.Enabled := False;
end;

procedure TfrmEditorRetrabajo.BuscarClick(Sender: TObject);
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

procedure TfrmEditorRetrabajo.btnAceptarClick(Sender: TObject);
var SQLStr,sOrden : String;
Qry2 : TADOQuery;
begin
  sOrden := '';
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        SQLStr := '';

        conn.Execute(SQLStr);

        Qry.Last;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['RET_Motivo'] := txtMotivo.Text;
        Qry['RET_Area'] := cmbTareas.Text;
        Qry['RET_Empleado'] := LeftStr(cmbEmpleados.Text,3);
        Qry['RET_Detectado'] := cmbDetectado.Text;
        Qry.Post;

        LoadData();
        Qry.Locate('RET_ID',lblId.Caption,[loPartialKey] )
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar la orden en Retrabajo: ' +
                      txtOrden.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
              if MessageDlg('De verdad estas seguro que quieres borrar la orden en Retrabajo: ' +
                            txtOrden.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                        sOrden := txtOrden.Text;
                        Qry.Delete;

                        validarOrden(gsYear + '-' + sOrden);

              end;
        end;
  end
  else if giOpcion = 4 then
  begin
        if txtOrden.Text <> '' then
        begin
              if not Qry.Locate('ITE_Nombre',gsYear + '-' + txtOrden.Text ,[loPartialKey] ) then
                begin
                    MessageDlg('No se encontro ningun Orden con el numero:' + txtOrden.Text + '.', mtInformation,[mbOk], 0);
                    txtOrden.SetFocus;
                    Exit;
                end;
        end
  end;

  EnableControls(True);
  btnAceptar.Enabled := False;
  btnCancelar.Enabled := False;
  Editar.Enabled := False;
  Borrar.Enabled := False;
  Buscar.Enabled := False;
  Nuevo.Enabled := True;
  if Qry.RecordCount > 0 Then
  begin
        Editar.Enabled := True;
        Borrar.Enabled := True;
        Buscar.Enabled := True;
  end;
  BindData();
  giOpcion := 0;
  EnableFormButtons(gbButtons, sPermits);  
end;

function TfrmEditorRetrabajo.ValidateData():Boolean;
var i:Integer;
bfound : boolean;
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

        bfound := False;
        for i:= 0 to cmbTareas.Items.Count do
        begin
                if cmbTareas.Text = cmbTareas.Items[i] then
                begin
                     bfound := True;
                     break;
                end;
        end;

        if bfound = false then
          begin
            MessageDlg('Area Responsable Incorrecta : ' + cmbTareas.Text +
                       '. Seleccionelo de la lista.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        bfound := False;
        for i:= 0 to cmbEmpleados.Items.Count do
        begin
                if cmbEmpleados.Text = cmbEmpleados.Items[i] then
                begin
                     bfound := True;
                     break;
                end;
        end;

        if bfound = false then
          begin
            MessageDlg('Empleado Responsable Incorrecto : ' + cmbEmpleados.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
          end;

        bfound := False;
        for i:= 0 to cmbDetectado.Items.Count do
        begin
                if cmbDetectado.Text = cmbDetectado.Items[i] then
                begin
                     bfound := True;
                     break;
                end;
        end;

        if bfound = false then
          begin
            MessageDlg('Detectado por Incorrecto : ' + cmbDetectado.Text +
                       '. Seleccionelo de la lista.' , mtInformation,[mbOk], 0);
            result :=  False;
          end;

end;

procedure TfrmEditorRetrabajo.LoadData();
begin
    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT R.*,' +
                    'CASE WHEN R.RET_Empleado IS NULL THEN ''000 - Desconocido'' ELSE ' +
                    'R.RET_Empleado + '' - '' + E.Nombre END AS Responsable ' +
                    'FROM tblRetrabajo R ' +
                    'LEFT OUTER JOIN tblEmpleados E ON R.RET_Empleado = E.ID ' +
                    'WHERE Left(R.ITE_Nombre,2) = ' + QuotedStr(gsYear) + ' ' +
                    'ORDER BY R.ITE_Nombre ';

    Qry.Open;


end;

procedure TfrmEditorRetrabajo.validarOrden(order: String);
var SQLStr : String;
Qry2 : TADOQuery;
begin
  SQLStr := 'SELECT * FROM tblItemTasks WHERE ITE_Nombre = ' +
            QuotedStr(order) + ' AND TAS_ID = 18 AND ITS_Status = 3';

  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
        ShowMessage('La orden esta en retrabajo Calidad, sera puesta en listo Calidad.');

        SQLStr := 'UPDATE tblItemTasks SET ITS_Status = 0 WHERE ITE_Nombre = ' +
                  QuotedStr(order) + ' AND TAS_ID = 18 AND ITS_Status = 3';

        Conn.Execute(SQLStr);
  end;
end;

end.
