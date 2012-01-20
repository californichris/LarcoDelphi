unit Empleados;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBCtrls, StdCtrls, ComCtrls, ADODB,DB,IniFiles,All_Functions,chris_Functions,
  ScrollView, CustomGridViewControl, CustomGridView, GridView,StrUtils,sndkey32,ComObj,Larco_Functions;

type
  TfrmEmpleados = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    gbButtons: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    txtId: TEdit;
    txtNombre: TEdit;
    txtDepa: TEdit;
    txtTurno: TEdit;
    txtPercep: TEdit;
    txtFecha: TEdit;
    txtPuesto: TEdit;
    txtCosto: TEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    gvEmpleados: TGridView;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    btnAceptar: TButton;
    btnCancelar: TButton;
    btnExport: TButton;
    SaveDialog1: TSaveDialog;
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    Procedure BindEmpleado();
    Procedure BindGrid();
    Procedure ClearData();
    Procedure EnableControls(Value:Boolean);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnAceptarClick(Sender: TObject);
    function ValidateData():Boolean;
    procedure TabSheet2Show(Sender: TObject);
    procedure BuscarClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure SendTab(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEmpleados: TfrmEmpleados;
  giOpcion : Integer;
  Conn : TADOConnection;
  Qry : TADOQuery;
  sPermits : String;  
implementation

uses Main;

{$R *.dfm}

procedure TfrmEmpleados.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action :=caFree;
end;

procedure TfrmEmpleados.FormCreate(Sender: TObject);
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT * FROM tblEmpleados ORDER BY Id';
    Qry.Open;

    if Qry.RecordCount > 0 then
        BindEmpleado();

    self.Width := 527;
    self.Height := 466;
    sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.BindEmpleado();
begin
        txtId.Text := VarToStr(Qry['Id']);
        txtNombre.Text := VarToStr(Qry['Nombre']);
        txtDepa.Text := VarToStr(Qry['Departamento']);
        txtPuesto.Text := VarToStr(Qry['Puesto']);
        txtTurno.Text := VarToStr(Qry['Turno']);
        txtPercep.Text := VarToStr(Qry['Percepciones']);
        txtCosto.Text := VarToStr(Qry['CostoHora']);
        txtFecha.Text := VarToStr(Qry['FechaNac']);
end;

procedure TfrmEmpleados.ClearData();
begin
        txtId.Text := '';
        txtNombre.Text := '';
        txtDepa.Text := '';
        txtPuesto.Text := '';
        txtTurno.Text := '';
        txtPercep.Text := '';
        txtCosto.Text := '';
        txtFecha.Text := '';
end;

procedure TfrmEmpleados.EnableControls(Value:Boolean);
begin
        txtNombre.ReadOnly := Value;
        txtDepa.ReadOnly := Value;
        txtPuesto.ReadOnly := Value;
        txtTurno.ReadOnly := Value;
        txtPercep.ReadOnly := Value;
        txtCosto.ReadOnly := Value;
        txtFecha.ReadOnly := Value;
end;

procedure TfrmEmpleados.Button1Click(Sender: TObject);
begin
Qry.First;
BindEmpleado;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.Button2Click(Sender: TObject);
begin
Qry.Prior;
BindEmpleado;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.Button3Click(Sender: TObject);
begin
Qry.Next;
BindEmpleado;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.Button4Click(Sender: TObject);
begin
Qry.Last;
BindEmpleado;
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.NuevoClick(Sender: TObject);
begin
ClearData();
EnableControls(False);
txtId.SetFocus;
giOpcion := 1;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Editar.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmEmpleados.btnCancelarClick(Sender: TObject);
begin
EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;

Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;

BindEmpleado;
EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmEmpleados.EditarClick(Sender: TObject);
begin
giOpcion := 2;
EnableControls(False);
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Borrar.Enabled := False;
Buscar.Enabled := False;
txtNombre.SetFocus;
end;

procedure TfrmEmpleados.BorrarClick(Sender: TObject);
begin
giOpcion := 3;
btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Buscar.Enabled := False;
end;

procedure TfrmEmpleados.btnAceptarClick(Sender: TObject);
var SQLStr,SQLWhere : String;
Qry2 : TADOQuery;
begin

  if giOpcion = 1 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Insert;
        Qry['Nombre'] := txtNombre.Text;
        Qry['Departamento'] := txtDepa.Text;
        Qry['Puesto'] := txtPuesto.Text;
        Qry['Turno'] := txtTurno.Text;
        Qry['Percepciones'] := txtPercep.Text;
        Qry['CostoHora'] := txtCosto.Text;
        Qry['FechaNac'] := txtFecha.Text;
        Qry.Post;
  end
  else if giOpcion = 2 then
  begin
        if not ValidateData() then
          Exit;

        Qry.Edit;
        Qry['Nombre'] := txtNombre.Text;
        Qry['Departamento'] := txtDepa.Text;
        Qry['Puesto'] := txtPuesto.Text;
        Qry['Turno'] := txtTurno.Text;
        Qry['Percepciones'] := txtPercep.Text;
        Qry['CostoHora'] := txtCosto.Text;
        Qry['FechaNac'] := txtFecha.Text;
        Qry.Post;
  end
  else if giOpcion = 3 then
  begin
        if MessageDlg('Estas seguro que quieres borrar al empleado ' +
                      txtNombre.Text + '?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              Qry.Delete;
  end
  else if giOpcion = 4 then
  begin

        if txtId.Text <> '' Then
          begin
                if not Qry.Locate('Id',txtId.Text,[loPartialKey] ) then
                   begin
                           MessageDlg('No se encontro ningun Empleado con este Numero ' + txtId.Text + '.', mtInformation,[mbOk], 0);
                           txtId.SetFocus;
                           Exit;
                   end;
          end
        else
          begin
              SQLStr := 'SELECT Id FROM tblEmpleados ';
              if txtNombre.Text <> '' Then
                 SQLWhere := SQLWhere + 'Nombre LIKE ' + QuotedStr('%' + txtNombre.Text + '%') + ' AND ';

              if txtDepa.Text <> '' Then
                 SQLWhere := SQLWhere + 'Departamento LIKE ' + QuotedStr('%' + txtDepa.Text + '%') + ' AND ';

              if txtPuesto.Text <> '' Then
                 SQLWhere := SQLWhere + 'Puesto LIKE ' + QuotedStr('%' + txtPuesto.Text + '%') + ' AND ';

              if txtTurno.Text <> '' Then
                 SQLWhere := SQLWhere + 'Turno LIKE ' + QuotedStr('%' + txtTurno.Text + '%') + ' AND ';

              if SQLWhere = '' Then
                begin
                        MessageDlg('Por favor escribe algo en No.Empleado,Nombre,Departamento,Puesto o Turno.', mtInformation,[mbOk], 0);
                        txtNombre.SetFocus;
                        Exit;
                end
              else
                begin
                        SQLWhere := LeftStr(SQLWhere,Length(SQLWhere) - 5);
                        Qry2 := TADOQuery.Create(nil);
                        Qry2.Connection :=Conn;

                        Qry2.SQL.Clear;
                        Qry2.SQL.Text := SQLStr + 'WHERE ' + SQLWhere;
                        Qry2.Open;

                        if Qry2.RecordCount > 0 then
                           begin
                              if not Qry.Locate('Id',Qry2['Id'],[loPartialKey] ) then
                                 MessageDlg('No se encontro ningun Empleado con estos datos.', mtInformation,[mbOk], 0);
                           end
                        else
                           begin
                                 MessageDlg('No se encontro ningun Empleado con estos datos.', mtInformation,[mbOk], 0);
                                 txtNombre.SetFocus;
                                 Exit;
                           end;
                end;
          end;
        txtId.ReadOnly := True;
  end;

EnableControls(True);
btnAceptar.Enabled := False;
btnCancelar.Enabled := False;
Nuevo.Enabled := True;
Editar.Enabled := True;
Borrar.Enabled := True;
Buscar.Enabled := True;
BindEmpleado;
giOpcion := 0;
EnableFormButtons(gbButtons, sPermits);
end;

function TfrmEmpleados.ValidateData():Boolean;
begin
        ValidateData := True;
        if txtNombre.Text = '' Then
          begin
            MessageDlg('El nombre del empleado no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if txtDepa.Text = '' Then
          begin
            MessageDlg('El Departamento del empleado no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if txtPuesto.Text = '' Then
          begin
            MessageDlg('El Departamento del empleado no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if txtTurno.Text = '' Then
          begin
            MessageDlg('El Turno del empleado no puede estar vacio.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtTurno.Text) Then
          begin
            MessageDlg('El Turno debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtPercep.Text) Then
          begin
            MessageDlg('Las percepciones deben de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsNumeric(txtPercep.Text) Then
          begin
            MessageDlg('El Costo por Hora debe de ser un valor numerico.', mtInformation,[mbOk], 0);
            result :=  False;
          end;

        if not IsDate(txtFecha.Text) Then
          begin
            MessageDlg('Por favor escriba una fecha valida.', mtInformation,[mbOk], 0);
            result :=  False;
          end;
end;

procedure TfrmEmpleados.TabSheet2Show(Sender: TObject);
begin
BindGrid();
end;

procedure TfrmEmpleados.BindGrid();
var iCurrent : Integer;
begin
  iCurrent := Qry.RecNo;

  gvEmpleados.ClearRows;
  Qry.First;
  while not Qry.Eof do
  begin
      gvEmpleados.AddRow(1);
      gvEmpleados.Cells[0,gvEmpleados.RowCount -1] := VarToStr(Qry['Id']);
      gvEmpleados.Cells[1,gvEmpleados.RowCount -1] := VarToStr(Qry['Nombre']);
      gvEmpleados.Cells[2,gvEmpleados.RowCount -1] := VarToStr(Qry['Departamento']);
      gvEmpleados.Cells[3,gvEmpleados.RowCount -1] := VarToStr(Qry['Puesto']);
      gvEmpleados.Cells[4,gvEmpleados.RowCount -1] := VarToStr(Qry['Turno']);
      gvEmpleados.Cells[5,gvEmpleados.RowCount -1] := VarToStr(Qry['Percepciones']);
      gvEmpleados.Cells[6,gvEmpleados.RowCount -1] := VarToStr(Qry['CostoHora']);
      gvEmpleados.Cells[7,gvEmpleados.RowCount -1] := VarToStr(Qry['FechaNac']);
      Qry.Next;
  end;

  Qry.RecNo := iCurrent;
end;

procedure TfrmEmpleados.BuscarClick(Sender: TObject);
begin
ClearData();
txtId.ReadOnly := False;
EnableControls(False);
txtId.SetFocus;
giOpcion := 4;

btnAceptar.Enabled := True;
btnCancelar.Enabled := True;

Nuevo.Enabled := False;
Editar.Enabled := False;
Borrar.Enabled := False;
end;

procedure TfrmEmpleados.btnExportClick(Sender: TObject);
var sFileName: String;
begin

  SaveDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if SaveDialog1.Execute then
  begin
    sFileName := SaveDialog1.FileName;
    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ExportGrid(gvEmpleados,sFileName);

  end;
end;

procedure TfrmEmpleados.FormResize(Sender: TObject);
begin
        PageControl1.Height := self.Height - 49;
        PageControl1.Width  := self.Width - 22;

        gbButtons.Height := PageControl1.Height - 40;
        gbButtons.Width := PageControl1.Width - 30;

        gvEmpleados.Height := gbButtons.Height - 32;
        gvEmpleados.Width := gbButtons.Width - 4;

        btnExport.Left := gvEmpleados.Left + gvEmpleados.Width - btnExport.Width;
        btnExport.Top := gvEmpleados.Top + gvEmpleados.Height + 7;
        //showmessage(IntToStr(gvEmpleados.Top + gvEmpleados.Height));
end;

procedure TfrmEmpleados.SendTab(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end;
end;

procedure TfrmEmpleados.ExportGrid(Grid: TGridView;sFileName: String);
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
      Sheet.Name := 'Empleados';

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
end;

end.
