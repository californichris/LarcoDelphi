unit Larco_Functions;

interface

uses
  Windows, SysUtils, Dialogs,  Db, ADODB, Classes,Variants,Winsock,
  NMsmtp,DateUtils,StdCtrls,CustomGridViewControl, CustomGridView, GridView,
  IdTrivialFTPBase,ComObj,Controls,LTCUtils;

  procedure BindComboEmpleados(sConnString: String; cmbEmpleados: TComboBox);
  procedure BindComboTareasDetectado(sConnString: String; cmbTareas: TComboBox; cmbDetectado: TComboBox);
  procedure BindGridClientes(sConnString: String; gvClientes: TGridView);
  function getFormYear(sConnString: String; sFormName: String): String;
  function getUserPermits(sConnString: String; sFormName: String; sUserLogin: String): String;
  procedure EnableFormButtons(gbButtons: TGroupBox; sPermits: String);
  function getStringBoolean(value: String):Boolean;
  function ValidateEmpleado(sConnString: String; Id: String):Boolean;
implementation

function ValidateEmpleado(sConnString: String; Id: String):Boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := False;
    if UT(Id) = '' then
        Exit;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection := Conn;

      SQLStr := 'SELECT Nombre FROM tblEmpleados WHERE Activo = 1 AND Id =  ' + IntToStr(StrToInt(Id));

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      if Qry.RecordCount > 0 then
          Result := True;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure BindComboEmpleados(sConnString: String; cmbEmpleados: TComboBox);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT ID,Nombre FROM tblEmpleados Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      cmbEmpleados.Items.Clear;
      cmbEmpleados.Items.Add('Todos');
      cmbEmpleados.Items.Add('000 - Desconocido');
      while not Qry.Eof do begin
          cmbEmpleados.Items.Add(FormatFloat('000',Qry['ID']) + ' - ' + Qry['Nombre']);
          Qry.Next;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    cmbEmpleados.Text := '';
end;

procedure BindComboTareasDetectado(sConnString: String;
                                   cmbTareas: TComboBox; cmbDetectado: TComboBox);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      cmbTareas.Items.Clear;
      cmbDetectado.Items.Clear;
      cmbTareas.Items.Add('Todos');
      cmbDetectado.Items.Add('Todos');
      while not Qry.Eof do begin
          cmbTareas.Items.Add(Qry['Nombre']);
          cmbDetectado.Items.Add(Qry['Nombre']);
          Qry.Next;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    cmbTareas.Text := '';
    cmbDetectado.Text := '';
end;

procedure BindGridClientes(sConnString: String; gvClientes: TGridView);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '060,062,699,799,899,999,960';

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := sConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Distinct Clave FROM tblClientes Order By Clave';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvClientes.ClearRows;
      while not Qry.Eof do begin
          gvClientes.AddRow(1);
          gvClientes.Cells[0,gvClientes.RowCount -1] := VarToStr(Qry['Clave']);
          if (slClientes.IndexOf(VarToStr(Qry['Clave'])) = -1) then begin
                  gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
          end;

          Qry.Next;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

function getFormYear(sConnString: String; sFormName: String): String;
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    Result := '';
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblScreens WHERE SCR_FormName = ' + QuotedStr(sFormName);

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        if Qry.RecordCount > 0 then begin
                Result := VarToStr(Qry['SCR_Year']);
        end
        else begin
                Result := WordToStr( YearOf(Date) );
        end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

function getUserPermits(sConnString: String; sFormName: String; sUserLogin: String): String;
var SQLStr : String;
Conn : TADOConnection;
Qry : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    Result := '';
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT U.USE_Login,S.SCR_FormName,GS.Nuevo,GS.Editar,GS.Borrar,GS.Buscar ' +
                  'FROM tblUsers U ' +
                  'INNER JOIN tblUser_Groups UG ON U.USE_ID = UG.USE_ID ' +
                  'INNER JOIN tblGroup_Screens GS on UG.Group_ID = GS.Group_ID ' +
                  'INNER JOIN tblScreens S ON GS.SCR_ID = S.SCR_ID ' +
                  'WHERE USE_Login = ' + QuotedStr(sUserLogin) +
                  ' AND SCR_FormName = ' + QuotedStr(sFormName);

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        if Qry.RecordCount > 0 then begin
                Result := VarToStr(Qry['Nuevo']) + ',' + VarToStr(Qry['Editar']) + ',' +
                          VarToStr(Qry['Borrar']) + ',' + VarToStr(Qry['Buscar']);
        end
        else begin
                Result := '0,0,0,0';
        end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure EnableFormButtons(gbButtons: TGroupBox; sPermits: String);
var slPermits : TStringList;
Nuevo, Editar, Borrar, Buscar : TControl;
begin
  if (sPermits = '') then sPermits := '0,0,0,0';

  slPermits := TStringList.Create;
  slPermits.CommaText := sPermits;

  Nuevo := gbButtons.FindChildControl('Nuevo');
  Editar := gbButtons.FindChildControl('Editar');
  Borrar := gbButtons.FindChildControl('Borrar');
  Buscar := gbButtons.FindChildControl('Buscar');
  if (nil <> Nuevo) then
  begin
        Nuevo.Enabled := getStringBoolean(slPermits[0]);
  end;

  if (nil <> Editar) then
  begin
        Editar.Enabled := getStringBoolean(slPermits[1]);
  end;

  if (nil <> Borrar) then
  begin
        Borrar.Enabled := getStringBoolean(slPermits[2]);
  end;

  if (nil <> Buscar) then
  begin
        Buscar.Enabled := getStringBoolean(slPermits[3]);
  end;
end;

function getStringBoolean(value: String):Boolean;
begin
    Result := False;
    if (value <> '0') then begin
        Result := True;
    end;

end;

end.
