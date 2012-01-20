unit Productos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView,LTCUtils, CustomGridViewControl, CustomGridView,
  GridView,Columns,ColumnClasses,ADODB,DB,IniFiles,All_Functions, Menus,Larco_Functions;

type
  TfrmProductos = class(TForm)
    gbButtons: TGroupBox;
    gvProductos: TGridView;
    Label1: TLabel;
    txtNombre: TEdit;
    Button1: TButton;
    PopupMenu1: TPopupMenu;
    Borrar1: TMenuItem;
    Editar1: TMenuItem;
    btnCancelar: TButton;
    GroupBox2: TGroupBox;
    gvTareas: TGridView;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    procedure BindProductos();
    procedure BindTareas();
    procedure BindSelectedTareas(Product: String);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    function ProductoExists(Producto: String):Boolean;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure gvProductosSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure BorrarProducto();
    procedure SaveRoutes();
    procedure SaveAggregateValues();
    procedure NuevoClick(Sender: TObject);
    Procedure EnableControls(Value:Boolean);
    procedure EditarClick(Sender: TObject);
    procedure BorrarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure txtHorasSetUpKeyPress(Sender: TObject; var Key: Char);
    procedure txtPrecioHoraKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProductos: TfrmProductos;
  giOpcion : Integer;
  giRow : Integer;
  sPermits : String;  
implementation

uses Main;

{$R *.dfm}

procedure TfrmProductos.BindProductos();
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

    SQLStr := 'SELECT * FROM tblProductos Order By Nombre';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    gvProductos.ClearRows;
    While not Qry.Eof do
    Begin
        gvProductos.AddRow(1);
        gvProductos.Cells[0,gvProductos.RowCount -1] := VarToStr(Qry['Id']);
        gvProductos.Cells[1,gvProductos.RowCount -1] := VarToStr(Qry['Nombre']);
        //gvProductos.Cells[2,gvProductos.RowCount -1] := VarToStr(Qry['PrecioHora']);
        //gvProductos.Cells[3,gvProductos.RowCount -1] := VarToStr(Qry['HorasSetUp']);
        //gvProductos.Cells[4,gvProductos.RowCount -1] := VarToStr(Qry['Es062']);
        Qry.Next;
    End;


    Qry.Close;
    Conn.Close;

  sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
  EnableFormButtons(gbButtons, sPermits);

end;

procedure TfrmProductos.FormCreate(Sender: TObject);
begin
    giOpcion := 0;
    BindProductos();
    BindTareas();
    txtNombre.Text := gvProductos.Cells[1,gvProductos.SelectedRow];
    //txtPrecioHora.Text := gvProductos.Cells[2,gvProductos.SelectedRow];
    //txtHorasSetUp.Text := gvProductos.Cells[3,gvProductos.SelectedRow];
    //chkEs062.Checked := getStringBoolean(gvProductos.Cells[4,gvProductos.SelectedRow]);
    BindSelectedTareas(gvProductos.Cells[1,gvProductos.SelectedRow]);

    sPermits := getUserPermits(frmMain.sConnString, Self.Name, frmMain.sUserLogin);
    EnableFormButtons(gbButtons, sPermits);
end;

procedure TfrmProductos.Button1Click(Sender: TObject);
var Conn : TADOConnection;
SQLStr : String;
sProducto,sPrecioHora, sHorasSetUp, sId : String; //sEs062, oldEs062
begin

    If giOpcion <> 3 Then
    begin
        If txtNombre.Text = '' then
          begin
            MessageDlg('Por favor escriba un nombre de Producto.', mtInformation,[mbOk], 0);
            Exit;
          end;

        if ProductoExists(txtNombre.Text) then
          begin
            MessageDlg('Ya existe un producto con este nombre.', mtInformation,[mbOk], 0);
            Exit;
          end;
    end;

    sProducto := gvProductos.Cells[1,gvProductos.SelectedRow];
    sPrecioHora := gvProductos.Cells[2,gvProductos.SelectedRow];
    sHorasSetUp := gvProductos.Cells[3,gvProductos.SelectedRow];
    //oldEs062 := gvProductos.Cells[4,gvProductos.SelectedRow];

    sId := gvProductos.Cells[0,gvProductos.SelectedRow];
    //sEs062 := '0';
    //if chkEs062.checked then sEs062 := '1';

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;

    if giOpcion = 2 Then
    begin
        If (UT(txtNombre.Text) <>  UT(sProducto)) then
          begin
            SQLStr := 'UPDATE tblItems SET PRO_Nombre = ' +  QuotedStr(txtNombre.Text) +
                      ' WHERE PRO_Nombre = ' +  QuotedStr(sProducto);

            Conn.Execute(SQLStr);

            SQLStr := 'UPDATE tblProductos SET Nombre = ' +  QuotedStr(txtNombre.Text) +
                      ' WHERE Id = ' +  sId;
                      {', Es062 = ' +  sEs062 +
                      ', PrecioHora = ' +  QuotedStr(txtPrecioHora.Text) +
                      ', HorasSetUp = ' +  QuotedStr(txtHorasSetUp.Text) +
                      }


            Conn.Execute(SQLStr);
          end;

        SaveRoutes();
        SaveAggregateValues();
    end
    else if giOpcion = 1 Then
    begin
        SQLStr := 'INSERT INTO tblProductos(Nombre) ' +
                  'VALUES(' + QuotedStr(txtNombre.Text) + ')';

        Conn.Execute(SQLStr);

        SaveRoutes();
    end
    else if giOpcion = 3 Then
    begin
        BorrarProducto();
    end;


    Conn.Close;

    BindProductos();
    EnableControls(True);
    gvProductos.SelectCell(1,gvProductos.SelectedRow);
    gvProductos.SetFocus;

    giOpcion := 0;
end;

function TfrmProductos.ProductoExists(Producto: String):Boolean;
var i:integer;
begin
        ProductoExists := False;
        for i:=0 to gvProductos.RowCount -1 do
          begin
                if (giOpcion = 1) and (i <> giRow ) then
                  if UT(Producto) = UT(gvProductos.Cells[1,i]) then
                    ProductoExists := True;
          end;
end;

procedure TfrmProductos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmProductos.BindTareas();
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
        gvTareas.Cells[3,gvTareas.RowCount -1] := VarToStr(Qry['Nombre']);
        Qry.Next;
    End;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmProductos.gvProductosSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
  txtNombre.Text := gvProductos.Cells[1,ARow];
  //txtPrecioHora.Text := gvProductos.Cells[2,ARow];
  //txtHorasSetUp.Text := gvProductos.Cells[3,ARow];
  //chkEs062.Checked := getStringBoolean(gvProductos.Cells[4,ARow]);
  BindSelectedTareas(gvProductos.Cells[1,ARow]);
  giRow := gvProductos.SelectedRow;
end;


procedure TfrmProductos.BindSelectedTareas(Product: String);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr, Produc_ID : String;
i: Integer;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT Rou_From,Rou_To FROM tblRouting WHERE Nombre IN (''*'',' +
              QuotedStr(Product) + ') ORDER BY Rou_From';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    //Clear selected tasks
    for i:= 0 to gvTareas.RowCount - 1 do
      begin
            gvTareas.Cell[2,i].AsBoolean  := False;
            gvTareas.Cell[4,i].AsBoolean  := False;
      end;

    While not Qry.Eof do
    Begin
        for i:= 0 to gvTareas.RowCount - 1 do
          begin
                if (gvTareas.Cells[0,i] = Qry['Rou_From']) or (gvTareas.Cells[0,i] = Qry['Rou_To']) then
                   begin
                        gvTareas.Cell[2,i].AsBoolean  := True;
                        gvTareas.Cell[4,i].AsBoolean  := True;
                   end;
          end;
        Qry.Next;
    End;


    Produc_ID := gvProductos.Cells[0,gvProductos.SelectedRow];

    SQLStr := 'SELECT * FROM tblAggregateValue WHERE Product_ID = ' + Produc_ID + ' ORDER BY Task_ID';

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    While not Qry.Eof do
    Begin
        for i:= 0 to gvTareas.RowCount - 1 do
          begin
                if (gvTareas.Cells[0,i] = Qry['Task_ID']) then
                   begin
                        gvTareas.Cells[5,i] := VarToStr(Qry['Value']);
                        gvTareas.Cells[6,i] := VarToStr(Qry['Value']);
                        break;
                   end;
          end;
        Qry.Next;
    End;


    Qry.Close;
    Conn.Close;
end;

procedure TfrmProductos.BorrarProducto();
var sId,sProducto : string;
Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
  sProducto := gvProductos.Cells[1,gvProductos.SelectedRow];
  sId := gvProductos.Cells[0,gvProductos.SelectedRow];

  if MessageDlg('Estas seguro que quieres borrar el Producto ' + sProducto + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;
  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  //validar que no haya items de este producto dentro del flujo de trabajo.
  SQLStr := 'SELECT TOP 1 ITE_Nombre FROM tblItems WHERE PRO_Nombre = ' + QuotedStr(sProducto);

  Qry.SQL.Clear;
  Qry.SQL.Text := SQLStr;
  Qry.Open;

  if Qry.RecordCount > 0 Then
    begin
          ShowMessage('No puedes borrar el producto porque existen ' + chr(13) +
                      'Ordenes de trabajo de este Producto dentro del flujo de trabajo.');
    end
  else
    begin
      //Borrar Producto
      SQLStr := 'DELETE FROM tblProductos WHERE Id = ' + sId;
      Conn.Execute(SQLStr);

      //Borrar Rutas de este producto
      SQLStr := 'DELETE FROM tblRouting WHERE Nombre = ' + QuotedStr(sProducto);
      Conn.Execute(SQLStr);
    end;

  Qry.Close;
  Conn.Close;
  BindProductos();
  gvProductos.SelectCell(1,0);
  gvProductos.SetFocus;
end;

procedure TfrmProductos.NuevoClick(Sender: TObject);
var i:Integer;
begin
  txtNombre.Text := '';
  //txtPrecioHora.Text := '';
  //txtHorasSetUp.Text := '';
  //chkEs062.Checked := false;

  //Clear selected tasks
  for i:= 0 to gvTareas.RowCount - 1 do
    begin
          gvTareas.Cell[2,i].AsBoolean  := False;
          gvTareas.Cell[4,i].AsBoolean  := False;
          gvTareas.Cells[5,i] := '1.00';
          gvTareas.Cells[6,i] := '1.00';
    end;

  Editar.Enabled := False;
  Borrar.Enabled := False;
  EnableControls(False);
  giOpcion := 1;
  txtNombre.SetFocus;
end;

procedure TfrmProductos.EnableControls(Value:Boolean);
begin
        txtNombre.ReadOnly := Value;
        //txtPrecioHora.ReadOnly := Value;
        //txtHorasSetUp.ReadOnly := Value;
        //chkEs062.Enabled := not Value;

        Button1.Enabled := not Value;
        btnCancelar.Enabled := not Value;

        if Value then
        begin
          Nuevo.Enabled := Value;
          Borrar.Enabled := Value;
          Editar.Enabled := Value;
        end;

        if (not Value) Then begin
                gvTareas.Columns[2].Options := gvTareas.Columns[2].Options + [coEditing];
                gvTareas.Columns[5].Options := gvTareas.Columns[5].Options + [coEditing];
        end
        else begin
                gvTareas.Columns[2].Options := gvTareas.Columns[2].Options - [coEditing];
                gvTareas.Columns[5].Options := gvTareas.Columns[5].Options - [coEditing];                
        end;

  EnableFormButtons(gbButtons, sPermits);

end;


procedure TfrmProductos.EditarClick(Sender: TObject);
begin
  Nuevo.Enabled := False;
  Borrar.Enabled := False;
  giOpcion := 2;
  EnableControls(False);
  txtNombre.SetFocus;
end;

procedure TfrmProductos.BorrarClick(Sender: TObject);
begin
  Nuevo.Enabled := False;
  Editar.Enabled := False;
  giOpcion := 3;
  EnableControls(False);
end;

procedure TfrmProductos.btnCancelarClick(Sender: TObject);
begin
    EnableControls(True);
    giOpcion := 0;
    gvProductos.SelectCell(1,gvProductos.SelectedRow);
    gvProductos.SetFocus;
end;

procedure TfrmProductos.SaveRoutes();
var i : Integer;
bChanged : Boolean;
Conn : TADOConnection;
SQLStr,sOldProduct,sProduct : String;
RouFrom,RouTo : String;
begin

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  //Old Product Name
  sOldProduct := gvProductos.Cells[1,gvProductos.SelectedRow];
  sProduct := txtNombre.Text;

  if giOpcion = 2 Then
  begin

      //Check if there were changes on the routes
      bChanged := False;
      for i:= 0 to gvTareas.RowCount - 1 do
            if  gvTareas.Cell[2,i].AsBoolean <> gvTareas.Cell[4,i].AsBoolean then
               begin
                    bChanged := True;
                    Break;
               end;

      // If there a product name change update the name on Routes table
      if UT(sOldProduct) <> UT(sProduct) then
        begin
              SQLStr := 'UPDATE tblRouting SET Nombre = ' + QuotedStr(sProduct) +
              ' WHERE Nombre = ' + QuotedStr(sOldProduct);
              Conn.Execute(SQLStr);
        end;


      if bChanged = False then
        begin

            Conn.Close;
            Exit;
        end;

       //delete previos routes
       SQLStr := 'DELETE FROM tblRouting WHERE Nombre = ' + QuotedStr(sOldProduct);
       Conn.Execute(SQLStr);
  end;

  RouFrom := '';
  RouTo := '';

  for i:= 0 to gvTareas.RowCount - 1 do
    begin
          if gvTareas.Cell[2,i].AsBoolean = True  Then
            begin
                if RouFrom = '' then
                  begin
                      RouFrom := gvTareas.Cells[0,i];
                  end
                else
                  begin
                      RouTo := gvTareas.Cells[0,i];
                      // insert route
                      SQLStr := 'INSERT INTO tblRouting(Nombre,Rou_From,Rou_Code,Rou_To) ' +
                                'VALUES(' + QuotedStr(sProduct) + ',' + RouFrom + ',' + QuotedStr('OK') +
                                ',' + RouTo + ')';

                      Conn.Execute(SQLStr);
                      RouFrom := gvTareas.Cells[0,i];
                  end;
            end;
    end;



  Conn.Close;
end;

procedure TfrmProductos.SaveAggregateValues();
var i : Integer;
Conn : TADOConnection;
SQLStr,Produc_ID : String;
begin

  //Create Connection
  Conn := TADOConnection.Create(nil);
  Conn.ConnectionString := frmMain.sConnString;
  Conn.LoginPrompt := False;

  Produc_ID := gvProductos.Cells[0,gvProductos.SelectedRow];
  for i:= 0 to gvTareas.RowCount - 1 do
    begin
           if gvTareas.Cell[2,i].AsBoolean = True Then
             begin
                SQLStr := 'UPDATE tblAggregateValue SET [Value] = ' + gvTareas.Cells[5,i] +
                          ' WHERE Product_ID = ' + Produc_ID + ' AND Task_ID = ' + gvTareas.Cells[0,i];

                Conn.Execute(SQLStr);
             end
           else
             begin
                SQLStr := 'UPDATE tblAggregateValue SET [Value] = 0.0' +
                          ' WHERE Product_ID = ' + Produc_ID + ' AND Task_ID = ' + gvTareas.Cells[0,i];

                Conn.Execute(SQLStr);
             end;
          {if gvTareas.Cells[5,i] <> gvTareas.Cells[6,i]  Then
            begin
                SQLStr := 'UPDATE tblAggregateValue SET [Value] = ' + gvTareas.Cells[5,i] +
                          ' WHERE Product_ID = ' + Produc_ID + ' AND Task_ID = ' + gvTareas.Cells[0,i];

                Conn.Execute(SQLStr);
            end;  }
    end;

  Conn.Close;
end;

procedure TfrmProductos.txtHorasSetUpKeyPress(Sender: TObject; var Key: Char);
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

procedure TfrmProductos.txtPrecioHoraKeyPress(Sender: TObject;  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar((Sender as TEdit).Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;
end;

end.
