unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, IniFiles,ComCtrls,Productos,Monitor,Grupos,Tareas,Empleados,Routing,
  Rutas,Ventas,Scrap,Clientes,Year,FechaEntrega,Editor, StdCtrls, ExtCtrls,ADODB,DB,
  ReporteScrap,RelacionOrdenCompra,LTCUtils,ComObj;

type
  TfrmMain = class(TForm)
    StatusBar: TStatusBar;
    MainMenu1: TMainMenu;
    Monitor1: TMenuItem;
    Catalogos1: TMenuItem;
    Productos1: TMenuItem;
    areas1: TMenuItem;
    areas2: TMenuItem;
    Empleados1: TMenuItem;
    Rutas1: TMenuItem;
    Ventas1: TMenuItem;
    Clientes1: TMenuItem;
    Reportes1: TMenuItem;
    Empleados2: TMenuItem;
    Sistema1: TMenuItem;
    EstablecerA1: TMenuItem;
    EditarOrden1: TMenuItem;
    Usuarios1: TMenuItem;
    Salir1: TMenuItem;
    ProgressBar1: TProgressBar;
    lblScrap: TLabel;
    Timer1: TTimer;
    ReportedeScrap1: TMenuItem;
    RelaciondeOrdenesdeCompra1: TMenuItem;
    Screens1: TMenuItem;
    ReportedeRetrabajo1: TMenuItem;
    ReportedeScrapenDinero1: TMenuItem;
    ipodeCambio1: TMenuItem;
    ReportedeCumplimiento1: TMenuItem;
    ReportedePromediodeCumplimiento1: TMenuItem;
    ReportedeRetrabajoenDinero1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    Categories1: TMenuItem;
    GruposdeUsuarios1: TMenuItem;
    Menu1: TMenuItem;
    Permisos1: TMenuItem;
    EditordeScrap1: TMenuItem;
    N6: TMenuItem;
    ReportedeCargadeTrabajo1: TMenuItem;
    ReportedeProductividad1: TMenuItem;
    Facturacion1: TMenuItem;
    Facturacion2: TMenuItem;
    PendientesdeFacturar1: TMenuItem;
    N7: TMenuItem;
    Contribuyente1: TMenuItem;
    Materiales1: TMenuItem;
    ProductosTerminados1: TMenuItem;
    Windows1: TMenuItem;
    Tile1: TMenuItem;
    Cascade1: TMenuItem;
    N8: TMenuItem;
    CloseAll1: TMenuItem;
    EditordeRetrabajo1: TMenuItem;
    UnidadesdeMedida1: TMenuItem;
    iposdeMaterial1: TMenuItem;
    Paises1: TMenuItem;
    Entradas1: TMenuItem;
    Proveedores1: TMenuItem;
    SalidasAlmacen1: TMenuItem;
    Configuracion1: TMenuItem;
    ReporteEntradasSalidasBorradas1: TMenuItem;
    ReportedePiezasTerminadas1: TMenuItem;
    ReporteEntradasSalidasAlmacen1: TMenuItem;
    ReporteEntradasSalidasLarco1: TMenuItem;
    ReportedeMaterialesEscasos1: TMenuItem;
    SalidasLarco1: TMenuItem;
    ReporteProductividadEmpleado1: TMenuItem;
    ReporteProductividadEmpleado2: TMenuItem;
    ReportedeMaterialesporOrdendeTrabajo1: TMenuItem;
    DiasInhabiles1: TMenuItem;
    ReportedeCumplimientodeTiempodeEntrega1: TMenuItem;
    Planos1: TMenuItem;
    Stock1: TMenuItem;
    EntradasStock1: TMenuItem;
    EntradasvsSalidas1: TMenuItem;
    EntradasvsSalidasPorPlano1: TMenuItem;
    TotalPiezasStock1: TMenuItem;
    PiezasenStock1: TMenuItem;
    procedure Productos1Click(Sender: TObject);
    procedure Monitor1Click(Sender: TObject);
    procedure areas1Click(Sender: TObject);
    procedure areas2Click(Sender: TObject);
    procedure Empleados1Click(Sender: TObject);
    procedure Rutas1Click(Sender: TObject);
    function FormIsRunning(FormName: String):Boolean;
    procedure Ventas1Click(Sender: TObject);
    procedure Clientes1Click(Sender: TObject);
    procedure EstablecerA1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Empleados2Click(Sender: TObject);
    procedure EditarOrden1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Usuarios1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure lblScrapClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ReportedeScrap1Click(Sender: TObject);
    procedure RelaciondeOrdenesdeCompra1Click(Sender: TObject);
    procedure Screens1Click(Sender: TObject);
    procedure ReportedeRetrabajo1Click(Sender: TObject);
    procedure ReportedeScrapenDinero1Click(Sender: TObject);
    procedure ipodeCambio1Click(Sender: TObject);
    procedure ReportedeCumplimiento1Click(Sender: TObject);
    procedure ReportedePromediodeCumplimiento1Click(Sender: TObject);
    procedure ReportedeRetrabajoenDinero1Click(Sender: TObject);
    procedure Categories1Click(Sender: TObject);
    procedure GruposdeUsuarios1Click(Sender: TObject);
    procedure Menu1Click(Sender: TObject);
    procedure Permisos1Click(Sender: TObject);
    procedure BindMenu(userId : String);
    procedure assignEvent(formName : String; menuItem : TMenuItem);
    procedure EditordeScrap1Click(Sender: TObject);
    procedure ReportedeCargadeTrabajo1Click(Sender: TObject);
    procedure Facturacion2Click(Sender: TObject);
    procedure PendientesdeFacturar1Click(Sender: TObject);
    procedure ReportedeProductividad1Click(Sender: TObject);
    procedure Contribuyente1Click(Sender: TObject);
    procedure Materiales1Click(Sender: TObject);
    procedure ProductosTerminados1Click(Sender: TObject);
    procedure Tile1Click(Sender: TObject);
    procedure Cascade1Click(Sender: TObject);
    procedure CloseAll1Click(Sender: TObject);
    procedure EditordeRetrabajo1Click(Sender: TObject);
    procedure UnidadesdeMedida1Click(Sender: TObject);
    procedure TiposdeMaterial1Click(Sender: TObject);
    procedure Paises1Click(Sender: TObject);
    procedure Entradas1Click(Sender: TObject);
    procedure Proveedores1Click(Sender: TObject);
    procedure SalidasAlmacen1Click(Sender: TObject);
    procedure Configuracion1Click(Sender: TObject);
    procedure ReporteEntradasSalidasBorradas1Click(Sender: TObject);
    procedure ReportedePiezasTerminadas1Click(Sender: TObject);
    procedure ReporteEntradasSalidasAlmacen1Click(Sender: TObject);
    procedure ReporteEntradasSalidasLarco1Click(Sender: TObject);
    procedure ReportedeMaterialesEscasos1Click(Sender: TObject);
    procedure SalidasLarco1Click(Sender: TObject);
    procedure ReporteProductividadEmpleado1Click(Sender: TObject);
    procedure ReporteProductividadEmpleado2Click(Sender: TObject);
    procedure ReportedeMaterialesporOrdendeTrabajo1Click(Sender: TObject);
    procedure DiasInhabiles1Click(Sender: TObject);
    procedure ReportedeCumplimientodeTiempodeEntrega1Click(
      Sender: TObject);
    procedure Planos1Click(Sender: TObject);
    procedure EntradasStock1Click(Sender: TObject);
    procedure EntradasvsSalidas1Click(Sender: TObject);
    procedure EntradasvsSalidasPorPlano1Click(Sender: TObject);
    procedure TotalPiezasStock1Click(Sender: TObject);
    procedure PiezasenStock1Click(Sender: TObject);
  private
    { Private declarations }
    FirstTimeLogin : Boolean;
  public
    { Public declarations }
    sConnString : String;
    sUserID : String;
    sUserLogin : String;
    sUserPassword : String;
    iIntervalo: Integer;
  end;

var
  frmMain: TfrmMain;
  gsConnString,StartDDir: String;

implementation

uses Login, Users, PorcentajeScrap, Screens, PorcentajeRetrabajo,
  PorcentajeScrapDinero, ExchangeRate, ReporteCumplimiento,
  ReportePromedioCump, PorcentajeRetrabajoDinero, CatalogoScreens,
  CatalogoCategories, CatalogoGrupos, MenuEditor, CatalogoPermisos,
  EditorDeScrap, ReporteCargaTrabajo, Facturacion, PendientesFacturar,
  ReporteProductividad, CatalogoContribuyente, CatalogoMateriales,
  CatalogoPrpductosTerminados, EditorDeRetrabajo, CatalogoUnidadMedida,
  CatalogoTipoMaterial, CatalogoPaises, Entradas, CatalogoProvedores,
  SalidasAlmacen, InventariosConfiguracion, ReporteEntradasSalidasBorradas,
  ReportePiezasTerminadas, ReporteEntradasSalidasAlmacen,
  ReporteEntradasSalidasLarco, ReporteMaterialesEscasos, SalidasLarco,
  ReporteProductividadEmpleado, ReporteProductividadEmpleadoDinero,
  ReporteMaterialesPorOrden, CatalogoDiasInhabiles,
  ReporteCumplimientoFechaEntrega, CatalogoPlanos, EntradasSalidasStock,
  ReporteEntradasSalidasStock, ReporteEntradasSalidasPlano,
  ReporteTotalPiezasStock, ReportePiezasStock;

{$R *.dfm}

procedure TfrmMain.Productos1Click(Sender: TObject);
begin
if FormIsRunning('frmProductos') Then
  begin
        setActiveWindow(frmProductos.Handle);
        frmProductos.WindowState := wsNormal;
  end
else
  begin
    Application.CreateForm(TfrmProductos,frmProductos);
    frmProductos.Show;
  end;
end;

procedure TfrmMain.Monitor1Click(Sender: TObject);
begin
if FormIsRunning('Form1') Then
  begin
        setActiveWindow(Form1.Handle);
        Form1.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TForm1,Form1);
        Form1.Show;
  end;
end;

procedure TfrmMain.areas1Click(Sender: TObject);
begin
if FormIsRunning('frmGrupos') Then
  begin
        setActiveWindow(frmGrupos.Handle);
        frmGrupos.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmGrupos,frmGrupos);
        frmGrupos.Show;
  end;
end;

procedure TfrmMain.areas2Click(Sender: TObject);
begin
if FormIsRunning('frmTareas') Then
  begin
        setActiveWindow(frmTareas.Handle);
        frmTareas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmTareas,frmTareas);
        frmTareas.Show;
  end;
end;

procedure TfrmMain.Empleados1Click(Sender: TObject);
begin
if FormIsRunning('frmEmpleados') Then
  begin
        setActiveWindow(frmEmpleados.Handle);
        frmEmpleados.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEmpleados,frmEmpleados);
        frmEmpleados.Show;
  end;
end;

procedure TfrmMain.Rutas1Click(Sender: TObject);
begin
if FormIsRunning('frmRutas') Then
  begin
        setActiveWindow(frmRutas.Handle);
        frmRutas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmRutas,frmRutas);
        frmRutas.Show;
  end;
end;

function TfrmMain.FormIsRunning(FormName: String):Boolean;
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

procedure TfrmMain.Ventas1Click(Sender: TObject);
begin
if FormIsRunning('frmVentas') Then
  begin
        setActiveWindow(frmVentas.Handle);
        frmVentas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmVentas,frmVentas);
        frmVentas.Show;
  end;
end;

procedure TfrmMain.Clientes1Click(Sender: TObject);
begin
if FormIsRunning('frmClientes') Then
  begin
        setActiveWindow(frmClientes.Handle);
        frmClientes.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmClientes,frmClientes);
        frmClientes.Show;
  end;
end;

procedure TfrmMain.EstablecerA1Click(Sender: TObject);
begin
if FormIsRunning('frmYear') Then
  begin
        setActiveWindow(frmYear.Handle);
        frmYear.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmYear,frmYear);
        frmYear.Show;
  end;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var sUser,sPassword,sServer,sDB,sYear : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    //ShowMessage(StartDDir);
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','');
    sDB := IniFile.ReadString('Conn','DB','');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','');
    sYear := IniFile.ReadString('System','Year','');
    iIntervalo := StrToInt(IniFile.ReadString('Relacion','Refresh','30000'));

    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;

    sConnString :=  gsConnString;

    StatusBar.Panels[0].Text := sServer;
    StatusBar.Panels[1].Text := sDB;
    StatusBar.Panels[2].Text := sUser;
    StatusBar.Panels[3].Text := DateToStr(Now);

    ProgressBar1.Parent := StatusBar;  //adopt the Progressbar
    ProgressBar1.Top    := 3 ;      //set size of
    ProgressBar1.Left   := StatusBar.Panels.Items[0].Width +
                           StatusBar.Panels.Items[1].Width + StatusBar.Panels.Items[2].Width +
                           StatusBar.Panels.Items[3].Width + 2;

    lblScrap.Visible := False;
    lblScrap.Parent := StatusBar;
    lblScrap.Top    := 3 ;      //set size of
    lblScrap.Left   := StatusBar.Panels.Items[0].Width +
                       StatusBar.Panels.Items[1].Width + StatusBar.Panels.Items[2].Width +
                       StatusBar.Panels.Items[3].Width + StatusBar.Panels.Items[4].Width + 2;

    Self.Caption := Self.Caption + ' 3.5';
    Timer1Timer(nil);

end;

procedure TfrmMain.Empleados2Click(Sender: TObject);
begin
if FormIsRunning('frmEntrega') Then
  begin
        setActiveWindow(frmEntrega.Handle);
        frmEntrega.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEntrega,frmEntrega);
        frmEntrega.Show;
  end;
end;

procedure TfrmMain.EditarOrden1Click(Sender: TObject);
begin
if FormIsRunning('frmEditor') Then
  begin
        setActiveWindow(frmEditor.Handle);
        frmEditor.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEditor,frmEditor);
        frmEditor.Show;
  end;

end;

procedure TfrmMain.FormActivate(Sender: TObject);
begin
   If Not FirstTimeLogin Then
      Begin
        Application.Initialize;
        Application.CreateForm(TfrmLogin, frmLogin);
        frmLogin.Show;
        FirstTimeLogin := True;
   End;
end;

procedure TfrmMain.Usuarios1Click(Sender: TObject);
begin
if FormIsRunning('frmUsers') Then
  begin
        setActiveWindow(frmUsers.Handle);
        frmUsers.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmUsers,frmUsers);
        frmUsers.Show;
  end;
end;

procedure TfrmMain.Salir1Click(Sender: TObject);
begin
Application.Terminate;
end;

procedure TfrmMain.lblScrapClick(Sender: TObject);
begin
if FormIsRunning('frmScrap') Then
  begin
        setActiveWindow(frmScrap.Handle);
        frmScrap.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmScrap,frmScrap);
        frmScrap.Show;
  end;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var   Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin

    lblScrap.Caption := '';
    Application.ProcessMessages;
    lblScrap.Visible := False;
    application.ProcessMessages;
    application.ProcessMessages;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      // Update all items that are active or terminated in almacen task
      SQLStr := 'UPDATE tblScrap SET SCR_Activo = 1 ' +
                'FROM TBLSCRAP S ' +
                'INNER JOIN tblItemTasks I ON S.SCR_NewItem = I.ITE_Nombre ' +
                'WHERE I.TAS_ID = 2 AND (I.ITS_Status IS NOT NULL AND I.ITS_Status <> 0)';

      Conn.Execute(SQLStr);

      //select all items that need to be reschedule
      Qry.SQL.Clear;
      Qry.SQL.Text := 'SELECT COUNT(*) AS Ordenes FROM tblScrap WHERE SCR_Activo = 0';
      Qry.Open;

      if Qry['Ordenes'] > 0 then
      begin
          lblScrap.Visible := True;
          lblScrap.Caption := 'Scrap Pendiente de Reprogramar.';
          application.ProcessMessages;
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

procedure TfrmMain.ReportedeScrap1Click(Sender: TObject);
begin
if FormIsRunning('frmScrapPorcen') Then
  begin
        setActiveWindow(frmScrapPorcen.Handle);
        frmScrapPorcen.WindowState := wsNormal;
  end
else                      
  begin
        Application.CreateForm(TfrmScrapPorcen,frmScrapPorcen);
        frmScrapPorcen.Show;
  end;
end;

procedure TfrmMain.RelaciondeOrdenesdeCompra1Click(Sender: TObject);
begin
if FormIsRunning('frmRelacionOC') Then
  begin
        setActiveWindow(frmRelacionOC.Handle);
        frmRelacionOC.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmRelacionOC,frmRelacionOC);
        frmRelacionOC.Show;
  end;
end;

procedure TfrmMain.Screens1Click(Sender: TObject);
begin
if FormIsRunning('frmCatalogoScreens') Then
  begin
        setActiveWindow(frmCatalogoScreens.Handle);
        frmCatalogoScreens.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCatalogoScreens,frmCatalogoScreens);
        frmCatalogoScreens.Show;
  end;
end;

procedure TfrmMain.ReportedeRetrabajo1Click(Sender: TObject);
begin
if FormIsRunning('frmRetrabajo') Then
  begin
        setActiveWindow(frmRetrabajo.Handle);
        frmRetrabajo.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmRetrabajo,frmRetrabajo);
        frmRetrabajo.Show;
  end;
end;

procedure TfrmMain.ReportedeScrapenDinero1Click(Sender: TObject);
//var InputString: string;
begin
{InputString:= InputBox('Confirmacion...', 'Proporciona la clave : ', '');

if InputString <> 'Dinero123' then begin
        ShowMessage('Clave Incorrecta...');
        Exit;
end;
}
if FormIsRunning('frmScrapDinero') Then
  begin
        setActiveWindow(frmScrapDinero.Handle);
        frmScrapDinero.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmScrapDinero,frmScrapDinero);
        frmScrapDinero.Show;
  end;
end;

procedure TfrmMain.ipodeCambio1Click(Sender: TObject);
begin

if FormIsRunning('frmExchangeRate') Then
  begin
        setActiveWindow(frmExchangeRate.Handle);
        frmExchangeRate.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmExchangeRate,frmExchangeRate);
        frmExchangeRate.Show;
  end;

end;

procedure TfrmMain.ReportedeCumplimiento1Click(Sender: TObject);
begin

if FormIsRunning('frmCumplimiento') Then
  begin
        setActiveWindow(frmCumplimiento.Handle);
        frmCumplimiento.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCumplimiento,frmCumplimiento);
        frmCumplimiento.Show;
  end;

end;

procedure TfrmMain.ReportedePromediodeCumplimiento1Click(Sender: TObject);
begin
if FormIsRunning('frmPromCumpli') Then
  begin
        setActiveWindow(frmPromCumpli.Handle);
        frmPromCumpli.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmPromCumpli,frmPromCumpli);
        frmPromCumpli.Show;
  end;
end;

procedure TfrmMain.ReportedeRetrabajoenDinero1Click(Sender: TObject);
//var InputString: string;
begin
{InputString:= InputBox('Confirmacion...', 'Proporciona la clave : ', '');

if InputString <> 'Dinero123' then begin
        ShowMessage('Clave Incorrecta...');
        Exit;
end;
}
if FormIsRunning('frmRetrabajoDinero') Then
  begin
        setActiveWindow(frmRetrabajoDinero.Handle);
        frmRetrabajoDinero.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmRetrabajoDinero,frmRetrabajoDinero);
        frmRetrabajoDinero.Show;
  end;
end;

procedure TfrmMain.Categories1Click(Sender: TObject);
begin
if FormIsRunning('frmCatalogoCategories') Then
  begin
        setActiveWindow(frmCatalogoCategories.Handle);
        frmCatalogoCategories.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCatalogoCategories,frmCatalogoCategories);
        frmCatalogoCategories.Show;
  end;
end;

procedure TfrmMain.GruposdeUsuarios1Click(Sender: TObject);
begin
if FormIsRunning('frmCatalogoGrupos') Then
  begin
        setActiveWindow(frmCatalogoGrupos.Handle);
        frmCatalogoGrupos.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCatalogoGrupos,frmCatalogoGrupos);
        frmCatalogoGrupos.Show;
  end;
end;

procedure TfrmMain.Menu1Click(Sender: TObject);
begin
if FormIsRunning('frmMenuEditor') Then
  begin
        setActiveWindow(frmMenuEditor.Handle);
        frmMenuEditor.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmMenuEditor,frmMenuEditor);
        frmMenuEditor.Show;
  end;
end;

procedure TfrmMain.Permisos1Click(Sender: TObject);
begin
if FormIsRunning('frmCatalogoPermisos') Then
  begin
        setActiveWindow(frmCatalogoPermisos.Handle);
        frmCatalogoPermisos.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCatalogoPermisos,frmCatalogoPermisos);
        frmCatalogoPermisos.Show;
  end;
end;

procedure TfrmMain.BindMenu(userId : String);
var item,subItem : TMenuItem;
SQLStr, sLastItem : String;
Conn : TADOConnection;
Qry : TADOQuery;
Qry2 : TADOQuery;
begin
    Conn := nil;
    Qry := nil;
    Qry2 := nil;
    MainMenu1.Items.Clear;
    try
    begin
        Conn := TADOConnection.Create(nil);
        Conn.ConnectionString := frmMain.sConnString;
        Conn.LoginPrompt := False;
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;
        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;


        SQLStr := 'SELECT DISTINCT C.Category_Order,C.Category_ID, C.Category_Name ' +
                  'FROM tblCategories C ' +
                  'INNER JOIN tblCategory_Screens CS ON  C.Category_ID = CS.Category_ID ' +
                  'INNER JOIN tblScreens S ON  S.SCR_ID = CS.SCR_ID ' +
                  'INNER JOIN tblGroup_Screens GS ON GS.SCR_ID = S.SCR_ID ' +
                  'WHERE GS.Group_ID IN (SELECT Group_ID FROM tblUser_Groups WHERE USE_ID = ' +
                  userId + ') ' +
                  'ORDER BY C.Category_Order,C.Category_ID';

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        While not Qry.Eof do
        begin
            item := TMenuItem.Create(nil);
            item.Caption := Qry['Category_Name'];

            SQLStr := 'SELECT  DISTINCT CS.SCR_Order, S.SCR_ID, S.SCR_Name, S.SCR_FormName ' +
                      'FROM tblCategories C ' +
                      'INNER JOIN tblCategory_Screens CS ON  C.Category_ID = CS.Category_ID ' +
                      'INNER JOIN tblScreens S ON  S.SCR_ID = CS.SCR_ID ' +
                      'INNER JOIN tblGroup_Screens GS ON GS.SCR_ID = S.SCR_ID ' +
                      'WHERE GS.Group_ID IN (SELECT Group_ID FROM tblUser_Groups WHERE USE_ID = ' +
                      userId + ') ' + 'AND C.Category_Name = ' + QuotedStr(item.Caption) +
                      ' ORDER BY CS.SCR_Order, S.SCR_ID ';


            Qry2.SQL.Clear;
            Qry2.SQL.Text := SQLStr;
            Qry2.Open;

            While not Qry2.Eof do
            begin
                subItem := TMenuItem.Create(nil);
                subItem.Caption := Qry2['SCR_Name'];
                sLastItem := subItem.Caption;

                if sLastItem <> '-' then begin
                        assignEvent(Qry2['SCR_FormName'], subItem);
                end;

                item.Add(subItem);

                Qry2.Next;
            end;

            if not ( (sLastItem = '-') and (Qry2.RecordCount = 1) ) then begin
                    MainMenu1.Items.Add(item);
            end;

            Qry.Next;
        end;

        item := TMenuItem.Create(nil);
        item.Caption := 'Windows';

        subItem := TMenuItem.Create(nil);
        subItem.Caption := 'Cascada';
        subItem.OnClick := Cascade1Click;
        item.Add(subItem);
        
        subItem := TMenuItem.Create(nil);
        subItem.Caption := 'Tile';
        subItem.OnClick := Tile1Click;
        item.Add(subItem);

        subItem := TMenuItem.Create(nil);
        subItem.Caption := 'Cerrar Todas';
        subItem.OnClick := CloseAll1Click;
        item.Add(subItem);


        subItem := TMenuItem.Create(nil);
        subItem.Caption := '-';
        item.Add(subItem);


        MainMenu1.Items.Add(item);
    end
    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : BindGrid');
    end;

    if Qry2 <> nil then begin
      Qry2.Close;
      Qry2.Free;
    end;
    if Qry <> nil then begin
      Qry.Close;
      Qry.Free;
    end;
    if Conn <> nil then begin
      Conn.Close;
      Conn.Free
    end;

end;


procedure TfrmMain.assignEvent(formName : String; menuItem : TMenuItem);
begin

      if UT(formName) = UT('form1') then begin
        menuItem.OnClick := Monitor1Click;
      end
      else if UT(formName) = UT('frmCatalogoGrupos') then begin
        menuItem.OnClick := GruposdeUsuarios1Click;
      end
      else if UT(formName) = UT('frmCatalogoCategories') then begin
        menuItem.OnClick := Categories1Click;
      end
      else if UT(formName) = UT('frmClientes') then begin
        menuItem.OnClick := Clientes1Click;
      end
      else if UT(formName) = UT('frmCumplimiento') then begin
        menuItem.OnClick := ReportedeCumplimiento1Click;
      end
      else if UT(formName) = UT('frmEditor') then begin
        menuItem.OnClick := EditarOrden1Click;
      end
      else if UT(formName) = UT('frmEmpleados') then begin
        menuItem.OnClick := Empleados1Click;
      end
      else if UT(formName) = UT('frmEntrega') then begin
        menuItem.OnClick := Empleados2Click;
      end
      else if UT(formName) = UT('frmMenuEditor') then begin
        menuItem.OnClick := Menu1Click;
      end
      else if UT(formName) = UT('frmProductos') then begin
        menuItem.OnClick := Productos1Click;
      end
      else if UT(formName) = UT('frmPromCumpli') then begin
        menuItem.OnClick := ReportedePromediodeCumplimiento1Click;
      end
      else if UT(formName) = UT('frmRelacionOC') then begin
        menuItem.OnClick := RelaciondeOrdenesdeCompra1Click;
      end
      else if UT(formName) = UT('frmRetrabajo') then begin
        menuItem.OnClick := ReportedeRetrabajo1Click;
      end
      else if UT(formName) = UT('frmRetrabajoDinero') then begin
        menuItem.OnClick := ReportedeRetrabajoenDinero1Click;
      end
      else if UT(formName) = UT('frmRutas') then begin
        menuItem.OnClick := Rutas1Click;
      end
      else if UT(formName) = UT('frmScrapDinero') then begin
        menuItem.OnClick := ReportedeScrapenDinero1Click;
      end
      else if UT(formName) = UT('frmScrapPorcen') then begin
        menuItem.OnClick := ReportedeScrap1Click;
      end
      else if UT(formName) = UT('frmCatalogoScreens') then begin
        menuItem.OnClick := Screens1Click;
      end
      else if UT(formName) = UT('frmTareas') then begin
        menuItem.OnClick := areas2Click;
      end
      else if UT(formName) = UT('frmUsers') then begin
        menuItem.OnClick := Usuarios1Click;
      end
      else if UT(formName) = UT('frmVentas') then begin
        menuItem.OnClick := Ventas1Click;
      end
      else if UT(formName) = UT('frmYear') then begin
        menuItem.OnClick := EstablecerA1Click;
      end
      else if UT(formName) = UT('frmCatalogoGrupos') then begin
        menuItem.OnClick := GruposdeUsuarios1Click;
      end
      else if UT(formName) = UT('frmCatalogoPermisos') then begin
        menuItem.OnClick := Permisos1Click;
      end
      else if UT(formName) = UT('frmScrapEditor') then begin
        menuItem.OnClick := EditordeScrap1Click;
      end
      else if UT(formName) = UT('frmCargaTrabajo') then begin
        menuItem.OnClick := ReportedeCargadeTrabajo1Click;
      end
      else if UT(formName) = UT('frmFacturacion') then begin
        menuItem.OnClick := Facturacion2Click;
      end
      else if UT(formName) = UT('frmPendientesFact') then begin
        menuItem.OnClick := PendientesdeFacturar1Click;
      end
      else if UT(formName) = UT('frmProductividad') then begin
        menuItem.OnClick := ReportedeProductividad1Click;
      end
      else if UT(formName) = UT('frmContribuyente') then begin
        menuItem.OnClick := Contribuyente1Click;
      end
      else if UT(formName) = UT('frmMateriales') then begin
        menuItem.OnClick := Materiales1Click;
      end
      else if UT(formName) = UT('frmProductosTerminados') then begin
        menuItem.OnClick := ProductosTerminados1Click;
      end
      else if UT(formName) = UT('frmEditorRetrabajo') then begin
        menuItem.OnClick := EditordeRetrabajo1Click;
      end
      else if UT(formName) = UT('frmUnidadMedida') then begin
        menuItem.OnClick := UnidadesdeMedida1Click;
      end
      else if UT(formName) = UT('frmTipoMaterial') then begin
        menuItem.OnClick := TiposdeMaterial1Click;
      end
      else if UT(formName) = UT('frmPaises') then begin
        menuItem.OnClick := Paises1Click;
      end
      else if UT(formName) = UT('frmEntradas') then begin
        menuItem.OnClick := Entradas1Click;
      end
      else if UT(formName) = UT('frmProvedores') then begin
        menuItem.OnClick := Proveedores1Click;
      end
      else if UT(formName) = UT('frmSalidasAlmacen') then begin
        menuItem.OnClick := SalidasAlmacen1Click;
      end
      else if UT(formName) = UT('frmInventariosConf') then begin
        menuItem.OnClick := Configuracion1Click;
      end
      else if UT(formName) = UT('frmEntradasSalidasBorradas') then begin
        menuItem.OnClick := ReporteEntradasSalidasBorradas1Click;
      end
      else if UT(formName) = UT('frmPiezasTerminadas') then begin
        menuItem.OnClick := ReportedePiezasTerminadas1Click;
      end
      else if UT(formName) = UT('frmESAlmacen') then begin
        menuItem.OnClick := ReporteEntradasSalidasAlmacen1Click;
      end
      else if UT(formName) = UT('frmESLarco') then begin
        menuItem.OnClick := ReporteEntradasSalidasLarco1Click;
      end
      else if UT(formName) = UT('frmEscasos') then begin
        menuItem.OnClick := ReportedeMaterialesEscasos1Click;
      end
      else if UT(formName) = UT('frmSalidasLarco') then begin
        menuItem.OnClick := SalidasLarco1Click;
      end
      else if UT(formName) = UT('frmProdEmpleado') then begin
        menuItem.OnClick := ReporteProductividadEmpleado1Click;
      end
      else if UT(formName) = UT('frmProdEmpleadoDinero') then begin
        menuItem.OnClick := ReporteProductividadEmpleado2Click;
      end
      else if UT(formName) = UT('frmMaterialesPorOrden') then begin
        menuItem.OnClick := ReportedeMaterialesporOrdendeTrabajo1Click;
      end
      else if UT(formName) = UT('frmDiasInhabiles') then begin
        menuItem.OnClick := DiasInhabiles1Click;
      end
      else if UT(formName) = UT('frmCumplimientoTiempoEntrega') then begin
        menuItem.OnClick := ReportedeCumplimientodeTiempodeEntrega1Click;
      end
      else if UT(formName) = UT('frmCatalogoPlanos') then begin
        menuItem.OnClick := Planos1Click;
      end
      else if UT(formName) = UT('frmESStock') then begin
        menuItem.OnClick := EntradasStock1Click;
      end
      else if UT(formName) = UT('frmReporteESStock') then begin
        menuItem.OnClick := EntradasvsSalidas1Click;
      end
      else if UT(formName) = UT('frmReporteESPlano') then begin
        menuItem.OnClick := EntradasvsSalidasPorPlano1Click;
      end
      else if UT(formName) = UT('frmReporteTotalPiezasStock') then begin
        menuItem.OnClick := TotalPiezasStock1Click;
      end
      else if UT(formName) = UT('frmReportePiezasStock') then begin
        menuItem.OnClick := PiezasenStock1Click;
      end;
end;

procedure TfrmMain.EditordeScrap1Click(Sender: TObject);
begin
if FormIsRunning('frmScrapEditor') Then
  begin
        setActiveWindow(frmScrapEditor.Handle);
        frmScrapEditor.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmScrapEditor,frmScrapEditor);
        frmScrapEditor.Show;
  end;
end;

procedure TfrmMain.ReportedeCargadeTrabajo1Click(Sender: TObject);
begin
if FormIsRunning('frmCargaTrabajo') Then
  begin
        setActiveWindow(frmCargaTrabajo.Handle);
        frmCargaTrabajo.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCargaTrabajo,frmCargaTrabajo);
        frmCargaTrabajo.Show;
  end;
end;

procedure TfrmMain.Facturacion2Click(Sender: TObject);
begin
if FormIsRunning('frmFacturacion') Then
  begin
        setActiveWindow(frmFacturacion.Handle);
        frmFacturacion.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmFacturacion,frmFacturacion);
        frmFacturacion.Show;
  end;
end;

procedure TfrmMain.PendientesdeFacturar1Click(Sender: TObject);
begin
if FormIsRunning('frmPendientesFact') Then
  begin
        setActiveWindow(frmPendientesFact.Handle);
        frmPendientesFact.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmPendientesFact,frmPendientesFact);
        frmPendientesFact.Show;
  end;
end;

procedure TfrmMain.ReportedeProductividad1Click(Sender: TObject);
begin
if FormIsRunning('frmProductividad') Then
  begin
        setActiveWindow(frmProductividad.Handle);
        frmProductividad.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmProductividad,frmProductividad);
        frmProductividad.Show;
  end;
end;

procedure TfrmMain.Contribuyente1Click(Sender: TObject);
begin
if FormIsRunning('frmContribuyente') Then
  begin
        setActiveWindow(frmContribuyente.Handle);
        frmContribuyente.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmContribuyente,frmContribuyente);
        frmContribuyente.Show;
  end;

end;

procedure TfrmMain.Materiales1Click(Sender: TObject);
begin
if FormIsRunning('frmMateriales') Then
  begin
        setActiveWindow(frmMateriales.Handle);
        frmMateriales.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmMateriales,frmMateriales);
        frmMateriales.Show;
  end;
end;

procedure TfrmMain.ProductosTerminados1Click(Sender: TObject);
begin
if FormIsRunning('frmProductosTerminados') Then
  begin
        setActiveWindow(frmProductosTerminados.Handle);
        frmProductosTerminados.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmProductosTerminados,frmProductosTerminados);
        frmProductosTerminados.Show;
  end;

end;

procedure TfrmMain.Tile1Click(Sender: TObject);
begin
 Tile;
end;

procedure TfrmMain.Cascade1Click(Sender: TObject);
begin
  Cascade;
end;

procedure TfrmMain.CloseAll1Click(Sender: TObject);
var i: Integer;
begin
 for i:= 0 to MdiChildCount - 1 do
  MDIChildren[i].Close;

end;

procedure TfrmMain.EditordeRetrabajo1Click(Sender: TObject);
begin
if FormIsRunning('frmEditorRetrabajo') Then
  begin
        setActiveWindow(frmEditorRetrabajo.Handle);
        frmEditorRetrabajo.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEditorRetrabajo,frmEditorRetrabajo);
        frmEditorRetrabajo.Show;
  end;
end;

procedure TfrmMain.UnidadesdeMedida1Click(Sender: TObject);
begin
if FormIsRunning('frmUnidadMedida') Then
  begin
        setActiveWindow(frmUnidadMedida.Handle);
        frmUnidadMedida.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmUnidadMedida,frmUnidadMedida);
        frmUnidadMedida.Show;
  end;
end;

procedure TfrmMain.TiposdeMaterial1Click(Sender: TObject);
begin
if FormIsRunning('frmTipoMaterial') Then
  begin
        setActiveWindow(frmTipoMaterial.Handle);
        frmTipoMaterial.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmTipoMaterial,frmTipoMaterial);
        frmTipoMaterial.Show;
  end;
end;

procedure TfrmMain.Paises1Click(Sender: TObject);
begin
if FormIsRunning('frmPaises') Then
  begin
        setActiveWindow(frmPaises.Handle);
        frmPaises.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmPaises,frmPaises);
        frmPaises.Show;
  end;
end;

procedure TfrmMain.Entradas1Click(Sender: TObject);
begin
if FormIsRunning('frmEntradas') Then
  begin
        setActiveWindow(frmEntradas.Handle);
        frmEntradas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEntradas,frmEntradas);
        frmEntradas.Show;
  end;
end;

procedure TfrmMain.Proveedores1Click(Sender: TObject);
begin
if FormIsRunning('frmProvedores') Then
  begin
        setActiveWindow(frmProvedores.Handle);
        frmProvedores.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmProvedores,frmProvedores);
        frmProvedores.Show;
  end;
end;

procedure TfrmMain.SalidasAlmacen1Click(Sender: TObject);
begin
if FormIsRunning('frmSalidasAlmacen') Then
  begin
        setActiveWindow(frmSalidasAlmacen.Handle);
        frmSalidasAlmacen.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmSalidasAlmacen,frmSalidasAlmacen);
        frmSalidasAlmacen.Show;
  end;
end;

procedure TfrmMain.Configuracion1Click(Sender: TObject);
begin
if FormIsRunning('frmInventariosConf') Then
  begin
        setActiveWindow(frmInventariosConf.Handle);
        frmInventariosConf.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmInventariosConf, frmInventariosConf);
        frmInventariosConf.Show;
  end;
end;

procedure TfrmMain.ReporteEntradasSalidasBorradas1Click(Sender: TObject);
begin
if FormIsRunning('frmEntradasSalidasBorradas') Then
  begin
        setActiveWindow(frmEntradasSalidasBorradas.Handle);
        frmEntradasSalidasBorradas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEntradasSalidasBorradas, frmEntradasSalidasBorradas);
        frmEntradasSalidasBorradas.Show;
  end;
end;

procedure TfrmMain.ReportedePiezasTerminadas1Click(Sender: TObject);
begin
if FormIsRunning('frmPiezasTerminadas') Then
  begin
        setActiveWindow(frmPiezasTerminadas.Handle);
        frmPiezasTerminadas.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmPiezasTerminadas, frmPiezasTerminadas);
        frmPiezasTerminadas.Show;
  end;
end;

procedure TfrmMain.ReporteEntradasSalidasAlmacen1Click(Sender: TObject);
begin
if FormIsRunning('frmESAlmacen') Then
  begin
        setActiveWindow(frmESAlmacen.Handle);
        frmESAlmacen.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmESAlmacen, frmESAlmacen);
        frmESAlmacen.Show;
  end;
end;

procedure TfrmMain.ReporteEntradasSalidasLarco1Click(Sender: TObject);
begin
if FormIsRunning('frmESLarco') Then
  begin
        setActiveWindow(frmESLarco.Handle);
        frmESLarco.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmESLarco, frmESLarco);
        frmESLarco.Show;
  end;
end;

procedure TfrmMain.ReportedeMaterialesEscasos1Click(Sender: TObject);
begin
if FormIsRunning('frmEscasos') Then
  begin
        setActiveWindow(frmEscasos.Handle);
        frmEscasos.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmEscasos, frmEscasos);
        frmEscasos.Show;
  end;
end;

procedure TfrmMain.SalidasLarco1Click(Sender: TObject);
begin
if FormIsRunning('frmSalidasLarco') Then
  begin
        setActiveWindow(frmSalidasLarco.Handle);
        frmSalidasLarco.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmSalidasLarco, frmSalidasLarco);
        frmSalidasLarco.Show;
  end;
end;

procedure TfrmMain.ReporteProductividadEmpleado1Click(Sender: TObject);
begin
if FormIsRunning('frmProdEmpleado') Then
  begin
        setActiveWindow(frmProdEmpleado.Handle);
        frmProdEmpleado.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmProdEmpleado, frmProdEmpleado);
        frmProdEmpleado.Show;
  end;
end;

procedure TfrmMain.ReporteProductividadEmpleado2Click(Sender: TObject);
begin
if FormIsRunning('frmProdEmpleadoDinero') Then
  begin
        setActiveWindow(frmProdEmpleadoDinero.Handle);
        frmProdEmpleadoDinero.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmProdEmpleadoDinero, frmProdEmpleadoDinero);
        frmProdEmpleadoDinero.Show;
  end;
end;

procedure TfrmMain.ReportedeMaterialesporOrdendeTrabajo1Click(
  Sender: TObject);
begin
if FormIsRunning('frmMaterialesPorOrden') Then
  begin
        setActiveWindow(frmMaterialesPorOrden.Handle);
        frmMaterialesPorOrden.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmMaterialesPorOrden, frmMaterialesPorOrden);
        frmMaterialesPorOrden.Show;
  end;
end;

procedure TfrmMain.DiasInhabiles1Click(Sender: TObject);
begin
if FormIsRunning('frmDiasInhabiles') Then
  begin
        setActiveWindow(frmDiasInhabiles.Handle);
        frmDiasInhabiles.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmDiasInhabiles, frmDiasInhabiles);
        frmDiasInhabiles.Show;
  end;
end;

procedure TfrmMain.ReportedeCumplimientodeTiempodeEntrega1Click(
  Sender: TObject);
begin
if FormIsRunning('frmCumplimientoTiempoEntrega') Then
  begin
        setActiveWindow(frmCumplimientoTiempoEntrega.Handle);
        frmCumplimientoTiempoEntrega.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCumplimientoTiempoEntrega, frmCumplimientoTiempoEntrega);
        frmCumplimientoTiempoEntrega.Show;
  end;
end;

procedure TfrmMain.Planos1Click(Sender: TObject);
begin
if FormIsRunning('frmCatalogoPlanos') Then
  begin
        setActiveWindow(frmCatalogoPlanos.Handle);
        frmCatalogoPlanos.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmCatalogoPlanos, frmCatalogoPlanos);
        frmCatalogoPlanos.Show;
  end;
end;

procedure TfrmMain.EntradasStock1Click(Sender: TObject);
begin
if FormIsRunning('frmESStock') Then
  begin
        setActiveWindow(frmESStock.Handle);
        frmESStock.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmESStock, frmESStock);
        frmESStock.Show;
  end;
end;

procedure TfrmMain.EntradasvsSalidas1Click(Sender: TObject);
begin
if FormIsRunning('frmReporteESStock') Then
  begin
        setActiveWindow(frmReporteESStock.Handle);
        frmReporteESStock.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmReporteESStock, frmReporteESStock);
        frmReporteESStock.Show;
  end;
end;

procedure TfrmMain.EntradasvsSalidasPorPlano1Click(Sender: TObject);
begin
if FormIsRunning('frmReporteESPlano') Then
  begin
        setActiveWindow(frmReporteESPlano.Handle);
        frmReporteESPlano.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmReporteESPlano, frmReporteESPlano);
        frmReporteESPlano.Show;
  end;
end;

procedure TfrmMain.TotalPiezasStock1Click(Sender: TObject);
begin
if FormIsRunning('frmReporteTotalPiezasStock') Then
  begin
        setActiveWindow(frmReporteTotalPiezasStock.Handle);
        frmReporteTotalPiezasStock.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmReporteTotalPiezasStock, frmReporteTotalPiezasStock);
        frmReporteTotalPiezasStock.Show;
  end;
end;

procedure TfrmMain.PiezasenStock1Click(Sender: TObject);
begin
if FormIsRunning('frmReportePiezasStock') Then
  begin
        setActiveWindow(frmReportePiezasStock.Handle);
        frmReportePiezasStock.WindowState := wsNormal;
  end
else
  begin
        Application.CreateForm(TfrmReportePiezasStock, frmReportePiezasStock);
        frmReportePiezasStock.Show;
  end;
end;

end.
