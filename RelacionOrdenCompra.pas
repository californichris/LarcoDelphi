unit RelacionOrdenCompra;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, sndkey32,
  ExtCtrls, StdCtrls, CellEditors, ScrollView, ComCtrls,ComObj,
  CustomGridViewControl, CustomGridView, GridView, Menus,Clipbrd,LTCUtils,Larco_functions;

type
  TfrmRelacionOC = class(TForm)
    GridView1: TGridView;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label4: TLabel;
    Button1: TButton;
    cmbClientes: TComboBox;
    Button2: TButton;
    btnBuscar: TButton;
    cmbOrdenes: TComboBox;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    Imprimir: TButton;
    Stock: TButton;
    PopupMenu2: TPopupMenu;
    Copiar1: TMenuItem;
    OpenDialog1: TOpenDialog;
    Editar: TButton;
    Borrar: TButton;
    GroupBox2: TGroupBox;
    chkTotal: TCheckBox;
    txtCantidad: TEdit;
    Label1: TLabel;
    Cancel: TButton;
    OK: TButton;
    lblAct: TLabel;
    lblCount: TLabel;
    ProgressBar1: TProgressBar;
    lblStep: TLabel;
    Label3: TLabel;
    cmbPartes: TComboBox;
    Label5: TLabel;
    cmbCompra: TComboBox;
    Button5: TButton;
    GroupBox5: TGroupBox;
    gvClientes: TGridView;
    CheckBox1: TCheckBox;
    btnOK: TButton;
    btnTodos: TButton;
    Button3: TButton;
    txtCliente: TEdit;
    GroupBox3: TGroupBox;
    gvOrdenes: TGridView;
    CheckBox2: TCheckBox;
    btnOK2: TButton;
    btnTodos2: TButton;
    txtOrden: TEdit;
    Button7: TButton;
    Label6: TLabel;
    ddlAnio: TComboBox;
    deRecibido1: TDateEditor;
    deRecibido2: TDateEditor;
    chkRecibido: TCheckBox;
    Label7: TLabel;
    cmbProductos: TComboBox;
    cmbPlanos: TComboBox;
    Label8: TLabel;
    ExportROC: TButton;
    gvColorCode: TGridView;
    lblWeek: TLabel;
    Label9: TLabel;
    cmbUrgente: TComboBox;
    procedure IdleLoop; virtual;
    function FormIsRunning(FormName: String):Boolean;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    function getBitVal(value : String):String;
    procedure FormCreate(Sender: TObject);
    procedure BindClientes();
    Procedure BindOrdenes(Query: String);
    procedure BindGrid();
    procedure ExportGrid(Grid:TGridView;sFileName: String);
    procedure ExportGridROC(Grid: TGridView;sFileName: String);
    function BitToBoolean(Value:String):Boolean;
    procedure Copiar1Click(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure ImportFile(FileName: String);
    Function RightStrOf(Str : String; C : String): String;
    procedure GridView1SelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure BorrarClick(Sender: TObject);
    procedure StockClick(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure CancelClick(Sender: TObject);
    procedure OKClick(Sender: TObject);
    procedure chkTotalClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnTodosClick(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure btnOK2Click(Sender: TObject);
    procedure btnTodos2Click(Sender: TObject);
    procedure chkRecibidoClick(Sender: TObject);
    procedure cmbPartesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cmbCompraKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ExportROCClick(Sender: TObject);
    procedure cmbCompraKeyPress(Sender: TObject; var Key: Char);
    procedure cmbPartesKeyPress(Sender: TObject; var Key: Char);
    function getClient(clientId : String; date: String) :String;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelacionOC: TfrmRelacionOC;
  gsYear: String;
  giCantidad, giCantCliente:Integer;
  Qry : TADOQuery;
  Conn : TADOConnection;
  bFirst : Boolean;

implementation

uses Main, ImpresionOrden, Ventas, ReporteRelacion, ReporteOC,
  DetalleOrdenes;

{$R *.dfm}

procedure TfrmRelacionOC.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

procedure TfrmRelacionOC.FormCreate(Sender: TObject);
begin
    bFirst := False;
    gsYear := RightStr(getFormYear(frmMain.sConnString,Self.Name),2);
    Timer1.Interval := frmMain.iIntervalo;

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := frmMain.sConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    Qry.SQL.Clear;
    Qry.SQL.Text := 'SELECT DISTINCT LEFT(ITE_NOMBRE,2) AS Year FROM tblOrdenes ORDER BY LEFT(ITE_NOMBRE,2) DESC';
    Qry.Open;

    ddlAnio.Clear;
    While not Qry.Eof do
    Begin
        ddlAnio.Items.Add('20' + VarToStr(Qry['Year']));
        Qry.Next;
    end;

    ddlAnio.ItemIndex := 0;

    cmbClientes.Text := 'Todos';
    cmbOrdenes.Text := 'Todos';
    cmbPartes.Text := 'Todos';
    cmbCompra.Text := 'Todos';
    cmbProductos.Text := 'Todos';
    cmbPlanos.Text := 'Todos';
    cmbUrgente.Text := 'Todos';

    cmbUrgente.Items.Add('Todos');
    cmbUrgente.Items.Add('No');
    cmbUrgente.Items.Add('Si');

    deRecibido1.Date := Now;
    deRecibido2.Date := Now;

    BindClientes();
    CheckBox1.Checked := False;
    btnOKClick(nil);
    BindGrid();


    gvColorCode.ClearRows;
    gvColorCode.AddRow();
    gvColorCode.Cells[1,gvColorCode.RowCount -1] := 'Stock';
    gvColorCode.Cell[0, gvColorCode.RowCount -1].Color := clBlue;

    gvColorCode.AddRow();
    gvColorCode.Cells[1,gvColorCode.RowCount -1] := 'Terminado';
    gvColorCode.Cell[0, gvColorCode.RowCount -1].Color := clYellow;

    gvColorCode.AddRow();
    gvColorCode.Cells[1,gvColorCode.RowCount -1] := 'Sin Plano';
    gvColorCode.Cell[0, gvColorCode.RowCount -1].Color := clSkyBlue;

    gvColorCode.AddRow();
    gvColorCode.Cells[1,gvColorCode.RowCount -1] := 'Stock Parcial';
    gvColorCode.Cell[0, gvColorCode.RowCount -1].Color := clLime;

    gvColorCode.AddRow();
    gvColorCode.Cells[1,gvColorCode.RowCount -1] := 'Mezclado';
    gvColorCode.Cell[0, gvColorCode.RowCount -1].Color := clSilver;

end;

procedure TfrmRelacionOC.BindGrid();
var SQLStr,SQLWhere,sItem,sOrden,additionalWhere, mezclado, prevOrden : String;
Col,i : Integer;
begin
    lblCount.Caption := '';
    giCantidad := 0;
    giCantCliente := 0;
    additionalWhere := '';

    if cmbUrgente.Text <> 'Todos' then
    begin
        if cmbUrgente.Text = 'No' then
        begin
          additionalWhere := ' O.Urgente = 0 '
        end
        else if cmbUrgente.Text = 'Si' then
        begin
          additionalWhere := ' O.Urgente = 1 '
        end;
    end;

    if chkRecibido.Checked = True then
    begin
        if(additionalWhere <> '') then
        begin
            additionalWhere := additionalWhere + ' AND '
        end;

        additionalWhere := additionalWhere + ' (O.Recibido >= ' + QuotedStr(deRecibido1.Text) + ' AND O.Recibido <= ' + QuotedStr(deRecibido2.Text  + ' 23:59:59.99') + ')';
    end;

    SQLStr := 'Relacion_Orden_Compra ' + QuotedStr(txtCliente.Text) + ','
              + QuotedStr(txtOrden.Text) + ',' + QuotedStr(cmbCompra.Text) + ','
              + QuotedStr(cmbPartes.Text) + ',' + QuotedStr(cmbProductos.Text) + ','
              + QuotedStr(cmbPlanos.Text) ;

    if bFirst = False then begin
        SQLStr := SQLStr + ',1,' + QuotedStr(RightStr(ddlAnio.Text,2)) + ',' + QuotedStr(additionalWhere);
        bFirst := True;
    end
    else begin
        SQLStr := SQLStr + ',0,' + QuotedStr(RightStr(ddlAnio.Text,2)) + ',' + QuotedStr(additionalWhere);
    end;

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    sOrden := cmbOrdenes.Text;
    cmbOrdenes.Clear;
    cmbOrdenes.Sorted := True;
    cmbOrdenes.Text := sOrden;

    sOrden := cmbCompra.Text;
    cmbCompra.Clear;
    cmbCompra.Sorted := True;
    cmbCompra.Text := sOrden;

    sOrden := cmbPartes.Text;
    cmbPartes.Clear;
    cmbPartes.Sorted := True;
    cmbPartes.Text := sOrden;

    sOrden := cmbProductos.Text;
    cmbProductos.Clear;
    cmbProductos.Sorted := True;
    cmbProductos.Text := sOrden;

    sOrden := cmbPlanos.Text;
    cmbPlanos.Clear;
    cmbPlanos.Sorted := True;
    cmbPlanos.Text := sOrden;

    GridView1.ClearRows;

    lblStep.Visible := True;
    ProgressBar1.Visible := True;
    ProgressBar1.Position := 0;
    ProgressBar1.Max := Qry.RecordCount;
    ProgressBar1.Step := 1;
    application.ProcessMessages;

    prevOrden := '';
    while not Qry.Eof do
    begin
        mezclado := VarToStr(Qry['Mezclado']);
        sOrden := VarToStr(Qry['Orden']);
        if (mezclado = '1') and (prevOrden = sOrden) then begin
          GridView1.Cells[18,GridView1.RowCount -1] := GridView1.Cells[18,GridView1.RowCount -1] + ',' + VarToStr(Qry['MO_ITE_NOmbre']);

          Qry.Next;
          prevOrden := sOrden;
          Continue;
        end;
        prevOrden := sOrden;

        GridView1.AddRow(1);
        GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(Qry['Orden']);
        GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(Qry['Recibido']);
        GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(Qry['OrdenCompra']);
        GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(Qry['Cantidad']);
        giCantidad := giCantidad + StrToInt(GridView1.Cells[3,GridView1.RowCount -1]);
        GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(Qry['Cliente']);
        giCantCliente := giCantCliente + StrToInt(GridView1.Cells[4,GridView1.RowCount -1]);

        GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(Qry['Descripcion']);
        GridView1.Cells[6,GridView1.RowCount -1] := VarToStr(Qry['Plano']);
        GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(Qry['Numero']);
        GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(Qry['Entrega']);
        GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(Qry['Interna']);
        GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(Qry['Tarea']);
        GridView1.Cells[11,GridView1.RowCount -1] := VarToStr(Qry['Status']);
        GridView1.Cells[12,GridView1.RowCount -1] := VarToStr(Qry['Total']);
        GridView1.Cells[13,GridView1.RowCount -1] := getBitVal(VarToStr(Qry['Dolares']));
        GridView1.Cells[14,GridView1.RowCount -1] := getBitVal(VarToStr(Qry['Stock']));
        GridView1.Cells[15,GridView1.RowCount -1] := getBitVal(VarToStr(Qry['StockParcial']));
        GridView1.Cells[16,GridView1.RowCount -1] := VarToStr(Qry['StockParcialCantidad']);
        GridView1.Cells[17,GridView1.RowCount -1] := getBitVal(VarToStr(Qry['Mezclado']));
        GridView1.Cells[18,GridView1.RowCount -1] := VarToStr(Qry['MO_ITE_NOmbre']);
        GridView1.Cells[19,GridView1.RowCount -1] := VarToStr(Qry['Revision']);
        GridView1.Cells[20,GridView1.RowCount -1] := VarToStr(Qry['Requisicion']);

        if (GridView1.Cells[10,GridView1.RowCount -1] = 'VentasFinal') and (GridView1.Cells[11,GridView1.RowCount -1] = 'Terminado') then
           begin
                for Col := 0 to GridView1.Columns.Count - 1 do
                        GridView1.Cell[Col, GridView1.RowCount -1].Color := clYellow;

           end;

        if (GridView1.Cells[10,GridView1.RowCount -1] = 'Ventas') and (GridView1.Cells[11,GridView1.RowCount -1] = 'Activo') then
           begin
                for Col := 0 to GridView1.Columns.Count - 1 do
                        GridView1.Cell[Col, GridView1.RowCount -1].Color := clSkyBlue;

           end;


        if GridView1.Cells[14,GridView1.RowCount -1] = 'Si' then  //stock
           begin
                for Col := 0 to GridView1.Columns.Count - 1 do
                        GridView1.Cell[Col, GridView1.RowCount -1].Color := clBlue;

           end;

        if GridView1.Cells[15,GridView1.RowCount -1] = 'Si' then  //Stock Parcial
           begin
                for Col := 0 to GridView1.Columns.Count - 1 do
                        GridView1.Cell[Col, GridView1.RowCount -1].Color := clLime;

           end;

        if GridView1.Cells[17,GridView1.RowCount -1] = 'Si' then  //Mezclado
           begin
                for Col := 0 to GridView1.Columns.Count - 1 do
                        GridView1.Cell[Col, GridView1.RowCount -1].Color := clSilver;

           end;

        sItem := Copy(GridView1.Cells[0,GridView1.RowCount -1],5,3);
        if cmbOrdenes.Items.IndexOf(sItem) = -1 then
        begin
                cmbOrdenes.Items.Add(sItem);
        end;

        sItem := Trim(GridView1.Cells[2,GridView1.RowCount -1]);
        if ( (cmbCompra.Items.IndexOf(sItem) = -1) and (sItem <> '')) then
        begin
                cmbCompra.Items.Add(sItem);
        end;

        sItem := Trim(GridView1.Cells[6,GridView1.RowCount -1]);
        if ( (cmbPlanos.Items.IndexOf(sItem) = -1) and (sItem <> '') ) then
        begin
                cmbPlanos.Items.Add(sItem);
        end;

        sItem := Trim(GridView1.Cells[7,GridView1.RowCount -1]);
        if ( (cmbPartes.Items.IndexOf(sItem) = -1) and (sItem <> '') ) then
        begin
                cmbPartes.Items.Add(sItem);
        end;

        sItem := Trim(GridView1.Cells[5,GridView1.RowCount -1]);
        if ( (cmbProductos.Items.IndexOf(sItem) = -1) and (sItem <> '') ) then
        begin
                cmbProductos.Items.Add(sItem);
        end;

        Qry.Next;

        ProgressBar1.StepIt;
        application.ProcessMessages;
    End;

    cmbPartes.Items.Insert(0,'Todos');
    cmbCompra.Items.Insert(0,'Todos');
    cmbProductos.Items.Insert(0, 'Todos');
    cmbPlanos.Items.Insert(0, 'Todos');

    gvOrdenes.ClearRows;
    for i:= 0 to cmbOrdenes.Items.Count - 1 do begin
        gvOrdenes.AddRow(1);
        gvOrdenes.Cells[0,gvOrdenes.RowCount -1] := cmbOrdenes.Items.Strings[i];
    end;


    lblStep.Visible := False;
    ProgressBar1.Visible := False;
    application.ProcessMessages;
    lblCount.Caption := 'Total de Ordenes : ' + IntToStr(GridView1.RowCount) +
                        ' Cantidad Total Larco : ' + IntToStr(giCantidad) +
                        ' Cantidad Total Cliente : ' + IntToStr(giCantCliente);

end;

procedure TfrmRelacionOC.BindClientes();
var Qry2 : TADOQuery;
SQLStr : String;
slClientes : TStringList;
begin
    slClientes := TStringList.Create;
    slClientes.CommaText := '010,060,062,162,699,799,862,899,999,960';
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    SQLStr := 'SELECT Distinct Clave FROM tblClientes Order By Clave';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    gvClientes.ClearRows;
    While not Qry2.Eof do
    Begin
        gvClientes.AddRow(1);
        gvClientes.Cells[0,gvClientes.RowCount -1] := VarToStr(Qry2['Clave']);
        if (slClientes.IndexOf(VarToStr(Qry2['Clave'])) = -1) then begin
                gvClientes.Cell[1,gvClientes.RowCount -1].AsBoolean := True;
        end;
        Qry2.Next;
    End;

    Qry2.Close;
end;

procedure TfrmRelacionOC.BindOrdenes(Query: String);
var Qry2 : TADOQuery;
SQLStr : String;
begin
    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;


    SQLStr := 'SELECT DISTINCT SUBSTRING(O.ITE_Nombre,8,3) AS Orden ' +
              'FROM tblOrdenes O ' +
              'INNER JOIN tblItemTasks I ON O.ITE_ID = I.ITE_ID AND ITS_DTStart IS NOT NULL AND ITS_DTStop IS NULL ' +
              'INNER JOIN tblTareas T ON T.[ID] = I.TAS_ID ';

    if Query <> '' then SQLStr := SQLStr + ' WHERE ' + Query;

    SQLStr := SQLStr + ' ORDER BY SUBSTRING(O.ITE_Nombre,8,3) ';

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    cmbOrdenes.Items.Clear;
    cmbOrdenes.Items.Add('Todos');
    While not Qry2.Eof do
    Begin
        cmbOrdenes.Items.Add(Qry2['Orden']);
        Qry2.Next;
    End;

    Qry2.Close;
end;

function TfrmRelacionOC.BitToBoolean(Value:String):Boolean;
begin
        Result := False;
        if Value = '-1' then
                result := True;
end;

procedure TfrmRelacionOC.Copiar1Click(Sender: TObject);
begin
        if PopupMenu2.PopupComponent = GridView1 then
           Clipboard.AsText := GridView1.Cells[0,GridView1.SelectedRow]

end;

procedure TfrmRelacionOC.btnBuscarClick(Sender: TObject);
begin
GroupBox1.Enabled := False;
BindGrid();
GroupBox1.Enabled := True;
end;

procedure TfrmRelacionOC.Button1Click(Sender: TObject);
var sFileName: String;
begin
  if GridView1.RowCount = 0 then
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

    ExportGrid(GridView1,sFileName);

  end;

end;

procedure TfrmRelacionOC.Button5Click(Sender: TObject);
var sFileName: String;
begin

  OpenDialog1.Filter := 'Excel files (*.xls)|*.XLS';
  if OpenDialog1.Execute then
  begin
    sFileName := OpenDialog1.FileName;

    if UpperCase(Trim(rightStr(sFileName,4))) <> '.XLS' Then
          sFileName := sFileName + '.xls';

    ImportFile(sFileName);

  end;

end;

procedure TfrmRelacionOC.ImportFile(FileName: String);
const
  xlWorkSheet = -4167;

  aHeaders: array[0..17] of PChar =
   ('Orden','TipoProceso','CantidadCliente','CantidadLarco','Descripcion','Numero',
   'Terminal','OrdenCompra','FechaRecibido','FechaInterna','FechaCompromiso',
   'Nombre','Aprobacion','Unitario','Total','Dolares','Observaciones','Otras');

var XApp : Variant;
Sheet : Variant;
col,Row :Integer;
bExists,bHeaders : Boolean;
sColumn,SQLStr,sFileAs,sData,sMsg : String;
slData : TStringList;
QryFile : TADOQuery;
begin
      QryFile := TADOQuery.Create(nil);
      QryFile.Connection :=Conn;

      slData := TStringList.Create;

      if  Length(RightStrOf(FileName,'\')) > 50 then
      begin
         Showmessage('El nombre del archivo tiene que ser menor de 50 caracteres.');
         Exit;
      end;

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


      Try
      Begin
          XApp.Workbooks.open(FileName);
          Sheet := XApp.Workbooks[1].WorkSheets[1];

          // validar los headers para ver si estan en el orden que los necesito
          // y si estan todos los que necesito
          bHeaders := False;
          bExists := True;
          Col := 1;
          slData.Clear;
          sData := '';
          While bExists do
          Begin
              sColumn := TrimRight(sheet.Cells[1,Col]);

              If sColumn = '' Then
                      Break
              else
                      sData := sData + sColumn + ',';

              if UT(sColumn) <> UT(aHeaders[Col - 1]) then
                begin
                        bHeaders := True;
                        Break;
                end;


               Col := Col + 1;
          end;

          sData := LeftStr(sData,Length(sData) - 1);
          slData.Add(sData);


          if bHeaders then
            begin
                ShowMessage('Las columnas no estan en el orden solicitado.');
            end
          else if (Col - 1) <> 18 then
            begin
                ShowMessage('El archivo debe de tener 18 columnas y tiene ' + IntToStr(Col-1) + '.');
            end;

//***     validar celdas vacias agregar espacios para arreglar el problema del insert
          Row := 2;
          While bExists do
          Begin
              sData := '';
              for Col := 1 to 18 do
                begin
                    sColumn := TrimRight(sheet.Cells[Row,Col]);

                    If ( (sColumn = '') and (Col <> 1) ) Then
                        sheet.Cells[Row,Col] := ' '
                    else if ( (sColumn = '') and (Col = 1) ) then
                     begin
                        bExists := False;
                        break;
                     end;

                     sColumn := sheet.Cells[Row,Col];
                     sData := sData + sColumn + ',';
                end;

              sData := LeftStr(sData,Length(sData) - 1);
              slData.Add(sData);
              Row := Row + 1;
          end;


           //************* Termina validacion de headers ************************

           sFileAs := StartDDir +
           LeftStr(RightStrOf(FileName,'\'),Length(RightStrOf(FileName,'\')) - 4) + '.csv';

           slData.SaveToFile(sFileAs);
           //XApp.ActiveWorkBook.SaveAs(sFileAs,6); //CSV
           //XApp.ActiveWorkBook.SaveAs(sFileAs,-4158);    //txt tab delimited

           XApp.ActiveWorkBook.close();
           Sheet := Unassigned;
           XApp.Quit;
           XApp := Unassigned;

           SQLStr := 'Upload_File ' + QuotedStr(sFileAs) + ',' +
           QuotedStr(RightStrOf(FileName,'\')) + ',' + QuotedStr(frmMain.StatusBar.Panels[2].Text) +
           ',' + QuotedStr(gsYear);

           QryFile.SQL.Clear;
           QryFile.SQL.Text := SQLStr;
           QryFile.Open;

           sMsg := 'Total de ordenes : ' + IntToStr(Row - 3) + chr(13) +
                   'Ordenes Importadas :' + IntToStr(Row - 3 - QryFile.RecordCount) + chr(13) +
                   'Ordenes Rechazadas :' + IntToStr(QryFile.RecordCount);
                   
           if QryFile.RecordCount > 0 then begin
                   sMsg := sMsg + chr(13) +'Deseas ver el detalle de las ordenes rechazadas';

                   if MessageDlg(sMsg,mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                   begin
                        Application.CreateForm(TfrmDetalle,frmDetalle);

                        with frmDetalle do begin
                              GridView1.ClearRows;
                              QryFile.First;
                              while not QryFile.Eof do
                              begin
                                  GridView1.AddRow(1);
                                  GridView1.Cells[0,GridView1.RowCount -1] := VarToStr(QryFile['ITE_Nombre']);
                                  GridView1.Cells[1,GridView1.RowCount -1] := VarToStr(QryFile['Razon']);
                                  GridView1.Cells[2,GridView1.RowCount -1] := VarToStr(QryFile['TipoProceso']);
                                  GridView1.Cells[3,GridView1.RowCount -1] := VarToStr(QryFile['Requerida']);
                                  GridView1.Cells[4,GridView1.RowCount -1] := VarToStr(QryFile['Ordenada']);
                                  GridView1.Cells[5,GridView1.RowCount -1] := VarToStr(QryFile['Producto']);

                                  GridView1.Cells[7,GridView1.RowCount -1] := VarToStr(QryFile['Numero']);
                                  GridView1.Cells[8,GridView1.RowCount -1] := VarToStr(QryFile['Terminal']);
                                  GridView1.Cells[9,GridView1.RowCount -1] := VarToStr(QryFile['OrdenCompra']);
                                  GridView1.Cells[10,GridView1.RowCount -1] := VarToStr(QryFile['Recibido']);
                                  GridView1.Cells[11,GridView1.RowCount -1] := VarToStr(QryFile['Interna']);
                                  GridView1.Cells[12,GridView1.RowCount -1] := VarToStr(QryFile['Entrega']);
                                  GridView1.Cells[13,GridView1.RowCount -1] := VarToStr(QryFile['Nombre']);
                                  GridView1.Cells[14,GridView1.RowCount -1] := VarToStr(QryFile['Unitario']);
                                  GridView1.Cells[15,GridView1.RowCount -1] := VarToStr(QryFile['Aprobacion']);
                                  GridView1.Cells[16,GridView1.RowCount -1] := VarToStr(QryFile['Total']);
                                  GridView1.Cells[17,GridView1.RowCount -1] := VarToStr(QryFile['Dolares']);
                                  GridView1.Cells[18,GridView1.RowCount -1] := VarToStr(QryFile['Observaciones']);
                                  GridView1.Cells[19,GridView1.RowCount -1] := VarToStr(QryFile['Otras']);
                                  QryFile.Next;
                              end;
                        end;
                        frmDetalle.ShowModal;
                   end;
           end
           else begin
              MessageDlg(sMsg , mtInformation,[mbOk], 0);
           end;

           BindGrid();
           QryFile.Close;
      end
      finally
          //Sheet := Unassigned;
          //XApp.Quit;
          //XApp := Unassigned;
      end;

end;

Function TfrmRelacionOC.RightStrOf(str: String; C: String):String;
Var sRigth : String;
    i : Integer;
Begin
  for i:= Length(str) Downto 0 do
    Begin
      if str[i] = C then
        Begin
          sRigth := Copy(str,i+1,Length(str));
          Break;
        End;
    End;

  Result := sRigth
End;


procedure TfrmRelacionOC.GridView1SelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{If (GridView1.Cells[11,ARow] <> '') and (GridView1.Cells[15,ARow] = '0' ) Then
Begin
        Stock.Enabled := True;
        Imprimir.Enabled := True;
        Editar.Enabled := True;
        Borrar.Enabled := True;
End
Else
Begin
        Stock.Enabled := False;
        Imprimir.Enabled := False;
        Editar.Enabled := False;
        Borrar.Enabled := False;
End;
}

end;

procedure TfrmRelacionOC.BorrarClick(Sender: TObject);
//var SQLStr : String;
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{    if MessageDlg('Estas seguro que quieres borrar la orden : ' +
                  GridView1.Cells[0,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;


    SQLStr := 'DELETE FROM tblStockOrdenes WHERE ITE_Nombre = ' +
    QuotedStr(gsYear + '-' + GridView1.Cells[0,GridView1.SelectedRow]);

    Conn.Execute(SQLStr);
    Stock.Enabled := False;
    Imprimir.Enabled := False;
    Editar.Enabled := False;
    Borrar.Enabled := False;
    BindGrid();
    }
end;

procedure TfrmRelacionOC.StockClick(Sender: TObject);
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{    if MessageDlg('Estas seguro que quieres marcar la orden : ' +
                  GridView1.Cells[0,GridView1.SelectedRow] + ' como tomada del stock?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

    chkTotal.Checked := True;
    txtCantidad.Text := '';
    txtCantidad.Enabled := false;
    GroupBox2.Visible := True;

    Stock.Enabled := False;
    Imprimir.Enabled := False;
    Editar.Enabled := False;
    Borrar.Enabled := False;
}
end;

procedure TfrmRelacionOC.ImprimirClick(Sender: TObject);
//var SQLStr,sOrden : String;
//Qry2,Qry3 : TADOQuery;
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{    if MessageDlg('Estas seguro que quieres Imprimir la orden : ' +
                  GridView1.Cells[0,GridView1.SelectedRow] +
                  ', al imprimirla se va a meter al sistema la cantidad completa?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry3 := TADOQuery.Create(nil);
    Qry3.Connection :=Conn;


    sOrden := gsYear + '-' + GridView1.Cells[0,GridView1.SelectedRow];
    SQLStr := 'UPDATE tblStockOrdenes SET Programado = 1 WHERE ITE_Nombre = ' + QuotedStr(sOrden);

    Conn.Execute(SQLStr);

    SQLStr := 'SELECT * FROM tblStockOrdenes WHERE ITE_Nombre = ' + QuotedStr(sOrden);

    Qry2.SQL.Clear;
    Qry2.SQL.Text := SQLStr;
    Qry2.Open;

    SQLStr := 'Insert_Orden ' + QuotedStr(sOrden) + ',' + QuotedStr(VarToStr(Qry2['TipoProceso'])) +
              ',' + VarToStr(Qry2['Requerida']) + ',' + VarToStr(Qry2['Ordenada']) + ',' + QuotedStr(VarToStr(Qry2['Producto'])) +
              ',' + QuotedStr(VarToStr(Qry2['Numero'])) + ',' + QuotedStr(VarToStr(Qry2['Terminal'])) +
              ',' + QuotedStr(VarToStr(Qry2['Entrega'])) + ',' + QuotedStr(VarToStr(Qry2['Recibido'])) +
              ',' + QuotedStr(VarToStr(Qry2['Interna'])) +
              ',' + QuotedStr(VarToStr(Qry2['Nombre'])) + ',' + VarToStr(Qry2['Aprobacion']) +
              ',' + QuotedStr(VarToStr(Qry2['Observaciones'])) + ',' + QuotedStr(VarToStr(Qry2['Otras'])) +
              ',' + VarToStr(Qry2['Unitario']) + ',' + VarToStr(Qry2['Total']) + ',' + QuotedStr('Ventas') +
              //',' + QuotedStr('System') + ',' + QuotedStr(GetLocalIP) +
              ',' + QuotedStr(frmMain.sUserLogin) + ',' + QuotedStr(GetLocalIP) +
              ',' + QuotedStr(VarToStr(Qry2['OrdenCompra'])) + ',' + VarToStr(Qry2['Dolares']);

    Qry3.SQL.Clear;
    Qry3.SQL.Text := SQLStr;
    Qry3.Open;

    if VarToStr(Qry3['ERROR']) = '-1' Then
      begin
            ShowMessage(VarToStr(Qry3['MSG']));
            Exit;
      end;

    Qry3.Close;

    if Qry2.RecordCount <= 0 then
                Exit;

    Application.Initialize;
    Application.CreateForm(TqrImpresionOrden,qrImpresionOrden);

    qrImpresionOrden.QROrden.Caption := VarToStr(Qry2['ITE_Nombre']);
    qrImpresionOrden.QROrden2.Caption := VarToStr(Qry2['ITE_Nombre']);

    qrImpresionOrden.QRCode1.Caption := '*' + VarToStr(Qry2['ITE_Nombre']) + '*';
    qrImpresionOrden.QRCode2.Caption := '*' + VarToStr(Qry2['ITE_Nombre']) + '*';
    qrImpresionOrden.QRSemana.Caption := '';
    qrImpresionOrden.QRNumero.Caption := VarToStr(Qry2['Numero']);
    qrImpresionOrden.QRTerminal.Caption := VarToStr(Qry2['Terminal']);
    qrImpresionOrden.QRRecibido.Caption := VarToStr(Qry2['Recibido']);
    qrImpresionOrden.QREntrega.Caption := VarToStr(Qry2['Interna']);
    qrImpresionOrden.QREntrega2.Caption := VarToStr(Qry2['Interna']);
    qrImpresionOrden.QRNombre.Caption := VarToStr(Qry2['Nombre']);
    qrImpresionOrden.QRFirma.Caption := '';
    qrImpresionOrden.QRObs.Caption := VarToStr(Qry2['Observaciones']);
    qrImpresionOrden.QRDesc.Caption := VarToStr(Qry2['Producto']);
    qrImpresionOrden.QRCompra.Caption := VarToStr(Qry2['OrdenCompra']);
    qrImpresionOrden.QRProceso.Caption := VarToStr(Qry2['TipoProceso']);
    qrImpresionOrden.QRCantidad.Caption := VarToStr(Qry2['Ordenada']);

    qrImpresionOrden.QRMsg.Caption := 'Forma: Larco-015' + #13 +
                                      'Nivel de Revisi�n: D' + #13 +
                                      'Retenci�n: 1 a�o+uso';
    //qrImpresionOrden.Print;
    qrImpresionOrden.Preview;
    qrImpresionOrden.Free;

    Qry2.Close;

    BindGrid();}
end;

procedure TfrmRelacionOC.CancelClick(Sender: TObject);
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{    GroupBox2.Visible := False;


    Stock.Enabled := True;
    Imprimir.Enabled := True;
    Editar.Enabled := True;
    Borrar.Enabled := True;
}
end;

procedure TfrmRelacionOC.OKClick(Sender: TObject);
//var SQLStr,sOrden : String;
//iCant : Integer;
//Qry2,Qry3 : TADOQuery;
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{    if MessageDlg('Estas seguro que quieres tomar del stock ' + txtCantidad.Text + ' piezas para la orden ' +
                  GridView1.Cells[0,GridView1.SelectedRow] + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;



    if chktotal.Checked then
    begin
        SQLStr := 'UPDATE tblStockOrdenes SET Stock = 1 WHERE ITE_Nombre = ' +
        QuotedStr(gsYear + '-' + GridView1.Cells[0,GridView1.SelectedRow]);

        Conn.Execute(SQLStr);

        CancelClick(nil);
        BindGrid();
    end
    else
    begin
        if StrToInt(txtCantidad.Text) > StrToInt(GridView1.Cells[3,GridView1.SelectedRow]) then
        begin
                ShowMessage('La cantidad no puede ser mayor que la de la orden');
                Exit;
        end;

        if StrToInt(txtCantidad.Text) <= 0 then
        begin
                ShowMessage('La cantidad no puede ser menor o igual a cero.');
                Exit;
        end;
        sOrden := gsYear + '-' + GridView1.Cells[0,GridView1.SelectedRow];
        iCant := StrToInt(GridView1.Cells[3,GridView1.SelectedRow]) - StrToInt(txtCantidad.Text);

        Qry2 := TADOQuery.Create(nil);
        Qry2.Connection :=Conn;

        Qry3 := TADOQuery.Create(nil);
        Qry3.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblStockOrdenes WHERE ITE_Nombre = ' + QuotedStr(sOrden);

        Qry2.SQL.Clear;
        Qry2.SQL.Text := SQLStr;
        Qry2.Open;

        SQLStr := 'Insert_Orden ' + QuotedStr(sOrden) + ',' + QuotedStr(VarToStr(Qry2['TipoProceso'])) +
                  ',' + IntToStr(iCant) + ',' + IntToStr(iCant) + ',' + QuotedStr(VarToStr(Qry2['Producto'])) +
                  ',' + QuotedStr(VarToStr(Qry2['Numero'])) + ',' + QuotedStr(VarToStr(Qry2['Terminal'])) +
                  ',' + QuotedStr(VarToStr(Qry2['Entrega'])) + ',' + QuotedStr(VarToStr(Qry2['Recibido'])) +
                  ',' + QuotedStr(VarToStr(Qry2['Interna'])) +
                  ',' + QuotedStr(VarToStr(Qry2['Nombre'])) + ',' + VarToStr(Qry2['Aprobacion']) +
                  ',' + QuotedStr(VarToStr(Qry2['Observaciones'])) + ',' + QuotedStr(VarToStr(Qry2['Otras'])) +
                  ',' + VarToStr(Qry2['Unitario']) + ',' + VarToStr(Qry2['Total']) + ',' + QuotedStr('Ventas') +
                  ',' + QuotedStr(frmMain.sUserLogin) + ',' + QuotedStr(GetLocalIP) +
                  ',' + QuotedStr(VarToStr(Qry2['OrdenCompra'])) + ',' + VarToStr(Qry2['Dolares']);

        Qry3.SQL.Clear;
        Qry3.SQL.Text := SQLStr;
        Qry3.Open;

        if VarToStr(Qry3['ERROR']) = '-1' Then
          begin
                ShowMessage(VarToStr(Qry3['MSG']));
                Exit;
          end;

        Qry3.Close;


        SQLStr := 'UPDATE tblStockOrdenes SET Stock = 1, Cantidad = ' + GridView1.Cells[3,GridView1.SelectedRow] +
                  ', Requerida = ' + txtCantidad.Text + ', Ordenada = ' + txtCantidad.Text +
                  ' WHERE ITE_Nombre = ' +
        QuotedStr(sOrden);

        Conn.Execute(SQLStr);


        if Qry2.RecordCount <= 0 then
                    Exit;


        CancelClick(nil);
        BindGrid();

        if MessageDlg('Quieres imprimer la orden que se genero de la cantidad faltanta?',
                     mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
            Qry2.Close;
            Exit;
        end;

        Application.Initialize;
        Application.CreateForm(TqrImpresionOrden,qrImpresionOrden);

        qrImpresionOrden.QROrden.Caption := VarToStr(Qry2['ITE_Nombre']);
        qrImpresionOrden.QROrden2.Caption := VarToStr(Qry2['ITE_Nombre']);

        qrImpresionOrden.QRCode1.Caption := '*' + VarToStr(Qry2['ITE_Nombre']) + '*';
        qrImpresionOrden.QRCode2.Caption := '*' + VarToStr(Qry2['ITE_Nombre']) + '*';
        qrImpresionOrden.QRSemana.Caption := '';
        qrImpresionOrden.QRNumero.Caption := VarToStr(Qry2['Numero']);
        qrImpresionOrden.QRTerminal.Caption := VarToStr(Qry2['Terminal']);
        qrImpresionOrden.QRRecibido.Caption := VarToStr(Qry2['Recibido']);
        qrImpresionOrden.QREntrega.Caption := VarToStr(Qry2['Entrega']);
        qrImpresionOrden.QREntrega2.Caption := VarToStr(Qry2['Entrega']);
        qrImpresionOrden.QRNombre.Caption := VarToStr(Qry2['Nombre']);
        qrImpresionOrden.QRFirma.Caption := '';
        qrImpresionOrden.QRObs.Caption := VarToStr(Qry2['Observaciones']);
        qrImpresionOrden.QRDesc.Caption := VarToStr(Qry2['Producto']);
        qrImpresionOrden.QRCompra.Caption := VarToStr(Qry2['OrdenCompra']);
        qrImpresionOrden.QRProceso.Caption := VarToStr(Qry2['TipoProceso']);
        qrImpresionOrden.QRCantidad.Caption := IntToStr(iCant);

        qrImpresionOrden.QRMsg.Caption := 'Forma: Larco-015' + #13 +
                                          'Nivel de Revisi�n: C' + #13 +
                                          'Retenci�n: 1 a�o+uso';
        //qrImpresionOrden.Print;
        qrImpresionOrden.Preview;
        qrImpresionOrden.Free;

        Qry2.Close;
    end;
}
end;

procedure TfrmRelacionOC.chkTotalClick(Sender: TObject);
begin
txtCantidad.Enabled := Not chkTotal.Checked;
end;

procedure TfrmRelacionOC.EditarClick(Sender: TObject);
//var SQLStr,sOrden : String;
//Qry2 : TADOQuery;
begin
//Comente este codigo por que no lo estan usando era para la importacion de excel
{if GridView1.Cells[0,GridView1.SelectedRow] = '' then
    exit;


sOrden := gsYear + '-' + GridView1.Cells[0,GridView1.SelectedRow];

Qry2 := TADOQuery.Create(nil);
Qry2.Connection :=Conn;

SQLStr := 'SELECT * FROM tblStockOrdenes WHERE ITE_Nombre = ' + QuotedStr(sOrden);

Qry2.SQL.Clear;
Qry2.SQL.Text := SQLStr;
Qry2.Open;

if Qry2.RecordCount <= 0 then
            Exit;

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

lblAct.Caption := '1';
frmVentas.NuevoClick(nil);
frmVentas.lblId.Caption := VarToStr(Qry2['ITE_Id']);
frmVentas.txtOrden.Text := GridView1.Cells[0,GridView1.SelectedRow];
frmVentas.txtProceso.Text := VarToStr(Qry2['TipoProceso']);
frmVentas.txtRequerida.Text := VarToStr(Qry2['Requerida']);
frmVentas.txtOrdenada.Text := VarToStr(Qry2['Ordenada']);
frmVentas.txtNumero.Text := VarToStr(Qry2['Numero']);
frmVentas.txtTerminal.Text := VarToStr(Qry2['Terminal']);
frmVentas.deEntrega.Text := VarToStr(Qry2['Entrega']);
frmVentas.deInterna.Text := VarToStr(Qry2['Interna']);
frmVentas.txtRecibido.Text := VarToStr(Qry2['Recibido']);
frmVentas.txtUnitario.Text := VarToStr(Qry2['Unitario']);
frmVentas.txtObservaciones.Text := VarToStr(Qry2['Observaciones']);
frmVentas.txtOtras.Text := VarToStr(Qry2['Otras']);
frmVentas.txtTotal.Text := VarToStr(Qry2['Total']);
frmVentas.cmbEmpleados.Text := VarToStr(Qry2['Nombre']);
frmVentas.cmbProductos.Text := VarToStr(Qry2['Producto']);
frmVentas.chkAprobacion.Checked := StrToBool(VarToStr(Qry2['Aprobacion']));
frmVentas.chkDlls.Checked := StrToBool(VarToStr(Qry2['Dolares']));
frmVentas.lblStock.Caption := '1';

Qry2.Close;
frmVentas.txtOrden.SetFocus;
}
end;

function TfrmRelacionOC.FormIsRunning(FormName: String):Boolean;
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


procedure TfrmRelacionOC.FormActivate(Sender: TObject);
begin

    if lblAct.Caption = '1' then
    begin
            BindGrid();
            lblAct.Caption := '0';
    end;

end;

procedure TfrmRelacionOC.Button2Click(Sender: TObject);
begin
    Application.Initialize;
    Application.CreateForm(TqrRelacionOC, qrRelacionOC);
    qrRelacionOC.ReportTitle.Caption := 'Relacion de Orden de Compra por Cliente';

    qrRelacionOC.QRSubDetail1.DataSet := Qry;
    qrRelacionOC.Field1.DataSet := Qry;
    qrRelacionOC.Field1.DataField := 'Orden';

    qrRelacionOC.Field2.DataSet := Qry;
    qrRelacionOC.Field2.DataField := 'Recibido';

    qrRelacionOC.Field3.DataSet := Qry;
    qrRelacionOC.Field3.DataField := 'OrdenCompra';

    qrRelacionOC.Field4.DataSet := Qry;
    qrRelacionOC.Field4.DataField := 'Cantidad';

    qrRelacionOC.Field5.DataSet := Qry;
    qrRelacionOC.Field5.DataField := 'Descripcion';

    qrRelacionOC.Field6.DataSet := Qry;
    qrRelacionOC.Field6.DataField := 'Numero';

    qrRelacionOC.Field7.DataSet := Qry;
    qrRelacionOC.Field7.DataField := 'Entrega';

    qrRelacionOC.Field8.DataSet := Qry;
    qrRelacionOC.Field8.DataField := 'Tarea';

    qrRelacionOC.Field9.DataSet := Qry;
    qrRelacionOC.Field9.DataField := 'Status';

    qrRelacionOC.Preview;
    qrRelacionOC.Free;

end;

procedure TfrmRelacionOC.IdleLoop;
var
  Start: Integer;
begin
  //Override this method to process some activities when there's no 'Ready' Items for your Client
  Start := 1;
  while (Start < 2) do
    begin
      Start := Start + 1;
      Application.ProcessMessages;
    end;
end;


procedure TfrmRelacionOC.Button3Click(Sender: TObject);
begin
  if GroupBox5.Visible = True then
  begin
          GroupBox5.Visible := False;
  end
  else begin
      GroupBox5.Visible := True;
      if txtCliente.Text = 'Todos' then
      begin
              CheckBox1.Checked := True;
              gvClientes.Enabled := False;
              btnTodos.Enabled := False;
      end
      else
      begin
              CheckBox1.Checked := False;
              gvClientes.Enabled := True;
              btnTodos.Enabled := True;
      end;

      GroupBox5.Top := txtCliente.Top + txtCliente.Height + 5;
      GroupBox5.Left := txtCliente.Left + 10;
  end;
end;

procedure TfrmRelacionOC.CheckBox1Click(Sender: TObject);
begin
gvClientes.Enabled := not CheckBox1.Checked;
btnTodos.Enabled := not CheckBox1.Checked;
end;

procedure TfrmRelacionOC.btnOKClick(Sender: TObject);
var i: integer;
sClientes : String;
begin
  GroupBox5.Visible := False;
  if CheckBox1.Checked = True then begin
          txtCliente.Text := 'Todos';
  end
  else begin
        sClientes := '';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                if gvClientes.Cell[1,i].AsBoolean = True then
                begin
                        sClientes := sClientes + gvClientes.Cells[0,i] + ',';
                end;
        end;
        txtCliente.Text := 'Todos';
        if sClientes <> '' then
        begin
                txtCliente.Text :=  LeftStr(sClientes,Length(sClientes) - 1);
        end;
  end;
end;

procedure TfrmRelacionOC.btnTodosClick(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos.Caption) = UT('Seleccionar Todos') then begin
        btnTodos.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos.Caption := 'Seleccionar Todos';
        for i:= 0 to gvClientes.RowCount - 1 do
        begin
                gvClientes.Cell[1,i].AsBoolean := False;
        end;
  end;

end;

procedure TfrmRelacionOC.Button7Click(Sender: TObject);
begin
  if GroupBox3.Visible = True then
  begin
          GroupBox3.Visible := False;
  end
  else begin
      GroupBox3.Visible := True;
      if txtOrden.Text = 'Todos' then
      begin
              CheckBox2.Checked := True;
              gvOrdenes.Enabled := False;
              btnTodos2.Enabled := False;
      end
      else
      begin
              CheckBox2.Checked := False;
              gvOrdenes.Enabled := True;
              btnTodos2.Enabled := True;
      end;

      GroupBox3.Top := txtOrden.Top + txtOrden.Height + 5;
      GroupBox3.Left := txtOrden.Left + 10;
  end;

end;

procedure TfrmRelacionOC.CheckBox2Click(Sender: TObject);
begin
gvOrdenes.Enabled := not CheckBox2.Checked;
btnTodos2.Enabled := not CheckBox2.Checked;
end;

procedure TfrmRelacionOC.btnOK2Click(Sender: TObject);
var i: integer;
sOrdenes : String;
begin
  GroupBox3.Visible := False;
  if CheckBox2.Checked = True then begin
          txtOrden.Text := 'Todos';
  end
  else begin
        sOrdenes := '';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                if gvOrdenes.Cell[1,i].AsBoolean = True then
                begin
                        sOrdenes := sOrdenes + gvOrdenes.Cells[0,i] + ',';
                end;
        end;
        txtOrden.Text := 'Todos';
        if sOrdenes <> '' then
        begin
                txtOrden.Text :=  LeftStr(sOrdenes,Length(sOrdenes) - 1);
        end;
  end;
end;

procedure TfrmRelacionOC.btnTodos2Click(Sender: TObject);
var i: integer;
begin
  if UT(btnTodos2.Caption) = UT('Seleccionar Todos') then begin
        btnTodos2.Caption := 'Deseleccionar Todos';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                gvOrdenes.Cell[1,i].AsBoolean := True;
        end;
  end
  else begin
        btnTodos2.Caption := 'Seleccionar Todos';
        for i:= 0 to gvOrdenes.RowCount - 1 do
        begin
                gvOrdenes.Cell[1,i].AsBoolean := False;
        end;
  end;

end;

procedure TfrmRelacionOC.chkRecibidoClick(Sender: TObject);
begin
deRecibido1.Enabled := chkRecibido.Checked;
deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmRelacionOC.cmbPartesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      BindGrid();
  end
end;

procedure TfrmRelacionOC.cmbCompraKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If Key = vk_return then
  begin
      BindGrid();
  end
end;

procedure TfrmRelacionOC.ExportROCClick(Sender: TObject);
var sFileName: String;
begin
  if GridView1.RowCount = 0 then
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

    ExportGridROC(GridView1,sFileName);

  end;

end;

procedure TfrmRelacionOC.ExportGrid(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
  xlCSV = 6;
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
      Sheet.Name := 'Ordenes de Compra';

      for Col := 1 to Grid.Columns.Count do
              Sheet.Cells[1,Col] := Grid.Columns[Col - 1].Header.Caption;

      for Row := 1 to Grid.RowCount do
                for Col := 1 to Grid.Columns.Count do
                        Sheet.Cells[Row + 1,Col] := Grid.Cells[Col - 1,Row - 1];



      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

       showmessage('El archivo se exporto exitosamente.');
end;

procedure TfrmRelacionOC.ExportGridROC(Grid: TGridView;sFileName: String);
const
  xlWorkSheet = -4167;
  xlCSV = 6;
var XApp : Variant;
Sheet : Variant;
Row :Integer;
StartDDir : String;
begin
      try //Create the excel object
      begin
          XApp:= CreateOleObject('Excel.Application');
          XApp.Workbooks.Add(xlWorkSheet);
          StartDDir := ExtractFileDir(ParamStr(0)) + '\';
          XApp.WorkBooks.Open(StartDDir + 'RelacionOrdenCompra.xls');
          XApp.Visible := False;
          XApp.DisplayAlerts := False;
      end;
      except
          showmessage('No se pudo abrir Microsoft Excel,  parece que no esta instalado en el sistema.');
          exit;
      end;


      //Sheet := XApp.Workbooks[1].WorkSheets[1];
      //Sheet := XApp.ActiveWorkbook.ActiveSheet;
      Sheet := XApp.ActiveWorkbook.WorkSheets['PAGE 1'];

      //Sheet.Name := 'Ordenes de Compra';
      Sheet.Cells[5, 3] := UpperCase(cmbCompra.Text);
      Sheet.Cells[4, 6] := UpperCase(getClient(LeftStr(Grid.Cells[0, 1], 3), Grid.Cells[1, 1]));
      Sheet.Cells[5, 6] := LeftStr(Grid.Cells[0, 1], 3);
      Sheet.Cells[5, 8] := 'SEM ' + lblWeek.Caption;


      for Row := 1 to Grid.RowCount do begin
        Sheet.Cells[Row + 7, 1] := Grid.Cells[0, Row - 1];
        Sheet.Cells[Row + 7, 2] := Grid.Cells[1, Row - 1];
        Sheet.Cells[Row + 7, 3] := UpperCase(Grid.Cells[20, Row - 1]);
        Sheet.Cells[Row + 7, 4] := Grid.Cells[4, Row - 1];
        Sheet.Cells[Row + 7, 5] := UpperCase(Grid.Cells[5, Row - 1]);
        Sheet.Cells[Row + 7, 6] := UpperCase(Grid.Cells[7, Row - 1]);
        Sheet.Cells[Row + 7, 7] := Grid.Cells[19, Row - 1];
        Sheet.Cells[Row + 7, 8] := Grid.Cells[8, Row - 1];
      end;


      XApp.ActiveWorkBook.SaveAs(sFileName);
      Sheet := Unassigned;
      XApp.Quit;
      XApp := Unassigned;

      ShowMessage('El archivo se exporto exitosamente.');
end;

function TfrmRelacionOC.getBitVal(value : String):String;
begin
  result := 'No';
  if(value = '1') then result := 'Si';
end;

procedure TfrmRelacionOC.cmbCompraKeyPress(Sender: TObject; var Key: Char);
begin
  Key := upcase(Key);
end;

procedure TfrmRelacionOC.cmbPartesKeyPress(Sender: TObject; var Key: Char);
begin
  Key := upcase(Key);
end;

function TfrmRelacionOC.getClient(clientId : String; date: String) : String;
var SQLStr: String;
Qry2 : TADOQuery;
begin
  result := '';
  lblWeek.Caption := '';
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  SQLStr := 'SELECT *, dbo.ISOWeek(' + QuotedStr(date) + ') As Week FROM tblClientes WHERE Clave = ' + QuotedStr(clientId);

  Qry2.SQL.Clear;
  Qry2.SQL.Text := SQLStr;
  Qry2.Open;

  if Qry2.RecordCount > 0 then begin
    result := VarToStr(Qry2['Nombre']);
    lblWeek.Caption := VarToStr(Qry2['Week']);
  end;

  Qry2.Close;
  Qry2.Free;
end;

end.

