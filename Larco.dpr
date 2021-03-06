program Larco;

uses
  Forms,
  Monitor in 'Monitor.pas' {Form1},
  Productos in 'Productos.pas' {frmProductos},
  Main in 'Main.pas' {frmMain},
  Grupos in 'Grupos.pas' {frmGrupos},
  Tareas in 'Tareas.pas' {frmTareas},
  Empleados in 'Empleados.pas' {frmEmpleados},
  Routing in 'Routing.pas' {frmRouting},
  Rutas in 'Rutas.pas' {frmRutas},
  Ventas in 'Ventas.pas' {frmVentas},
  Clientes in 'Clientes.pas' {frmClientes},
  Year in 'Year.pas' {frmYear},
  FormaAtrasos in 'FormaAtrasos.pas' {frmAtrasos},
  FechaEntrega in 'FechaEntrega.pas' {frmEntrega},
  ReporteRelacion in 'ReporteRelacion.pas' {qrRelacionEntrega: TQuickRep},
  Editor in 'Editor.pas' {frmEditor},
  ImpresionOrden in 'ImpresionOrden.pas' {qrImpresionOrden: TQuickRep},
  Login in 'Login.pas' {frmLogin},
  Users in 'Users.pas' {frmUsers},
  Scrap in 'Scrap.pas' {frmScrap},
  ReporteScrap in 'ReporteScrap.pas' {frmScrapReport},
  ReporteScrapQr in 'ReporteScrapQr.pas' {qrReporteScrap: TQuickRep},
  RelacionOrdenCompra in 'RelacionOrdenCompra.pas' {frmRelacionOC},
  ReporteOC in 'ReporteOC.pas' {qrRelacionOC: TQuickRep},
  PorcentajeScrap in 'PorcentajeScrap.pas' {frmScrapPorcen},
  ReporteScrapGraficaQr in 'ReporteScrapGraficaQr.pas' {qrScrapGrafica: TQuickRep},
  CatalogoScreens in 'CatalogoScreens.pas' {frmCatalogoScreens},
  PorcentajeRetrabajo in 'PorcentajeRetrabajo.pas' {frmRetrabajo},
  ReporteRetrabajoGraficaQr in 'ReporteRetrabajoGraficaQr.pas' {qrRetrabajoGrafica: TQuickRep},
  PorcentajeScrapDinero in 'PorcentajeScrapDinero.pas' {frmScrapDinero},
  ReporteDineroScrapGraficaQr in 'ReporteDineroScrapGraficaQr.pas' {qrDineroScrapGrafica: TQuickRep},
  ExchangeRate in 'ExchangeRate.pas' {frmExchangeRate},
  ReporteCumplimiento in 'ReporteCumplimiento.pas' {frmCumplimiento},
  ReporteCumplimientoQr in 'ReporteCumplimientoQr.pas' {qrCumpliGrafica: TQuickRep},
  ReportePromedioCump in 'ReportePromedioCump.pas' {frmPromCumpli},
  PorcentajeRetrabajoDinero in 'PorcentajeRetrabajoDinero.pas' {frmRetrabajoDinero},
  Larco_Functions in 'Larco_Functions.pas',
  ReporteDineroRetrabajoGraficaQr in 'ReporteDineroRetrabajoGraficaQr.pas' {qrDineroRetrabajoGrafica: TQuickRep},
  DetalleOrdenes in 'DetalleOrdenes.pas' {frmDetalle},
  CatalogoCategories in 'CatalogoCategories.pas' {frmCatalogoCategories},
  CatalogoGrupos in 'CatalogoGrupos.pas' {frmCatalogoGrupos},
  MenuEditor in 'MenuEditor.pas' {frmMenuEditor},
  CatalogoPermisos in 'CatalogoPermisos.pas' {frmCatalogoPermisos},
  EditorDeScrap in 'EditorDeScrap.pas' {frmScrapEditor},
  ReporteCargaTrabajo in 'ReporteCargaTrabajo.pas' {frmCargaTrabajo},
  Facturacion in 'Facturacion.pas' {frmFacturacion},
  ReporteFactura in 'ReporteFactura.pas' {qrFactura: TQuickRep},
  PendientesFacturar in 'PendientesFacturar.pas' {frmPendientesFact},
  ReportePendientesFacurar in 'ReportePendientesFacurar.pas' {qrPendientesFacturar: TQuickRep},
  ReporteProductividad in 'ReporteProductividad.pas' {frmProductividad},
  ReporteCargaTrabajoDetalle in 'ReporteCargaTrabajoDetalle.pas' {frmCTDetail},
  CatalogoContribuyente in 'CatalogoContribuyente.pas' {frmContribuyente},
  CatalogoMateriales in 'CatalogoMateriales.pas' {frmMateriales},
  CatalogoPrpductosTerminados in 'CatalogoPrpductosTerminados.pas' {frmProductosTerminados},
  EditorDeRetrabajo in 'EditorDeRetrabajo.pas' {frmEditorRetrabajo},
  CatalogoUnidadMedida in 'CatalogoUnidadMedida.pas' {frmUnidadMedida},
  CatalogoTipoMaterial in 'CatalogoTipoMaterial.pas' {frmTipoMaterial},
  Entradas in 'Entradas.pas' {frmEntradas},
  CatalogoPaises in 'CatalogoPaises.pas' {frmPaises},
  CatalogoProvedores in 'CatalogoProvedores.pas' {frmProvedores},
  SalidasAlmacen in 'SalidasAlmacen.pas' {frmSalidasAlmacen},
  InventariosConfiguracion in 'InventariosConfiguracion.pas' {frmInventariosConf},
  ReporteEntradasSalidasBorradas in 'ReporteEntradasSalidasBorradas.pas' {frmEntradasSalidasBorradas},
  ReportePiezasTerminadas in 'ReportePiezasTerminadas.pas' {frmPiezasTerminadas},
  ReporteEntradasSalidasAlmacen in 'ReporteEntradasSalidasAlmacen.pas' {frmESAlmacen},
  ReporteEntradasSalidasLarco in 'ReporteEntradasSalidasLarco.pas' {frmESLarco},
  ReporteMaterialesEscasos in 'ReporteMaterialesEscasos.pas' {frmEscasos},
  SalidasLarco in 'SalidasLarco.pas' {frmSalidasLarco},
  ReporteProductividadEmpleado in 'ReporteProductividadEmpleado.pas' {frmProdEmpleado},
  ReporteProductividadEmpleadoDinero in 'ReporteProductividadEmpleadoDinero.pas' {frmProdEmpleadoDinero},
  ReporteMaterialesPorOrden in 'ReporteMaterialesPorOrden.pas' {frmMaterialesPorOrden},
  ReporteMaterialesPorOrdenQr in 'ReporteMaterialesPorOrdenQr.pas' {qrReporteMaterialesPorOrden: TQuickRep},
  CatalogoDiasInhabiles in 'CatalogoDiasInhabiles.pas' {frmDiasInhabiles},
  ReporteCumplimientoFechaEntrega in 'ReporteCumplimientoFechaEntrega.pas' {frmCumplimientoTiempoEntrega},
  ReporteCumplimientoFechaEntregaQr in 'ReporteCumplimientoFechaEntregaQr.pas' {qrCumplimientoTiempoEntrega: TQuickRep},
  CatalogoPlanos in 'CatalogoPlanos.pas' {frmCatalogoPlanos},
  EntradasSalidasStock in 'EntradasSalidasStock.pas' {frmESStock},
  ReporteEntradasSalidasStock in 'ReporteEntradasSalidasStock.pas' {frmReporteESStock},
  ReporteEntradasSalidasPlano in 'ReporteEntradasSalidasPlano.pas' {frmReporteESPlano},
  ReporteTotalPiezasStock in 'ReporteTotalPiezasStock.pas' {frmReporteTotalPiezasStock},
  ReportePiezasStock in 'ReportePiezasStock.pas' {frmReportePiezasStock},
  PrintLabel in 'PrintLabel.pas' {LabelReport: TQuickRep};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Larco';
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmLogin, frmLogin);
  Application.Run;
end.
