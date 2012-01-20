unit ReporteCumplimientoFechaEntregaQr;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TqrCumplimientoTiempoEntrega = class(TQuickRep)
    QRSubDetail1: TQRSubDetail;
    lblOrdendeTrabajo: TQRDBText;
    lblFechaEntrada: TQRDBText;
    lblFechaTerminacion: TQRDBText;
    lblDiasAdelantoAtrazo: TQRDBText;
    QRBand1: TQRBand;
    lblOrden: TQRLabel;
    lblEntrada: TQRLabel;
    lblTerminacion: TQRLabel;
    lblDias: TQRLabel;
    lblAdelantoAtrazo: TQRLabel;
    lblTitulo: TQRLabel;
    PageFooterBand1: TQRBand;
    QRExpr1: TQRExpr;
    QRExpr2: TQRExpr;
    lblDiasUtilizados: TQRDBText;
  private

  public

  end;

var
  qrCumplimientoTiempoEntrega: TqrCumplimientoTiempoEntrega;

implementation

{$R *.DFM}

end.
