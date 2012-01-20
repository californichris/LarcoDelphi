unit ReporteScrapGraficaQr;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls, TeEngine, Series, TeeProcs,
  Chart, DbChart, QRTEE;

type
  TqrScrapGrafica = class(TQuickRep)
    PageHeaderBand1: TQRBand;
    DetailBand1: TQRBand;
    ReportTitle: TQRLabel;
    lblLiberadas: TQRLabel;
    lblScrap: TQRLabel;
    lblTotal: TQRLabel;
    lblPorcentaje: TQRLabel;
    QRDBChart1: TQRDBChart;
    QRChart1: TQRChart;
    Series1: TPieSeries;
  private

  public

  end;

var
  qrScrapGrafica: TqrScrapGrafica;

implementation

{$R *.DFM}

end.
