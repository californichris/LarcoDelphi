unit ReporteCumplimientoQr;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls, TeEngine, Series, TeeProcs,
  Chart, DbChart, QRTEE;

type
  TqrCumpliGrafica = class(TQuickRep)
    PageHeaderBand1: TQRBand;
    ReportTitle: TQRLabel;
    DetailBand1: TQRBand;
    lblLiberadas: TQRLabel;
    lblScrap: TQRLabel;
    lblTotal: TQRLabel;
    lblPorcentaje: TQRLabel;
    QRChart1: TQRChart;
    QRDBChart1: TQRDBChart;
    Series1: TPieSeries;
  private

  public

  end;

var
  qrCumpliGrafica: TqrCumpliGrafica;

implementation

{$R *.DFM}

end.
