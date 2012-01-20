unit ImpresionOrden;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, Dialogs,QRCtrls,IniFiles, QRExport;

type
  TqrImpresionOrden = class(TQuickRep)
    QRCSVFilter1: TQRCSVFilter;
    QRTextFilter1: TQRTextFilter;
    DetailBand1: TQRBand;
    QRImage1: TQRImage;
    QRDesc: TQRLabel;
    QRRecibido: TQRLabel;
    QRNombre: TQRLabel;
    QRObs: TQRLabel;
    QRNumero: TQRLabel;
    QREntrega: TQRLabel;
    QRTerminal: TQRLabel;
    QRFirma: TQRLabel;
    QROrden: TQRLabel;
    QRSemana: TQRLabel;
    QRCode1: TQRLabel;
    QRCode2: TQRLabel;
    QRCompra: TQRLabel;
    QRProceso: TQRLabel;
    QRCantidad: TQRLabel;
    QROrden2: TQRLabel;
    QREntrega2: TQRLabel;
    QRMsg: TQRLabel;
    QRLabel1: TQRLabel;
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private

  public

  end;

var
  qrImpresionOrden: TqrImpresionOrden;

implementation

{$R *.DFM}

procedure TqrImpresionOrden.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var OPicture : TPicture;
StartDDir : String;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';

    OPicture := TPicture.Create;
    OPicture.LoadFromFile(StartDDir + 'Orden.bmp');
    QRImage1.Picture := OPicture;
end;

end.
