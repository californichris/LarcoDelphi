unit ReporteRelacionOC;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt;

type
  TfrmOCReport = class(TForm)
    QuickRep1: TQuickRep;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmOCReport: TfrmOCReport;

implementation

{$R *.dfm}

procedure TfrmOCReport.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

end.
