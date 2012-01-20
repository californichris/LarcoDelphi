unit ReporteCargaTrabajoDetalle;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView;

type
  TfrmCTDetail = class(TForm)
    GridView1: TGridView;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCTDetail: TfrmCTDetail;

implementation

{$R *.dfm}

procedure TfrmCTDetail.Button1Click(Sender: TObject);
begin
    Self.Close();
end;

procedure TfrmCTDetail.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
Action := caFree;
end;

end.
