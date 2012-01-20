unit DetalleOrdenes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ScrollView, CustomGridViewControl, CustomGridView,
  GridView;

type
  TfrmDetalle = class(TForm)
    GridView1: TGridView;
    btnCerrar: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCerrarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDetalle: TfrmDetalle;

implementation

{$R *.dfm}

procedure TfrmDetalle.FormClose(Sender: TObject; var Action: TCloseAction);
begin
        Action := caFree;
end;

procedure TfrmDetalle.btnCerrarClick(Sender: TObject);
begin
        Self.Close;
end;

end.
