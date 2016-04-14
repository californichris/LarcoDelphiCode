unit PrintLabel;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls, jpeg;

type
  TLabelReport = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    lblCliente: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRImage1: TQRImage;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRLabel8: TQRLabel;
    lblFecha: TQRLabel;
    lblOCompra: TQRLabel;
    lblDesc: TQRLabel;
    lblCantidad: TQRLabel;
    lblReq: TQRLabel;
    lblNoParte: TQRLabel;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    lblPartida: TQRLabel;
  private

  public

  end;

var
  LabelReport: TLabelReport;

implementation

{$R *.DFM}

end.
