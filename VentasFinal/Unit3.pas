unit Unit3;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TPrintReport = class(TQuickRep)
    QRSubDetail1: TQRSubDetail;
    Field1: TQRDBText;
    Field2: TQRDBText;
    Field3: TQRDBText;
    Field4: TQRDBText;
    Field5: TQRDBText;
    Field6: TQRDBText;
    PageFooterBand1: TQRBand;
    QRExpr1: TQRExpr;
    QRExpr2: TQRExpr;
    TitleBand1: TQRBand;
    ReportTitle: TQRLabel;
    THeader1: TQRLabel;
    THeader2: TQRLabel;
    THeader3: TQRLabel;
    THeader4: TQRLabel;
    THeader5: TQRLabel;
    THeader6: TQRLabel;
    lblType: TQRLabel;
    ColumnHeaderBand1: TQRBand;
    Header1: TQRLabel;
    Header2: TQRLabel;
    Header3: TQRLabel;
    Header4: TQRLabel;
    Header5: TQRLabel;
    Header6: TQRLabel;
  private

  public

  end;

var
  PrintReport: TPrintReport;

implementation

{$R *.DFM}

end.
