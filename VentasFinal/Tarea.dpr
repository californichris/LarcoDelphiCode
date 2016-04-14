program Tarea;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  Imprimir in 'Imprimir.pas' {frmImprimir},
  Unit3 in 'Unit3.pas' {PrintReport: TQuickRep},
  PrintLabel in 'PrintLabel.pas' {LabelReport: TQuickRep};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Larco Ventas Final App';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
