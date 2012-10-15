program MultipleTask;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  Imprimir in 'Imprimir.pas' {frmImprimir},
  Unit3 in 'Unit3.pas' {PrintReport: TQuickRep};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Larco MultiTask App';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
