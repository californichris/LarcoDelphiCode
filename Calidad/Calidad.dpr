program Calidad;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  Imprimir in 'Imprimir.pas' {frmImprimir},
  Unit3 in 'Unit3.pas' {PrintReport: TQuickRep},
  Scrap in 'Scrap.pas' {frmScrap},
  Login in 'Login.pas' {frmLogin};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Larco Quality App';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
