program UpdaterApp;

uses
  Forms,
  Updater in 'Updater.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
