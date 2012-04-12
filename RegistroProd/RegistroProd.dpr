program RegistroProd;

uses
  Forms,
  RegistroProduccion in 'RegistroProduccion.pas' {frmRegistroProd},
  Login in 'Login.pas' {frmLogin};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmRegistroProd, frmRegistroProd);
  Application.CreateForm(TfrmLogin, frmLogin);
  Application.Run;
end.
