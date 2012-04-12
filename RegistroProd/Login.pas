unit Login;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls;

type
  TfrmLogin = class(TForm)
    cmdOk: TButton;
    cmdCancel: TButton;
    Panel1: TPanel;
    Label2: TLabel;
    txtPassword: TEdit;
    txtUser: TEdit;
    Label1: TLabel;
    procedure cmdCancelClick(Sender: TObject);
    procedure cmdOkClick(Sender: TObject);
    function IsValidUser(User: String;Password: String):boolean;
    procedure txtUserKeyPress(Sender: TObject; var Key: Char);
    procedure txtPasswordKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }

  end;

var
  frmLogin: TfrmLogin;
  gsConnString,StartDDir: String;


implementation

uses RegistroProduccion;

{$R *.dfm}

procedure TfrmLogin.cmdCancelClick(Sender: TObject);
begin
ModalResult := mrCancel;
Close;
end;

procedure TfrmLogin.cmdOkClick(Sender: TObject);
begin
   if  IsValidUser(txtUser.Text,txtPassword.Text ) then
       begin
            frmRegistroProd.lblEmpleado.Caption := txtUser.Text;
            ModalResult := mrOK;
       end
   else
       Begin
            ShowMessage('No tienes permisos para Autorizar.');
            txtPassword.Text := '';
            txtPassword.SetFocus;
       end;
end;

function TfrmLogin.IsValidUser(User: String;Password: String):boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := False;
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;
    Qry := TADOQuery.Create(nil);
    Qry.Connection :=Conn;

    SQLStr := 'SELECT TOP 1 U.USE_ID FROM tblusers U ' +
              'INNER JOIN tblUser_groups UG ON U.USE_ID = UG.USE_ID ' +
              'INNER JOIN tblGroups G ON UG.Group_ID = G.Group_ID ' +
              'WHERE UPPER(Group_Name) = UPPER(' + QuotedStr('TornosCNC') + ') ' +
              'AND U.USE_Login = ' + QuotedStr(User) + ' and U.USE_Password = ' + QuotedStr(Password);

    Qry.SQL.Clear;
    Qry.SQL.Text := SQLStr;
    Qry.Open;

    if Qry.RecordCount > 0 then begin
        Result := True;
    end;

    Qry.Close;
    Conn.Close;
end;

procedure TfrmLogin.txtUserKeyPress(Sender: TObject; var Key: Char);
begin
     if (key = chr(vk_return)) or (key = chr(vk_tab)) then
        txtPassword.SetFocus;
end;

procedure TfrmLogin.txtPasswordKeyPress(Sender: TObject; var Key: Char);
begin
     if (key = chr(vk_return)) or (key = chr(vk_tab)) then
        cmdOkClick(nil);
end;

procedure TfrmLogin.FormShow(Sender: TObject);
begin
frmRegistroProd.Enabled := False;
end;

procedure TfrmLogin.FormCreate(Sender: TObject);
var sUser,sPassword,sServer,sDB : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','BeltranC');
    sDB := IniFile.ReadString('Conn','DB','Larco');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','');
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;


end;

end.
