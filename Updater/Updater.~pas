unit Updater;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, chris_Functions,StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Memo1: TMemo;
    Button2: TButton;
    Edit1: TEdit;
    Memo2: TMemo;
    Memo3: TMemo;
    Memo4: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Edit1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  gsConnString,StartDDir: String;

  
implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
Var sUser,sPassword,sServer,sDB : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','LarcoVentas');
    sDB := IniFile.ReadString('Conn','DB','Larco');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','larco$a');
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;



end;

procedure TForm1.Button1Click(Sender: TObject);
var Conn : TADOConnection;
SQLStr : String;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;

    conn.Execute(Memo1.Text);
    conn.Execute(Memo2.Text);
    conn.Execute(Memo3.Text);
    conn.Execute(Memo4.Text);

    conn.Execute('INSERT INTO tblScreens(SCR_Name, SCR_FormName, SCR_Description, SCR_Year) VALUES(''Reporte de Productividad por Empleado en Dinero'',''frmProdEmpleadoDinero'',''Reporte de Productividad por Empleado en Dinero'',''2008'')');
    label1.Caption := 'La actualizacion se realizo exitosamente.';

    Button1.Enabled := False;
end;

procedure TForm1.Button2Click(Sender: TObject);
var Strings : TStringList;
begin
   Strings := TStringList.Create;
   Strings.Clear;
   Strings.Delimiter := #191;
   Strings.DelimitedText := Edit1.Text;
   ShowMessage(IntToStr(Strings.count));
end;

procedure TForm1.Edit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 Label1.Caption := Chr(Key)+' - '+IntToStr(Key);
end;

end.
