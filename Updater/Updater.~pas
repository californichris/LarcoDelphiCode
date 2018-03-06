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
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
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
Qry2 : TADOQuery;
begin
    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;

    Qry2 := TADOQuery.Create(nil);
    Qry2.Connection :=Conn;

    Qry2.SQL.Clear;
    Qry2.SQL.Text := Memo1.Text;
    Qry2.Open;


    if Qry2.RecordCount > 0 then
    begin
        Memo1.Text := '';
        while not Qry2.Eof do
          begin
            Memo1.Lines.Add(VarToStr(Qry2['ITE_Nombre']));

            Qry2.Next;
          end;

    end
    else begin
         Memo1.Text := 'No records found';
    end;

    Qry2.Close;


end;

procedure TForm1.Button2Click(Sender: TObject);
var Strings : TStringList;
begin

end;

end.
