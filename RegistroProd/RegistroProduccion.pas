unit RegistroProduccion;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions,StrUtils,chris_Functions, Mask, StdCtrls,sndkey32,
  ScrollView, CustomGridViewControl, CustomGridView, GridView, ComCtrls,ComObj,
  CellEditors, ExtCtrls, IdTrivialFTPBase, DateUtils;

type
  TfrmRegistroProd = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    txtLote: TEdit;
    txtMaquina: TEdit;
    txtPiezas: TEdit;
    txtScrap: TEdit;
    txtEmpleado: TEdit;
    txtHoras: TEdit;
    txtOtrasHoras: TEdit;
    txtActividad: TMemo;
    Label9: TLabel;
    GroupBox2: TGroupBox;
    Label10: TLabel;
    gvPendientes: TGridView;
    btnAutorizar: TButton;
    Nuevo: TButton;
    Editar: TButton;
    Borrar: TButton;
    Buscar: TButton;
    Aceptar: TButton;
    Cancelar: TButton;
    btnPendientes: TButton;
    lblEmpleado: TLabel;
    Timer1: TTimer;
    txtOrden: TMaskEdit;
    lblAnio: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ClearForm();
    procedure btnPendientesClick(Sender: TObject);
    procedure NuevoClick(Sender: TObject);
    procedure txtPiezasKeyPress(Sender: TObject; var Key: Char);
    procedure txtHorasKeyPress(Sender: TObject; var Key: Char);
    procedure CancelarClick(Sender: TObject);
    procedure txtOrdenKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DisableForm(value: Boolean);
    procedure AceptarClick(Sender: TObject);
    function getFormYear(sConnString: String; sFormName: String): String;
    function isValidData(orden: String): boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRegistroProd: TfrmRegistroProd;
  gsConnString, gsYear,gsOYear: String;
  giIntervalo: Integer;
  Conn : TADOConnection;
  giOpcion : Integer;

implementation

{$R *.dfm}

procedure TfrmRegistroProd.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frmRegistroProd.Timer1Timer(nil);
  setActiveWindow(frmRegistroProd.Handle);
  Action := caFree;
end;

procedure TfrmRegistroProd.FormCreate(Sender: TObject);
var sUser,sPassword,sServer,sDB,StartDDir : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','BeltranC');
    sDB := IniFile.ReadString('Conn','DB','Larco');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','');
    giIntervalo := StrToInt(IniFile.ReadString('Tasks','Refresh','30000'));
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;

    //Create Connection
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;


    Timer1.Interval := giIntervalo;
    lblAnio.Caption := getFormYear(gsConnString, Self.Name);
    gsOYear := RightStr(lblAnio.Caption,2);
    gsYear := gsOYear + '-';

end;

procedure TfrmRegistroProd.Timer1Timer(Sender: TObject);
var Qry : TADOQuery;
SQLStr : String;
begin
  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  SQLStr := 'SELECT * FROM tblLotes WHERE Autorizado = 0 ORDER BY ITE_Nombre, Lote';


  Qry.SQL.Clear;
  Qry.SQL.Text := SQLStr;
  Qry.Open;

  gvPendientes.ClearRows();
  While not Qry.Eof do
  Begin
      gvPendientes.AddRow(1);
      gvPendientes.Cells[0, gvPendientes.RowCount -1] := VarToStr(Qry['Lote_Id']);
      gvPendientes.Cells[1, gvPendientes.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
      gvPendientes.Cells[2, gvPendientes.RowCount -1] := VarToStr(Qry['Lote']);
      Qry.Next;
  End;

  Qry.Close;

end;

procedure TfrmRegistroProd.ClearForm();
begin
  txtOrden.Text := '';
  txtLote.Text := '';
  txtMaquina.Text := '';
  txtPiezas.Text := '';
  txtScrap.Text := '';
  txtEmpleado.Text := '';
  txtHoras.Text := '';
  txtOtrasHoras.Text := '';
  txtActividad.Text := '';
end;

procedure TfrmRegistroProd.DisableForm(value: Boolean);
begin
  txtOrden.Enabled := not value;
  txtLote.Enabled := not value;
  txtMaquina.Enabled := not value;
  txtPiezas.Enabled := not value;
  txtScrap.Enabled := not value;
  txtEmpleado.Enabled := not value;
  txtHoras.Enabled := not value;
  txtOtrasHoras.Enabled := not value;
  txtActividad.Enabled := not value;
end;

procedure TfrmRegistroProd.btnPendientesClick(Sender: TObject);
begin
  if(btnPendientes.Caption = '<')then begin
    frmRegistroProd.Width := 521;
    btnPendientes.Caption := '>';
    btnPendientes.Hint := 'Mostrar Pendientes de Autorizar';
  end
  else begin
    frmRegistroProd.Width := 847;
    btnPendientes.Caption := '<';
    btnPendientes.Hint := 'Esconder Pendientes de Autorizar';
  end;
end;

procedure TfrmRegistroProd.NuevoClick(Sender: TObject);
begin
  giOpcion := 1;
  Nuevo.Enabled := False;
  Aceptar.Enabled := True;
  Cancelar.Enabled := True;
  ClearForm();
  txtOrden.SetFocus;
end;

procedure TfrmRegistroProd.txtPiezasKeyPress(Sender: TObject;
  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
       else
                Key := #0;
end;

procedure TfrmRegistroProd.txtHorasKeyPress(Sender: TObject;
  var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key in ['.']) then
            begin
                if StrPos(PChar((Sender as TEdit).Text), '.') <> nil then
                  Key := #0;
            end
       else
                Key := #0;
end;

procedure TfrmRegistroProd.CancelarClick(Sender: TObject);
begin
  ClearForm();
end;

procedure TfrmRegistroProd.txtOrdenKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   If Key = vk_return then
   begin
        AppActivate(Application.Handle);
        SendKeys('{TAB}',False);
   end
   else if (Key = vk_Escape) and (Cancelar.Enabled = True)  then
    begin
            CancelarClick(nil);
    end;

end;

procedure TfrmRegistroProd.AceptarClick(Sender: TObject);
var SQLStr, orden : String;
Qry2 : TADOQuery;
begin
  Qry2 := TADOQuery.Create(nil);
  Qry2.Connection :=Conn;

  if giOpcion = 1 then
  begin
        orden := gsYear + txtOrden.Text;
        if not isValidData(orden) then
          Exit;

        SQLStr := 'INSERT INTO tblLotes(ITE_Nombre, Lote, Maquina, Piezas, Scrap, USE_ID, Horas, Otras_Horas, Actividad) ' +
                  'VALUES(' + QuotedStr(orden) + ',' + QuotedStr(txtLote.Text) + ',' + QuotedStr(txtMaquina.Text)
                  + ',' + QuotedStr(txtPiezas.Text) + ',' + QuotedStr(txtScrap.Text) + ',' + QuotedStr(txtEmpleado.Text)
                  + ',' + QuotedStr(txtHoras.Text) + ',' + QuotedStr(txtOtrasHoras.Text) + ',' + QuotedStr(txtActividad.Text) + ')';

        conn.Execute(SQLStr);
        ShowMessage('El registro de produccion fue grabado exitosamente.');
        ClearForm();
  end
  else if giOpcion = 2 then
  begin

  end
  else if giOpcion = 3 then
  begin

  end
  else if giOpcion = 4 then
  begin

          //Exit;
  end;

  giOpcion := 0;
  Timer1Timer(nil);
end;

function TfrmRegistroProd.getFormYear(sConnString: String; sFormName: String): String;
var SQLStr : String;
Qry : TADOQuery;
year: Word;
begin
    Qry := nil;
    Result := '';
    try
        Qry := TADOQuery.Create(nil);
        Qry.Connection :=Conn;

        SQLStr := 'SELECT * FROM tblScreens WHERE SCR_FormName = ' + QuotedStr(sFormName);

        Qry.SQL.Clear;
        Qry.SQL.Text := SQLStr;
        Qry.Open;

        if Qry.RecordCount > 0 then begin
                Result := VarToStr(Qry['SCR_Year']);
        end
        else begin
                Result := VarToStr(YearOf(Date));
        end;

    except
          on e : EOleException do
                ShowMessage('La base de datos no esta disponible. Por favor verifique que exista conectividad al servidor.');
          on e : Exception do
                ShowMessage(e.ClassName + ' error raised, with message : ' + e.Message + ' Method : getFormYear');
    end;

    Qry.Close;
end;

function TfrmRegistroProd.isValidData(orden: String): boolean;
var SQLStr : String;
Qry : TADOQuery;
begin
  result:= True;

  SQLStr := 'Select ITE_Nombre FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(orden);
  Qry := TADOQuery.Create(nil);
  Qry.Connection :=Conn;

  Qry.SQL.Clear;
  Qry.SQL.Text := SQLStr;
  Qry.Open;

  If Qry.RecordCount <= 0 Then begin
    result:= False;
    ShowMessage('La orden no existe en el sistema');
  end;

  SQLStr := 'SELECT Lote_Id FROM tblLotes WHERE Lote = ' + QuotedStr(txtLote.Text) + ' AND Maquina = ' + QuotedStr(txtMaquina.Text);

  Qry.SQL.Clear;
  Qry.SQL.Text := SQLStr;
  Qry.Open;

  If Qry.RecordCount > 0 Then begin
    result:= False;
    ShowMessage('El lote ya esta existe para este Lote y maquina.');
  end;

end;

end.
