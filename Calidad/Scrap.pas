unit Scrap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,ADODB,DB,IniFiles,All_Functions, chris_Functions,StrUtils;

type
  TfrmScrap = class(TForm)
    GroupBox1: TGroupBox;
    Button1: TButton;
    btnScrap: TButton;
    btnRetrabajo: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    txtMotivo: TEdit;
    cmbTareas: TComboBox;
    cmbEmpleados: TComboBox;
    btnOk: TButton;
    btnCancel: TButton;
    lblTask: TLabel;
    lblEmpleado: TLabel;
    lblOrden: TLabel;
    lblStatus: TLabel;
    lblCantidad: TLabel;
    txtCantidad: TEdit;
    chkParcial: TCheckBox;
    lblRepro: TLabel;
    txtRepro: TEdit;
    Label5: TLabel;
    cmbDetectado: TComboBox;
    lblCant: TLabel;
    Label6: TLabel;
    cmbEmpleadoDetecto: TComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCancelClick(Sender: TObject);
    procedure BindEmpleados();
    procedure BindTareas();
    procedure FormCreate(Sender: TObject);
    procedure btnScrapClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure chkParcialClick(Sender: TObject);
    procedure txtCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure txtReproKeyPress(Sender: TObject; var Key: Char);
    function BoolToStrInt(Value:Boolean):String;
    function ValidarCantidad(Item:String;Cantidad:Integer):Boolean;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmScrap: TfrmScrap;
  gsConnString,StartDDir: String;
  gsSeleccion : String;
  giConfirmPass : Integer;
implementation

uses Main, Login;
{$R *.dfm}

procedure TfrmScrap.FormClose(Sender: TObject; var Action: TCloseAction);
begin
frmMain.Timer2.Enabled := True;
frmMain.Timer1.Enabled := True;
frmMain.Timer1Timer(nil);

self.Hide;

frmMain.Enabled := True;
setActiveWindow(frmMain.Handle);
end;

procedure TfrmScrap.Button1Click(Sender: TObject);
begin
    frmMain.ChangeStatus(lblOrden.Caption,lblStatus.Caption, false);
    frmMain.txtOrden.Text := '';
    //frmMain.txtOrden.SetFocus;
    Close();
end;

procedure TfrmScrap.btnCancelClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmScrap.BindEmpleados();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT ID,Nombre FROM tblEmpleados Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      cmbEmpleados.Items.Clear;
      cmbEmpleados.Items.Add('000 - Desconocido');
      cmbEmpleadoDetecto.Items.Clear;
      cmbEmpleadoDetecto.Items.Add('000 - Desconocido');
      While not Qry.Eof do
      Begin
          cmbEmpleados.Items.Add(FormatFloat('000',Qry['ID']) + ' - ' + Qry['Nombre']);
          cmbEmpleadoDetecto.Items.Add(FormatFloat('000',Qry['ID']) + ' - ' + Qry['Nombre']);
          Qry.Next;
      End;

      cmbEmpleados.Text := '';
      cmbEmpleadoDetecto.Text := '';
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmScrap.BindTareas();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Nombre FROM tblTareas Order By Nombre';

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      cmbTareas.Items.Clear;
      cmbDetectado.Items.Clear;
      While not Qry.Eof do
      Begin
          cmbTareas.Items.Add(Qry['Nombre']);
          cmbDetectado.Items.Add(Qry['Nombre']);
          Qry.Next;
      End;

      cmbTareas.Text := '';
      cmbDetectado.Text := '';
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;


procedure TfrmScrap.FormCreate(Sender: TObject);
var sUser,sPassword,sServer,sDB : String;
IniFile: TIniFile;
begin
    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','BeltranC');
    sDB := IniFile.ReadString('Conn','DB','Larco');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','');
    giIntervalo := StrToInt(IniFile.ReadString('Tasks','Refresh','30000'));
    giMove := StrToInt(IniFile.ReadString('Tasks','Move','0'));
    giConfirm := StrToInt(IniFile.ReadString('Tasks','Confirm','0'));
    giConfirmPass := StrToInt(IniFile.ReadString('Tasks','ScrapPassword','0'));
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;

    self.Height := 113;
end;

procedure TfrmScrap.btnScrapClick(Sender: TObject);
begin
txtMotivo.Text := '';
self.Height := 406;
BindEmpleados;
BindTareas;
gsSeleccion := (Sender as TButton).Caption;

txtCantidad.Visible := False;
chkParcial.Visible := False;
lblCantidad.Visible := False;
if gsSeleccion = 'Scrap' then
begin
        txtCantidad.Visible := True;
        chkParcial.Visible := True;
        lblCantidad.Visible := True;
end

end;

procedure TfrmScrap.btnOkClick(Sender: TObject);
var Conn : TADOConnection;
SQLStr,sOpcion : String;
begin
  sOpcion := 'retrabajar';
  if gsSeleccion = 'Scrap' then
        sOpcion := 'scrapear';

  //********************* Validando informacion *********************
  if txtMotivo.Text = '' then
    begin
        ShowMessage('Por favor escribe un motivo...');
        Exit;
    end;

  if cmbTareas.Text = '' then
    begin
      MessageDlg('Area responsable requerida.', mtInformation,[mbOk], 0);
      Exit;
    end;

  if cmbEmpleados.Text = '' then
    begin
      MessageDlg('Empleado responsable requerida.', mtInformation,[mbOk], 0);
      Exit;
    end;

  if cmbDetectado.Text = '' then
    begin
      MessageDlg('Area detectado requerida.', mtInformation,[mbOk], 0);
      Exit;
    end;

  if cmbEmpleadoDetecto.Text = '' then
    begin
      MessageDlg('Empleado que lo detecto requerido.', mtInformation,[mbOk], 0);
      Exit;
    end;

  if (cmbTareas.Items.IndexOf(cmbTareas.Text) = -1) then begin
      MessageDlg('Area responsable Incorrecta : ' + cmbTareas.Text  + '.' + #13 +
                 'Seleccione una de la lista. ', mtInformation,[mbOk], 0);
      Exit;
  end;

  if (cmbEmpleados.Items.IndexOf(cmbEmpleados.Text) = -1) then begin
      MessageDlg('Empleado responsable Incorrecto : ' + cmbEmpleados.Text  + '.' + #13 +
                 'Seleccione uno de la lista. ', mtInformation,[mbOk], 0);
      Exit;
  end;

  if (cmbDetectado.Items.IndexOf(cmbDetectado.Text) = -1) then begin
      MessageDlg('Area detectado Incorrecta : ' + cmbDetectado.Text  + '.' + #13 +
                 'Seleccione una de la lista. ', mtInformation,[mbOk], 0);
      Exit;
  end;

  if (cmbEmpleadoDetecto.Items.IndexOf(cmbEmpleadoDetecto.Text) = -1) then begin
      MessageDlg('Empleado Detecto Incorrecto : ' + cmbEmpleadoDetecto.Text  + '.' + #13 +
                 'Seleccione uno de la lista. ', mtInformation,[mbOk], 0);
      Exit;
  end;

  if gsSeleccion = 'Scrap' then
  begin
        if txtCantidad.Text = '' then
        begin
          MessageDlg('Por favor capture una cantidad.', mtInformation,[mbOk], 0);
          Exit;
        end;

        if not ValidarCantidad( lblOrden.Caption, StrToInt(txtCantidad.Text) ) then
        begin
          MessageDlg('La cantidad scrapeada es mayor que la cantidad de la orden.', mtInformation,[mbOk], 0);
          Exit;
        end;

        if chkParcial.Checked then
        begin
              if txtRepro.Text = '' then
              begin
                MessageDlg('Por favor capture una cantidad a reprogramar.', mtInformation,[mbOk], 0);
                Exit;
              end;
        end;
  end;

  //***************************************************************

  if MessageDlg('Estas seguro que quieres ' + sOpcion + ' esta orden ' +
                lblOrden.Caption + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  begin
      Exit;
  end;

  Conn := nil;
  try
  begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;

    if gsSeleccion = 'Scrap' then
    begin
          if giConfirmPass = 1 then begin
              Application.CreateForm(TfrmLogin, frmLogin);
              if frmLogin.ShowModal <> mrOK then begin
                    ShowMessage('No tienes permiso para scrapear esta orden.');
                    Exit;
              end;
          end;


          if txtRepro.Text = '' then txtRepro.Text := '0';

          SQLStr := 'INSERT INTO tblScrap(ITE_Nombre,SCR_Motivo,SCR_Tarea,SCR_EmpleadoRes,SCR_Cantidad,' +
                    'SCR_Parcial,SCR_Repro,USE_Login,SCR_Fecha,SCR_NewItem,SCR_Impreso,SCR_Activo,SCR_Detectado,' +
                    'SCR_EmpleadoDetectado, Update_Date, Update_User) ' +
                    'VALUES(' + QuotedStr(lblOrden.Caption) + ',' +
                    QuotedStr(txtMotivo.Text) + ',' + QuotedStr(cmbTareas.Text) + ',' +
                    QuotedStr(LeftStr(cmbEmpleados.Text,3)) + ',' + txtCantidad.Text + ',' +
                    BoolToStrInt(chkParcial.Checked) + ',' + txtRepro.Text +
                    ',' + QuotedStr(lblEmpleado.Caption) + ',GetDate(),NULL,0,0,' +
                    QuotedStr(cmbDetectado.Text) + ',' + LeftStr(cmbEmpleadoDetecto.Text,3) + ', GetDate(),' +
                    lblEmpleado.Caption + ')';

          Conn.Execute(SQLStr);
      end
      else begin
          SQLStr := 'INSERT INTO tblRetrabajo VALUES(' + QuotedStr(lblOrden.Caption) + ',' +
                    QuotedStr(txtMotivo.Text) + ',' + QuotedStr(cmbTareas.Text) + ',' +
                    QuotedStr(LeftStr(cmbEmpleados.Text,3)) + ',GetDate(),NULL,' +
                    QuotedStr(cmbDetectado.Text) + ',' + LeftStr(cmbEmpleadoDetecto.Text,3) + ', GetDate(),' +
                    lblEmpleado.Caption + ')';

          Conn.Execute(SQLStr);

          frmMain.ChangeStatus(lblOrden.Caption, '5', false);
      end;
    end
    finally
      if Conn <> nil then Conn.Close;
    end;

    Close;
end;

procedure TfrmScrap.chkParcialClick(Sender: TObject);
begin
lblRepro.Visible := chkParcial.Checked;
txtRepro.Visible := chkParcial.Checked;
end;

procedure TfrmScrap.txtCantidadKeyPress(Sender: TObject; var Key: Char);
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

procedure TfrmScrap.txtReproKeyPress(Sender: TObject; var Key: Char);
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

function TfrmScrap.BoolToStrInt(Value:Boolean):String;
begin
        Result := '0';
        if Value Then
                Result := '1';
end;

function TfrmScrap.ValidarCantidad(Item:String;Cantidad:Integer):Boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := True;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT Top 1 Ordenada FROM tblOrdenes WHERE ITE_Nombre = ' + QuotedStr(Item);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      if Qry.RecordCount > 0 then
          if Cantidad > StrToInt(VarToStr(Qry['Ordenada'])) then
                  Result := False;

    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;

end;


procedure TfrmScrap.FormActivate(Sender: TObject);
begin
        Button1.SetFocus();
end;

end.
