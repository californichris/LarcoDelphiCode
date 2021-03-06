unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls, ImgList,Imprimir,chris_functions,Clipbrd,StrUtils,Larco_Functions;

type
  TfrmMain = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    txtEmpleado: TEdit;
    txtOrden: TEdit;
    gvListos: TGridView;
    gvActivos: TGridView;
    gvTerminados: TGridView;
    lblListo: TLabel;
    lblActivo: TLabel;
    lblTerm: TLabel;
    btnActivo: TButton;
    btnTerminado: TButton;
    gvPropiedades: TGridView;
    Timer1: TTimer;
    Button1: TButton;
    imlGrid: TImageList;
    lblTotal: TLabel;
    lblAtras: TLabel;
    lblTerminadas: TLabel;
    Image1: TImage;
    Label6: TLabel;
    Image2: TImage;
    Label7: TLabel;
    btnPrint: TButton;
    Label8: TLabel;
    ddlTareas: TComboBox;
    lblQuery: TLabel;
    lblTerminado: TLabel;
    Timer2: TTimer;
    PopupMenu1: TPopupMenu;
    Copiar1: TMenuItem;
    Copiarcomo1: TMenuItem;
    Separadoporcomas1: TMenuItem;
    Encomillas1: TMenuItem;
    CopiaraOrden1: TMenuItem;
    lblAntes: TLabel;
    procedure ActivarOrden(Orden,Task,User : String);
    function GetStatus(value:integer):String;
    function GetStatusDes(value:integer):String;
    procedure MoverOrden();
    function ValidateOrden(Orden: String; var Msg,Task,Status: String):Boolean;
    procedure ChangeStatus(Orden,Status,Task : String);
    function IsActive(Orden: String):Boolean;
    function IsReady(Orden: String):Boolean;

    procedure BindItemDetail(Item: String; Status: String; Task:String);
    procedure BindAll();
    procedure BindListos();
    procedure BindActivos();
    procedure BindTerminados();
    procedure BindAnteriores();
    procedure BindStats();
    procedure BindTareas();
    procedure FormCreate(Sender: TObject);
    procedure gvListosSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure txtEmpleadoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtOrdenKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtEmpleadoKeyPress(Sender: TObject; var Key: Char);
    procedure txtOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure gvListosKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Timer1Timer(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure ddlTareasChange(Sender: TObject);
    function FormIsRunning(FormName: String):Boolean;
    procedure Timer2Timer(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure Separadoporcomas1Click(Sender: TObject);
    procedure Encomillas1Click(Sender: TObject);
    procedure CopiaraOrden1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  gsConnString,StartDDir: String;
  gsTask,gsRoute,gsHideTask: String;
  giIntervalo,giDelete,giMove,giConfirm,giTerminados: Integer;
implementation
{$R WinXP.res}
{$R *.dfm}

procedure TfrmMain.FormCreate(Sender: TObject);
Var  i: Integer;
CurParam: String;
sUser,sPassword,sServer,sDB : String;
IniFile: TIniFile;
begin

  //Check if are Command Parameters
  If ParamCount = 0 then
    Begin
        MessageDlg('No Command Parameters Found', mtInformation,[mbOk], 0);
        Application.Terminate;
        Exit;
    End;

  //If Command Parameters Found
  for i := 1 to ParamCount do
    begin
      CurParam := ParamStr(i);
      case Clarify(ParamType(CurParam), ' ')[1] of
        'T': gsTask := ParamValue(CurParam);
        'R': gsRoute := ParamValue(CurParam);
      end;//Case
    end;//for


    StartDDir := ExtractFileDir(ParamStr(0)) + '\';
    IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

    sServer := IniFile.ReadString('Conn','Server','BeltranC');
    sDB := IniFile.ReadString('Conn','DB','Larco');
    sUser := IniFile.ReadString('Conn','User','sa');
    sPassword := IniFile.ReadString('Conn','Password','');
    giIntervalo := StrToInt(IniFile.ReadString('Tasks','Refresh','30000'));
    giDelete := StrToInt(IniFile.ReadString('Tasks','DeleteEmp','60000'));
    giMove := StrToInt(IniFile.ReadString('Tasks','Move','0'));
    giConfirm := StrToInt(IniFile.ReadString('Tasks','Confirm','0'));
    giTerminados := StrToInt(IniFile.ReadString('Tasks','Terminados','0'));
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;

    Self.Caption := Self.Caption + gsTask + ' 3.1';
    if giTerminados = 1 then begin
        lblTerminado.Caption := ' AND ITS_DTStop > DATEADD(dd,-4,GETDATE()) ';
    end;

    Application.Title := gsTask;
    Timer1.Interval := giIntervalo;
    Timer2.Interval := giDelete;

    BindTareas;
    BindAll;
end;

procedure TfrmMain.BindAll();
begin
    //BindListosActivosRetrabajo;
    BindListos;
    BindActivos;
    BindTerminados;
    BindAnteriores;
    BindStats();
end;

procedure TfrmMain.BindListos();
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre,' +
                'RIGHT(T.Nombre,LEN(T.Nombre) - LEN(' + QuotedStr(gsTask) + ')) AS Task, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ' +
                'WHERE I.ITS_Status = 0 AND ';

      if ddlTareas.Text = 'Todas' then
              SQLStr := SQLStr + 'LEFT(T.Nombre,LEN(' + QuotedStr(gsTask) + ')) = ' + QuotedStr(gsTask)
      else
              SQLStr := SQLStr + 'T.Nombre = ' + QuotedStr(ddlTareas.Text);

      SQLStr := SQLStr + lblQuery.Caption;
      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvListos.ClearRows;
      while not Qry.Eof do begin
        gvListos.AddRow(1);
        gvListos.Cells[0,gvListos.RowCount -1] := VarToStr(Qry['ITE_ID']);
        gvListos.Cells[1,gvListos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        gvListos.Cells[2,gvListos.RowCount -1] := VarToStr(Qry['Task']);
        gvListos.Cell[3,gvListos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
        Qry.Next;
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.BindActivos();
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre,' +
                'RIGHT(T.Nombre,LEN(T.Nombre) - LEN(' + QuotedStr(gsTask) + ')) AS Task, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ' +
                'WHERE I.ITS_Status = 1 AND ';

      if ddlTareas.Text = 'Todas' then
              SQLStr := SQLStr + 'LEFT(T.Nombre,LEN(' + QuotedStr(gsTask) + ')) = ' + QuotedStr(gsTask)
      else
              SQLStr := SQLStr + 'T.Nombre = ' + QuotedStr(ddlTareas.Text);

      SQLStr := SQLStr + lblQuery.Caption;
      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvActivos.ClearRows;
      while not Qry.Eof do begin
        gvActivos.AddRow(1);
        gvActivos.Cells[0,gvActivos.RowCount -1] := VarToStr(Qry['ITE_ID']);
        gvActivos.Cells[1,gvActivos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        gvActivos.Cells[2,gvActivos.RowCount -1] := VarToStr(Qry['Task']);
        gvActivos.Cell[3,gvActivos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
        Qry.Next;
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.BindTerminados();
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre,' +
                'RIGHT(T.Nombre,LEN(T.Nombre) - LEN(' + QuotedStr(gsTask) + ')) AS Task, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ' +
                'WHERE I.ITS_Status = 2 AND ';

      if ddlTareas.Text = 'Todas' then
              SQLStr := SQLStr + 'LEFT(T.Nombre,LEN(' + QuotedStr(gsTask) + ')) = ' + QuotedStr(gsTask)
      else
              SQLStr := SQLStr + 'T.Nombre = ' + QuotedStr(ddlTareas.Text);

      SQLStr := SQLStr + lblTerminado.Caption;
      SQLStr := SQLStr + lblQuery.Caption;
      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvTerminados.ClearRows;
      while not Qry.Eof do begin
        gvTerminados.AddRow(1);
        gvTerminados.Cells[0,gvTerminados.RowCount -1] := VarToStr(Qry['ITE_ID']);
        gvTerminados.Cells[1,gvTerminados.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
        gvTerminados.Cells[2,gvTerminados.RowCount -1] := VarToStr(Qry['Task']);
        gvTerminados.Cell[3,gvTerminados.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
        Qry.Next;
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.BindAnteriores();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
iAntes,i : Integer;
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

      iAntes := 0;
      if ddlTareas.Text = 'Todas' then
      begin

            for i:=0 to ddlTareas.Items.Count - 1 do
            begin
              if ddlTareas.Items.Strings[i] <> 'Todas' then begin
                  SQLStr := 'Ordenes_Antes_Tarea ' + QuotedStr(ddlTareas.Items.Strings[i]) + ',' + QuotedStr(lblQuery.Caption);

                  Qry.SQL.Clear;
                  Qry.SQL.Text := SQLStr;
                  Qry.Open;

                  iAntes := iAntes+ StrToInt( VarToStr(Qry['Antes']) );
              end;
            end;
      end
      else begin
            SQLStr := 'Ordenes_Antes_Tarea ' + QuotedStr(ddlTareas.Text) + ',' + QuotedStr(lblQuery.Caption);

            Qry.SQL.Clear;
            Qry.SQL.Text := SQLStr;
            Qry.Open;

            iAntes := StrToInt( VarToStr(Qry['Antes']) );
      end;

      lblAntes.Caption := 'Ordenes Antes de ' + gsTask + ': 0';
      if Qry.RecordCount >= 0 then begin
          lblAntes.Caption := 'Ordenes Antes de ' + gsTask + ' : ' + IntToStr(iAntes);
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.BindStats();
var i,iAtras:Integer;
begin
    iAtras := 0;
    lblTotal.Caption := 'Total de Ordenes en ' + gsTask + ' : ' + IntToStr(gvListos.RowCount + gvActivos.RowCount);
    lblTerminadas.Caption := 'Total de Ordenes Terminadas : ' + IntToStr(gvTerminados.RowCount);

    for i:=0 to gvListos.RowCount -1 do
      begin
            if gvListos.Cell[3,i].AsInteger = 1 then
                iAtras := iAtras + 1;
      end;

    for i:=0 to gvActivos.RowCount -1 do
      begin
            if gvActivos.Cell[3,i].AsInteger = 1 then
                iAtras := iAtras + 1;
      end;


    lblAtras.Caption := 'Total de Ordenes Atrasadas : ' + IntToStr(iAtras);
end;

procedure TfrmMain.BindTareas();
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

      SQLStr := 'SELECT Nombre FROM tblTareas ' +
                'WHERE LEFT(Nombre,LEN(' + QuotedStr(gsTask) + ')) = ' + QuotedStr(gsTask);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      ddlTareas.Clear;
      ddlTareas.Items.Add('Todas');
      while not Qry.Eof do begin
          ddlTareas.Items.Add(VarToStr(Qry['Nombre']));
          Qry.Next;
      end;

      ddlTareas.Text := 'Todas';
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;



procedure TfrmMain.BindItemDetail(Item: String; Status:String; Task:String);
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
      Qry.Connection := Conn;

      if Status = 'gvListos' then
        begin
              SQLStr := 'SELECT O.*,IT.*,I.ITE_Priority FROM tblOrdenes O ' +
                        'INNER JOIN tblItems I ON O.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblItemTasks IT ON IT.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblTareas T ON T.[ID] = IT.TAS_ID AND T.Nombre = ' + QuotedStr(gsTask + Task) + ' ' +
                        'WHERE O.ITE_ID = ' + Item;
        end
      else
        begin
              SQLStr := 'SELECT O.*,IT.*,CASE WHEN E.Nombre IS NULL THEN '''' ELSE E.Nombre END AS Empleado,' +
                        'I.ITE_Priority FROM tblOrdenes O ' +
                        'INNER JOIN tblItems I ON O.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblItemTasks IT ON IT.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblTareas T ON T.[ID] = IT.TAS_ID AND T.Nombre = ' + QuotedStr(gsTask + Task) + ' ' +
                        'LEFT OUTER JOIN tblEmpleados E ON IT.USE_LOGIN = E.[ID] ' +
                        'WHERE O.ITE_ID = ' + Item;
        end;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvPropiedades.ClearRows;
      if Qry.RecordCount > 0 then
      begin
          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Tipo Proceso';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['TipoProceso']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Cantidad Requerida';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Requerida']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Cantidad Ordenada';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Ordenada']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Descripcion';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Producto']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Numero';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Numero']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Terminal';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Terminal']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Fecha Recibido';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Recibido']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Fecha Interna';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Interna']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Fecha Entrega';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Entrega']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Nombre';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Nombre']);

          gvPropiedades.AddRow(1);
          gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Prioridad';
          gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['ITE_Priority']);

          if ( (VarToStr(Qry['ITS_Status']) <> '0') ) Then
            begin
                  gvPropiedades.AddRow(1);
                  gvPropiedades.AddRow(1);
                  gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Tarea';
                  gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := gsTask + Task;

                  //gvPropiedades.AddRow(1);
                  //gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'No.Empleado';
                  //gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := UT(VarToStr(Qry['USE_Login']));

                  gvPropiedades.AddRow(1);
                  gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Empleado';
                  gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['Empleado']);

                  gvPropiedades.AddRow(1);
                  gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Inicio';
                  gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['ITS_DTStart']);

                  gvPropiedades.AddRow(1);
                  gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Termino';
                  gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := VarToStr(Qry['ITS_DTStop']);
            end;
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.gvListosSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
BindItemDetail((Sender As TGridView).Cells[0,ARow],(Sender As TGridView).Name,(Sender As TGridView).Cells[2,ARow]);
end;

procedure TfrmMain.txtEmpleadoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = vk_Return then
  begin
        txtOrden.SetFocus;
  end
  else if key = vk_F5 then
  begin
    BindAll;
  end;
end;

procedure TfrmMain.txtOrdenKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if key = vk_Return then
  begin
      MoverOrden();
  end
  else if key = vk_F5 then
  begin
    BindAll;
  end;
end;

procedure TfrmMain.MoverOrden();
var Task,Status,Msg : String;
begin
  Timer2.Enabled := False;
  Timer1.Enabled := False;
  if not ValidateEmpleado(gsConnString, txtEmpleado.Text) Then
    begin
          ShowMessage('Numero de empleado incorrecto o empleado inactivo.');
          Timer1.Enabled := True;
          Exit;
    end;

   if ValidateOrden(txtOrden.Text,Msg,Task,Status) then
   begin
        if gsTask = Task then
        begin
          if giConfirm = 1 then
          begin
              if MessageDlg('Estas seguro que quieres ' + GetStatusDes(StrToInt(Status) + 1) +
              ' la orden ' + txtOrden.Text  + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
              begin
                  Timer1.Enabled := True;
                  Exit;
              end
              else begin
                  ChangeStatus(txtOrden.Text,Status,Task);
                  txtOrden.Text := '';
                  txtOrden.SetFocus;
                  Timer2.Enabled := True;
                  Timer1.Enabled := True;
                  Exit;
              end;
          end
          else begin
            ChangeStatus(txtOrden.Text,Status,Task);
            txtOrden.Text := '';
            txtOrden.SetFocus;
            Timer2.Enabled := True;
            Timer1.Enabled := True;
            Exit;
          end;
        end;

        if (Task = 'Ventas') and (Status = '1') then begin
            ShowMessage('La orden esta pendiente de Plano, no puede entrar a produccion.');
            Timer1.Enabled := True;
            Exit;
        end;

        if MessageDlg('La Orden esta ' + GetStatus(StrToInt(Status)) + ' en la tarea ' + Task + '.' + #13 +
         'Quieres activar la orden en esta tarea?',
          mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
            txtOrden.SetFocus;
            Timer1.Enabled := True;
            Exit;
        end
        else begin
                ActivarOrden(txtOrden.Text,gsTask,txtEmpleado.Text);
                txtOrden.Text := '';
                txtOrden.SetFocus;
                Timer2.Enabled := True;
                Timer1.Enabled := True;
                Exit;
        end;
   end
   else begin
        if LeftStr(Task,Length(gsTask)) = gsTask then begin
            if giConfirm = 1 then
            begin
                if MessageDlg('Estas seguro que quieres ' + GetStatusDes(StrToInt(Status) + 1) +
                ' la orden ' + txtOrden.Text  + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
                begin
                    Timer1.Enabled := True;
                    Exit;
                end
                else begin
                    ChangeStatus(txtOrden.Text,Status,Task);
                    txtOrden.Text := '';
                    txtOrden.SetFocus;
                    Timer1.Enabled := True;
                    Timer2.Enabled := True;
                    Exit;
                end;
            end
            else begin
              ChangeStatus(txtOrden.Text,Status,Task);
              txtOrden.Text := '';
              txtOrden.SetFocus;
              Timer2.Enabled := True;
              Timer1.Enabled := True;
              Exit;
            end;


        end
        else begin
            if Task = '' then
                    MessageDlg(Msg, mtInformation, [mbOK],0)
            else
                    MessageDlg(Msg + chr(13) + 'Se encuentra ' + GetStatus(StrToInt(Status)) + ' en ' + Task + '.', mtInformation, [mbOK],0);
        end;
   end;

   Timer1.Enabled := True;
   txtOrden.SetFocus;
end;

procedure TfrmMain.txtEmpleadoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['0'..'9'] then begin
  end
  else if (Key = Chr(vk_Back)) then begin
  end
  else
    Key := #0;

end;

procedure TfrmMain.txtOrdenKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['0'..'9'] then begin
  end
  else if (Key = Chr(vk_Back)) then begin
  end
  else if (Key = '-') then begin
  end
  else
    Key := #0;

end;

function TfrmMain.IsActive(Orden: String):Boolean;
var i:integer;
begin
  Result := False;
  for i:=0 to gvActivos.RowCount -1 do begin
    if UT(gvActivos.Cells[1,i]) = UT(Orden) then
    begin
        Result := True;
        gsHideTask := gvActivos.Cells[2,i]
    end;
  end;
end;

function TfrmMain.IsReady(Orden: String):Boolean;
var i:integer;
begin
  Result := False;
  for i:=0 to gvListos.RowCount -1 do begin
    if UT(gvListos.Cells[1,i]) = UT(Orden) then
    begin
        Result := True;
        gsHideTask := gvListos.Cells[2,i]
    end;
  end;

end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = vk_F5 then begin
    BindAll;
  end;
end;

procedure TfrmMain.gvListosKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = vk_F5 then begin
    BindAll;
  end;

end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
Var  IniFile: TIniFile;
begin
  StartDDir := ExtractFileDir(ParamStr(0)) + '\';
  IniFile := TiniFile.Create(StartDDir + 'Larco.ini');

  giIntervalo := StrToInt(IniFile.ReadString('Tasks','Refresh','30000'));
  Timer1.Interval := giIntervalo;

  BindAll;
end;

procedure TfrmMain.ChangeStatus(Orden,Status,Task : String);
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
      Qry.Connection := Conn;

      SQLStr := 'ChangeStatus ' + QuotedStr(Orden) + ',' + QuotedStr(Task) +
                ',' + Status + ',' + QuotedStr(txtEmpleado.Text) +
                ',' + QuotedStr(GetLocalIP);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      if Qry.RecordCount <= 0 then begin
         MessageDlg('Ocurrio un error mientras se cambiaba la orden de status.', mtInformation, [mbOK],0);
         Exit;
      end
      else
        begin
          if VarToStr(Qry['Error']) = '1' then begin
                 MessageDlg(VarToStr(Qry['MSG']), mtInformation, [mbOK],0);
                 Exit;
            end
          else
            begin
                 if Status = '1' Then
                         MessageDlg('Por favor lleve la orden ' + Orden + ' a la tarea ' + VarToStr(Qry['Msg']) + '.', mtInformation, [mbOK],0);
            end;

        end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;

  BindAll;
  txtempleado.SetFocus;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin

  if Button1.Caption = '<<' then begin
    Self.Width := 466;
    Button1.Caption := '>>';
    Button1.Hint := 'Mostrar propiedes de la orden.';
  end
  else begin
    Self.Width := 681;
    Button1.Caption := '<<';
    Button1.Hint := 'Esconder propiedes de la orden.';
  end;

end;

procedure TfrmMain.btnPrintClick(Sender: TObject);
begin
  if FormIsRunning('frmImprimir') then begin
    setActiveWindow(frmImprimir.Handle);
    frmImprimir.WindowState := wsNormal;
    frmImprimir.lblTask.Caption := gsTask;
    frmImprimir.lblTask2.Caption := ddlTareas.Text;
    frmImprimir.Visible := true;
  end
  else begin
    Application.CreateForm(TfrmImprimir,frmImprimir);
    frmImprimir.lblTask.Caption := gsTask;
    frmImprimir.lblTask2.Caption := ddlTareas.Text;
    frmImprimir.Show;
  end;
  self.Enabled := false;
end;

procedure TfrmMain.ddlTareasChange(Sender: TObject);
begin
  BindAll;
end;

function TfrmMain.FormIsRunning(FormName: String):Boolean;
var i:Integer;
begin
  Result := False;

  for  i := 0 to Screen.FormCount - 1 do
  begin
        if Screen.Forms[i].Name = FormName Then
          begin
                Result:= True;
                Break;
          end;
  end;

end;


procedure TfrmMain.Timer2Timer(Sender: TObject);
begin
  txtEmpleado.Text := '';
  Timer2.Enabled := False;
  txtEmpleado.SetFocus;
end;

function TfrmMain.ValidateOrden(Orden: String; var Msg,Task,Status: String):Boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    //Result := False;
    
    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection := Conn;

      SQLStr := 'ItemStatus ' + QuotedStr(Orden) + ',' + QuotedStr(gsTask);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      if Qry.RecordCount <= 0 then
      begin
          MessageDlg('La orden no existe en el sistema.', mtInformation, [mbOK],0);
          Result := False;
      end;

      Msg := Qry['Msg'];
      Task := Qry['Task'];
      Status := Qry['Status'];

      if Qry['Res'] = 0 then
      begin
          Result := False;
      end
      else begin
  {        Msg := Qry['Msg'];
          Task := VartoStr(Qry['Task']);
          Status := GetStatus(Qry['Status']);}
          Result := True;
      end;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

function TfrmMain.GetStatusDes(value:integer):String;
begin
if value = 0 then
        Result := 'Listo'
else if value = 1 then
        Result := 'Activar'
else if value = 2 then
        Result := 'Terminar';
end;

function TfrmMain.GetStatus(value:integer):String;
begin
if value = 0 then
        Result := 'Listo'
else if value = 1 then
        Result := 'Activo'
else if value = 2 then
        Result := 'Terminado';
end;

procedure TfrmMain.ActivarOrden(Orden,Task,User : String);
var Conn : TADOConnection;
SQLStr : String;
begin
  Conn := nil;
  try
  begin
    Conn := TADOConnection.Create(nil);
    Conn.ConnectionString := gsConnString;
    Conn.LoginPrompt := False;

    SQLStr := 'ActivarOrden ' + QuotedStr(Orden) + ',' + QuotedStr(Task) + ',' +
              QuotedStr(User) + ',' + QuotedStr(GetLocalIP);

    Conn.Execute(SQLStr);
  end
  finally
    if Conn <> nil then Conn.Close;
  end;

  BindAll;
end;


procedure TfrmMain.Copiar1Click(Sender: TObject);
begin
  if PopupMenu1.PopupComponent = gvActivos then
     Clipboard.AsText := gvActivos.Cells[1,gvActivos.SelectedRow]
  else if PopupMenu1.PopupComponent = gvListos then
     Clipboard.AsText := gvListos.Cells[1,gvListos.SelectedRow]
  else if PopupMenu1.PopupComponent = gvTerminados then
     Clipboard.AsText := gvTerminados.Cells[1,gvTerminados.SelectedRow];
end;

procedure TfrmMain.Separadoporcomas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
  sText := '';
  if PopupMenu1.PopupComponent = gvActivos then begin
     for i:= 0 to gvActivos.RowCount - 1 do
             sText := sText + gvActivos.Cells[1,i] + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end
  else if PopupMenu1.PopupComponent = gvListos then begin
     for i:= 0 to gvListos.RowCount - 1 do
             sText := sText + gvListos.Cells[1,i] + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end
  else if PopupMenu1.PopupComponent = gvTerminados then begin
     for i:= 0 to gvTerminados.RowCount - 1 do
             sText := sText + gvTerminados.Cells[1,i] + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end;
end;

procedure TfrmMain.Encomillas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
  sText := '';
  if PopupMenu1.PopupComponent = gvActivos then
  begin
     for i:= 0 to gvActivos.RowCount - 1 do
             sText := sText + QuotedStr(gvActivos.Cells[1,i]) + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end
  else if PopupMenu1.PopupComponent = gvListos then begin
     for i:= 0 to gvListos.RowCount - 1 do
             sText := sText + QuotedStr(gvListos.Cells[1,i]) + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end
  else if PopupMenu1.PopupComponent = gvTerminados then begin
     for i:= 0 to gvTerminados.RowCount - 1 do
             sText := sText + QuotedStr(gvTerminados.Cells[1,i]) + ',';

     Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
  end;
end;

procedure TfrmMain.CopiaraOrden1Click(Sender: TObject);
begin
  if PopupMenu1.PopupComponent = gvActivos then
     txtOrden.Text := gvActivos.Cells[1,gvActivos.SelectedRow]
  else if PopupMenu1.PopupComponent = gvListos then
     txtOrden.Text := gvListos.Cells[1,gvListos.SelectedRow]
  else if PopupMenu1.PopupComponent = gvTerminados then
     txtOrden.Text := gvTerminados.Cells[1,gvTerminados.SelectedRow];

  txtOrden.SetFocus;
end;

End.

