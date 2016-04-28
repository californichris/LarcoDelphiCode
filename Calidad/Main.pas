unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,ADODB,DB,IniFiles,All_Functions, chris_Functions,StdCtrls, ScrollView,
  CustomGridViewControl, CustomGridView, GridView, Menus,LTCUtils, Buttons,
  ExtCtrls, ImgList,Imprimir,Clipbrd,StrUtils,Larco_Functions;

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
    lblQuery: TLabel;
    PopupMenu1: TPopupMenu;
    Copiar1: TMenuItem;
    Copiarcomo1: TMenuItem;
    Separadoporcomas1: TMenuItem;
    Encomillas1: TMenuItem;
    CopiaraOrden1: TMenuItem;
    gvRetrabajo: TGridView;
    lblRetrabajo: TLabel;
    lblTerminado: TLabel;
    Timer2: TTimer;
    lblAntes: TLabel;
    function GetStatus(value:integer):String;
    function GetStatusDes(value:integer):String;
    procedure ChangeStatus(Orden,Status : String; update: Boolean);
    procedure ActivarOrden(Orden,Task,User : String);
    procedure MoverOrden();
    function IsActive(Orden: String):Boolean;
    function IsReady(Orden: String):Boolean;
    function ValidateOrden(Orden: String; var Msg,Task,Status: String):Boolean;
    procedure BindItemDetail(Item: String; Status: String);
    procedure BindAll();
    procedure BindListosActivosRetrabajo();
    procedure BindListos();
    procedure BindActivos();
    procedure BindRetrabajo();
    procedure BindTerminados();
    procedure BindAnteriores();
    procedure BindStats();
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
    function FormIsRunning(FormName: String):Boolean;
    procedure btnActivoClick(Sender: TObject);
    procedure btnTerminadoClick(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure Separadoporcomas1Click(Sender: TObject);
    procedure Encomillas1Click(Sender: TObject);
    procedure CopiaraOrden1Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  gsConnString,StartDDir: String;
  gsTask,gsRoute: String;
  giIntervalo,giMove,giConfirm,giTerminados,giDelete : Integer;
implementation

uses Scrap;
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
        MessageDlg('No hay parametros en el shortcut.', mtInformation,[mbOk], 0);
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

    if gsTask = '' Then
    Begin
        MessageDlg('La tarea que estableciste en el icono no es valida.', mtInformation,[mbOk], 0);
        Application.Terminate;
        Exit;
    End;

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
    gsConnString := 'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=' + sUser +
                   ';Password= ' + sPassword +'; Initial Catalog=' + sDB + ';Data Source=' + sServer;
    giTerminados := StrToInt(IniFile.ReadString('Tasks','Terminados','0'));

    Self.Caption := Self.Caption + gsTask + ' 2.3';

    if giTerminados = 1 then begin
        lblTerminado.Caption := ' AND I.ITS_DTStop > DATEADD(dd,-4,GETDATE()) ';
    end;

    Application.Title := gsTask;
    Timer1.Interval := giIntervalo;
    Timer2.Interval := giDelete;

    BindAll();
end;

procedure TfrmMain.BindAll();
begin
    BindListosActivosRetrabajo;
    //BindListos;
    //BindActivos;
    //BindRetrabajo;
    BindTerminados;
    BindAnteriores;
    BindStats();
end;

procedure TfrmMain.BindListosActivosRetrabajo();
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr, status : String;
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status, I.ITS_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] AND T.Nombre = ' + QuotedStr(gsTask) + ' AND I.ITS_Status in (0,1,3) ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblQuery.Caption <> '' then begin
        SQLStr := SQLStr + 'WHERE 1=1 ' + lblQuery.Caption;
      end;

      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvListos.ClearRows;
      gvActivos.ClearRows;
      gvRetrabajo.ClearRows;
      while not Qry.Eof do
      begin
          status := VarToStr(Qry['ITS_Status']);
          if status = '0' then begin
            gvListos.AddRow(1);
            gvListos.Cells[0,gvListos.RowCount -1] := VarToStr(Qry['ITE_ID']);
            gvListos.Cells[1,gvListos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
            gvListos.Cell[2,gvListos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
          end
          else if status = '1' then begin
            gvActivos.AddRow(1);
            gvActivos.Cells[0,gvActivos.RowCount -1] := VarToStr(Qry['ITE_ID']);
            gvActivos.Cells[1,gvActivos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
            gvActivos.Cell[2,gvActivos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
          end
          else if status = '3' then begin
            gvRetrabajo.AddRow(1);
            gvRetrabajo.Cells[0,gvRetrabajo.RowCount -1] := VarToStr(Qry['ITE_ID']);
            gvRetrabajo.Cells[1,gvRetrabajo.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
            gvRetrabajo.Cell[2,gvRetrabajo.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
          end;
          Qry.Next;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] AND T.Nombre = ' + QuotedStr(gsTask) + ' AND I.ITS_Status = 0 ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblQuery.Caption <> '' then begin
        SQLStr := SQLStr + 'WHERE ' + lblQuery.Caption;
      end;

      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      gvListos.ClearRows;
      while not Qry.Eof do
      begin
          gvListos.AddRow(1);
          gvListos.Cells[0,gvListos.RowCount -1] := VarToStr(Qry['ITE_ID']);
          gvListos.Cells[1,gvListos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
          gvListos.Cell[2,gvListos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] AND T.Nombre = ' + QuotedStr(gsTask) + ' AND I.ITS_Status = 1 ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblQuery.Caption <> '' then begin
        SQLStr := SQLStr + 'WHERE ' + lblQuery.Caption;
      end;

      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvActivos.ClearRows;
      while not Qry.Eof do
      begin
          gvActivos.AddRow(1);
          gvActivos.Cells[0,gvActivos.RowCount -1] := VarToStr(Qry['ITE_ID']);
          gvActivos.Cells[1,gvActivos.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
          gvActivos.Cell[2,gvActivos.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
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

      {SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] AND T.Nombre = ' + QuotedStr(gsTask) + ' AND I.ITS_Status = 2 ' + lblTerminado.Caption +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblQuery.Caption <> '' then begin
        SQLStr := SQLStr + 'WHERE ' + lblQuery.Caption;
      end;

      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;
      }
      SQLStr := 'Traer_Terminadas ' + QuotedStr(gsTask) + ',' + QuotedStr(lblQuery.Caption);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvTerminados.ClearRows;
      While not Qry.Eof do
      Begin
          gvTerminados.AddRow(1);
          gvTerminados.Cells[0,gvTerminados.RowCount -1] := VarToStr(Qry['ITE_ID']);
          gvTerminados.Cells[1,gvTerminados.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
          gvTerminados.Cell[2,gvTerminados.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
          Qry.Next;
      End;
    end
    finally
      if Qry <> nil then Qry.Close;
      if Conn <> nil then Conn.Close;
    end;
end;

procedure TfrmMain.BindRetrabajo();
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

      SQLStr := 'SELECT I.ITE_ID,RTRIM(I.ITE_Nombre) AS ITE_Nombre, ' +
                'CASE WHEN ITS_Status = 2 THEN 0 ' +
                'WHEN dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo THEN 1 ' +
                'WHEN dbo.GetHours(O.Interna,GETDATE()) > T.Interno THEN 1 ' +
                'WHEN I2.ITE_Priority > 0.00 THEN 2 ' +
                'ELSE 0 END AS Late, ' +
                'I2.ITE_Status ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] AND T.Nombre = ' + QuotedStr(gsTask) + ' AND I.ITS_Status = 3 ' +
                'INNER JOIN tblItems I2 ON I2.ITE_ID = I.ITE_ID ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblQuery.Caption <> '' then begin
        SQLStr := SQLStr + 'WHERE ' + lblQuery.Caption;
      end;

      SQLStr := SQLStr + ' ORDER BY I.ITS_DTStart Desc' ;
      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      gvRetrabajo.ClearRows;
      while not Qry.Eof do
      begin
          gvRetrabajo.AddRow(1);
          gvRetrabajo.Cells[0,gvRetrabajo.RowCount -1] := VarToStr(Qry['ITE_ID']);
          gvRetrabajo.Cells[1,gvRetrabajo.RowCount -1] := VarToStr(Qry['ITE_Nombre']);
          gvRetrabajo.Cell[2,gvRetrabajo.RowCount -1].AsInteger := StrToInt( VarToStr(Qry['Late']) );
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

      SQLStr := 'Ordenes_Antes_Tarea ' + QuotedStr(gsTask) + ',' + QuotedStr(lblQuery.Caption);

      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      lblAntes.Caption := 'Ordenes Antes de ' + gsTask + ': 0';
      if Qry.RecordCount >= 0 then begin
          lblAntes.Caption := 'Ordenes Antes de ' + gsTask + ' : ' + VarToStr(Qry['Antes']);
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;



procedure TfrmMain.BindStats();
var i,iAtras:Integer;
begin
        iAtras := 0;
        lblTotal.Caption := 'Ordenes en Tarea : ' + IntToStr(gvListos.RowCount + gvActivos.RowCount);
        lblTerminadas.Caption := 'Ordenes Terminadas : ' + IntToStr(gvTerminados.RowCount);

        for i:=0 to gvListos.RowCount -1 do
          begin
                if gvListos.Cell[2,i].AsInteger = 1 then
                    iAtras := iAtras + 1;
          end;

        for i:=0 to gvActivos.RowCount -1 do
          begin
                if gvActivos.Cell[2,i].AsInteger = 1 then
                    iAtras := iAtras + 1;
          end;


        lblAtras.Caption := 'Ordenes Atrasadas : ' + IntToStr(iAtras);
end;


procedure TfrmMain.BindItemDetail(Item: String; Status:String);
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
                        'INNER JOIN tblTareas T ON T.[ID] = IT.TAS_ID AND T.Nombre = ' + QuotedStr(gsTask) + ' ' +
                        'WHERE O.ITE_ID = ' + Item;
        end
      else
        begin
              SQLStr := 'SELECT O.*,IT.*,CASE WHEN E.Nombre IS NULL THEN '''' ELSE E.Nombre END AS Empleado,' +
                        'I.ITE_Priority FROM tblOrdenes O ' +
                        'INNER JOIN tblItems I ON O.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblItemTasks IT ON IT.ITE_ID = I.ITE_ID ' +
                        'INNER JOIN tblTareas T ON T.[ID] = IT.TAS_ID AND T.Nombre = ' + QuotedStr(gsTask) + ' ' +
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

                  //gvPropiedades.AddRow(1);
                  //gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'No.Empleado';
                  //gvPropiedades.Cells[1,gvPropiedades.RowCount -1] := UT(VarToStr(Qry['USE_Login']));

                  gvPropiedades.AddRow(1);
                  gvPropiedades.Cells[0,gvPropiedades.RowCount -1] := 'Nombre';
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
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;

procedure TfrmMain.gvListosSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
BindItemDetail((Sender As TGridView).Cells[0,ARow],(Sender As TGridView).Name);
end;

procedure TfrmMain.txtEmpleadoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Timer2.Enabled := False;
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
iInc : Integer;
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
                    if Status = '9' then begin
                      ShowMessage('Esta orden esta declarada como Scrap.');
                      Timer1.Enabled := True;
                      Exit;
                    end;

                    iInc := 1;
                    if Status = '3' then iInc := 0;
                    if MessageDlg('Estas seguros que quieres ' + GetStatusDes(StrToInt(Status) + iInc) +
                    ' la orden ' + txtOrden.Text  + '?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
                    begin
                        Timer1.Enabled := True;
                        Exit;
                    end
                    else begin
                        if StrToInt(Status) = 1 then
                        begin
                            //************** Execute form *********************************
                            if FormIsRunning('frmScrap') Then
                              begin
                                    setActiveWindow(frmScrap.Handle);
                                    frmScrap.WindowState := wsNormal;
                                    frmScrap.Height := 113;
                                    frmScrap.Visible := true;
                                    frmScrap.lblOrden.Caption := txtOrden.Text;
                                    frmScrap.lblEmpleado.Caption := txtEmpleado.Text;
                                    frmScrap.lblStatus.Caption := Status;
                              end
                            else
                              begin
                                    Application.CreateForm(TfrmScrap,frmScrap);
                                    frmScrap.lblTask.Caption := gsTask;
                                    frmScrap.lblOrden.Caption := txtOrden.Text;
                                    frmScrap.lblEmpleado.Caption := txtEmpleado.Text;
                                    frmScrap.lblStatus.Caption := Status;
                                    frmScrap.Show;
                              end;

                            self.Enabled := false;
                            Exit;
                            //**************************************************************8
                        end;

                        ChangeStatus(txtOrden.Text,Status, true);
                        txtOrden.Text := '';
                        txtOrden.SetFocus;
                        Timer2.Enabled := True;
                        Timer1.Enabled := True;
                        Exit;
                    end;
                end
                else begin
                  ChangeStatus(txtOrden.Text,Status, true);
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
              if Task = '' then
                      MessageDlg(Msg, mtInformation, [mbOK],0)
              else
                      MessageDlg(Msg + chr(13) + 'Se encuentra ' + GetStatus(StrToInt(Status)) + ' en ' + Task + '.', mtInformation, [mbOK],0);
         end;

         txtOrden.SetFocus;
         Timer1.Enabled := True;
end;



function TfrmMain.ValidateOrden(Orden: String; var Msg,Task,Status: String):Boolean;
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr : String;
begin
    Result := False;
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

      Msg := VartoStr(Qry['Msg']);
      Task := Qry['Task'];
      Status := Qry['Status'];

      if Qry['Res'] = 0 then
      begin
          Result := False;
      end
      else begin
          {Msg := Qry['Msg'];
          Task := VartoStr(Qry['Task']);
          Status := GetStatus(Qry['Status']);}
          Result := True;
      end;
    end
    finally
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;
end;



procedure TfrmMain.txtEmpleadoKeyPress(Sender: TObject; var Key: Char);
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

procedure TfrmMain.txtOrdenKeyPress(Sender: TObject; var Key: Char);
begin
        if Key in ['0'..'9'] then
            begin
            end
        else if (Key = Chr(vk_Back)) then
            begin
            end
        else if (Key = '-') then
            begin
            end
       else
                Key := #0;

end;

function TfrmMain.IsActive(Orden: String):Boolean;
var i:integer;
begin
        Result := False;
        for i:=0 to gvActivos.RowCount -1 do
          begin
                if UT(gvActivos.Cells[1,i]) = UT(Orden) then
                    Result := True;
          end;

end;

function TfrmMain.IsReady(Orden: String):Boolean;
var i:integer;
begin
        Result := False;
        for i:=0 to gvListos.RowCount -1 do
          begin
                if UT(gvListos.Cells[1,i]) = UT(Orden) then
                    Result := True;
          end;

end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
        if Key = vk_F5 then
        begin
          BindAll;
        end;
end;

procedure TfrmMain.gvListosKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = vk_F5 then
  begin
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

procedure TfrmMain.ChangeStatus(Orden,Status : String; update: Boolean);
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

      SQLStr := 'ChangeStatus ' + QuotedStr(Orden) + ',' + QuotedStr(gsTask)+
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
      if Qry <> nil then begin
        Qry.Close;
        Qry.Free;
      end;
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

    if(update) then begin
      BindAll;
    end;
//txtempleado.SetFocus;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin

if Button1.Caption = '<<' Then
  Begin
          Self.Width := Self.Width - 213;
          Button1.Caption := '>>';
          Button1.Hint := 'Mostrar propiedes de la orden.';
  end
  else
  begin
          Self.Width := Self.Width + 213;
          Button1.Caption := '<<';
          Button1.Hint := 'Esconder propiedes de la orden.';
  end;

end;

procedure TfrmMain.btnPrintClick(Sender: TObject);
begin
if FormIsRunning('frmImprimir') Then
  begin
        setActiveWindow(frmImprimir.Handle);
        frmImprimir.WindowState := wsNormal;
        frmImprimir.Visible := true;
  end
else
  begin
Application.CreateForm(TfrmImprimir,frmImprimir);
frmImprimir.lblTask.Caption := gsTask;
frmImprimir.Show;
  end;
self.Enabled := false;

end;

function TfrmMain.GetStatus(value:integer):String;
begin
if value = 0 then
        Result := 'Listo'
else if value = 1 then
        Result := 'Activo'
else if value = 2 then
        Result := 'Terminado'
else if value = 3 then
        Result := 'Retrabajo';

end;

function TfrmMain.GetStatusDes(value:integer):String;
begin
if value = 0 then
        Result := 'Listo'
else if value = 1 then
        Result := 'Activar'
else if value = 2 then
        Result := 'Terminar'
else if value = 3 then
        Result := 'Terminar';

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
      if Conn <> nil then begin
        Conn.Close;
        Conn.Free
      end;
    end;

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


procedure TfrmMain.btnActivoClick(Sender: TObject);
begin
MoverOrden();
end;

procedure TfrmMain.btnTerminadoClick(Sender: TObject);
begin
MoverOrden();
end;

procedure TfrmMain.Copiar1Click(Sender: TObject);
begin
        if PopupMenu1.PopupComponent = gvActivos then
           Clipboard.AsText := gvActivos.Cells[1,gvActivos.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvListos then
           Clipboard.AsText := gvListos.Cells[1,gvListos.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvTerminados then
           Clipboard.AsText := gvTerminados.Cells[1,gvTerminados.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvRetrabajo then
           Clipboard.AsText := gvRetrabajo.Cells[1,gvRetrabajo.SelectedRow];

end;

procedure TfrmMain.Separadoporcomas1Click(Sender: TObject);
var i : integer;
sText : String;
begin
        sText := '';
        if PopupMenu1.PopupComponent = gvActivos then
        begin
           for i:= 0 to gvActivos.RowCount - 1 do
                   sText := sText + gvActivos.Cells[1,i] + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end
        Else if PopupMenu1.PopupComponent = gvListos then
        begin
           for i:= 0 to gvListos.RowCount - 1 do
                   sText := sText + gvListos.Cells[1,i] + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end
        Else if PopupMenu1.PopupComponent = gvTerminados then
        begin
           for i:= 0 to gvTerminados.RowCount - 1 do
                   sText := sText + gvTerminados.Cells[1,i] + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end
        Else if PopupMenu1.PopupComponent = gvRetrabajo then
        begin
           for i:= 0 to gvRetrabajo.RowCount - 1 do
                   sText := sText + gvRetrabajo.Cells[1,i] + ',';

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
        Else if PopupMenu1.PopupComponent = gvListos then
        begin
           for i:= 0 to gvListos.RowCount - 1 do
                   sText := sText + QuotedStr(gvListos.Cells[1,i]) + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end
        Else if PopupMenu1.PopupComponent = gvTerminados then
        begin
           for i:= 0 to gvTerminados.RowCount - 1 do
                   sText := sText + QuotedStr(gvTerminados.Cells[1,i]) + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end
        Else if PopupMenu1.PopupComponent = gvRetrabajo then
        begin
           for i:= 0 to gvRetrabajo.RowCount - 1 do
                   sText := sText + QuotedStr(gvRetrabajo.Cells[1,i]) + ',';

           Clipboard.AsText := LeftStr(sText,Length(sText) - 1);
        end;
end;

procedure TfrmMain.CopiaraOrden1Click(Sender: TObject);
begin
        if PopupMenu1.PopupComponent = gvActivos then
           txtOrden.Text := gvActivos.Cells[1,gvActivos.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvListos then
           txtOrden.Text := gvListos.Cells[1,gvListos.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvTerminados then
           txtOrden.Text := gvTerminados.Cells[1,gvTerminados.SelectedRow]
        Else if PopupMenu1.PopupComponent = gvRetrabajo then
           txtOrden.Text := gvRetrabajo.Cells[1,gvRetrabajo.SelectedRow];

        txtOrden.SetFocus;
end;

procedure TfrmMain.Timer2Timer(Sender: TObject);
begin
        txtEmpleado.Text := '';
        Timer2.Enabled := False;
        txtEmpleado.SetFocus;        
end;

End.

