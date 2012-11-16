unit Imprimir;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ADODB,DB,IniFiles,All_Functions,CellEditors, ScrollView,
  CustomGridViewControl, CustomGridView, GridView,StrUtils,Unit3,chris_functions;

type
  TfrmImprimir = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    chkRecibido: TCheckBox;
    deRecibido1: TDateEditor;
    deRecibido2: TDateEditor;
    Panel1: TPanel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    chkInterna: TCheckBox;
    deInterna1: TDateEditor;
    deInterna2: TDateEditor;
    chkEntrega: TCheckBox;
    deEntrega1: TDateEditor;
    deEntrega2: TDateEditor;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GridView1: TGridView;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Label2: TLabel;
    Label3: TLabel;
    Panel2: TPanel;
    gvCampos: TGridView;
    Panel3: TPanel;
    gvSort: TGridView;
    Button4: TButton;
    ddlCampo: TComboBox;
    ddlCondicion: TComboBox;
    ddlValue: TComboBox;
    ddlConjuncion: TComboBox;
    rbtTodas: TRadioButton;
    rbtAtrasadas: TRadioButton;
    lblTask: TLabel;
    lblTask2: TLabel;
    Button5: TButton;
    chkTerminado: TCheckBox;
    deTerminado1: TDateEditor;
    deTerminado2: TDateEditor;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure chkRecibidoClick(Sender: TObject);
    procedure chkInternaClick(Sender: TObject);
    procedure chkEntregaClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure GridView1Click(Sender: TObject);
    procedure GridView1CellClick(Sender: TObject; ACol, ARow: Integer);
    procedure ddlCampoChange(Sender: TObject);
    procedure BindValues(Field : String);
    procedure Button4Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmImprimir: TfrmImprimir;
  gsConnString,StartDDir: String;

  Campos: array[0..7] of string = ('Nombre', 'Producto', 'Recibido', 'Interna', 'Entrega', 'Empleado','Requerida','Ordenada');
  DBNames: array[0..7] of string = ('I.ITE_Nombre', 'O.Producto', 'O.Recibido', 'O.Interna', 'O.Entrega', 'O.Nombre','O.Requerida','O.Ordenada');

implementation

uses Main;

{$R *.dfm}

procedure TfrmImprimir.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//Action := caFree;
self.Hide;
frmMain.Enabled := True;
setActiveWindow(frmMain.Handle);
end;

procedure TfrmImprimir.chkRecibidoClick(Sender: TObject);
begin
deRecibido1.Enabled := chkRecibido.Checked;
deRecibido2.Enabled := chkRecibido.Checked;
end;

procedure TfrmImprimir.chkInternaClick(Sender: TObject);
begin
deInterna1.Enabled := chkInterna.Checked;
deInterna2.Enabled := chkInterna.Checked;
end;

procedure TfrmImprimir.chkEntregaClick(Sender: TObject);
begin
deEntrega1.Enabled := chkEntrega.Checked;
deEntrega2.Enabled := chkEntrega.Checked;
end;

procedure TfrmImprimir.FormCreate(Sender: TObject);
var i : Integer;
sUser,sPassword,sServer,sDB : String;
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



deRecibido1.Date := Now;
deRecibido2.Date := Now;

deInterna1.Date := Now;
deInterna2.Date := Now;

deEntrega1.Date := Now;
deEntrega2.Date := Now;

deTerminado1.Date := DateAdd(Now,-3,daDays);
deTerminado2.Date := Now;

//GridView1.AddRow(1);


for i:=0 to 7 do
begin
        gvCampos.AddRow(1);
        gvCampos.Cell[0,gvCampos.RowCount - 1].AsString := DBNames[i];
        gvCampos.Cell[1,gvCampos.RowCount - 1].AsString := Campos[i];
        if i <= 5 then
                gvCampos.Cell[2,gvCampos.RowCount - 1].AsBoolean := True;

        gvSort.AddRow(1);
        gvSort.Cell[0,gvSort.RowCount - 1].AsString := DBNames[i];
        gvSort.Cell[1,gvSort.RowCount - 1].AsString := Campos[i];
end;






GridView1.SelectCell(0,0);
end;

procedure TfrmImprimir.Button1Click(Sender: TObject);
begin
GridView1.ClearRows;
ddlCampo.Text := '';
ddlCondicion.Text := '';
ddlValue.Text := '';
ddlConjuncion.Text := '';
ddlCampo.SetFocus;
//GridView1.AddRow(1);
//GridView1.Cell[0,0].AsString := 'Seleccione';
//GridView1.SelectCell(0,0);
end;

procedure TfrmImprimir.Button2Click(Sender: TObject);
begin
if (ddlCampo.Text = '') or (ddlCondicion.Text = '') or (ddlValue.Text = '') or (ddlConjuncion.Text = '') then
        Exit;

GridView1.AddRow(1);
GridView1.Cell[0,GridView1.RowCount - 1].AsString := ddlCampo.Text;
GridView1.Cell[1,GridView1.RowCount - 1].AsString := ddlCondicion.Text;
GridView1.Cell[2,GridView1.RowCount - 1].AsString := ddlValue.Text;
GridView1.Cell[3,GridView1.RowCount - 1].AsString := ddlConjuncion.Text;

ddlCampo.Text := '';
ddlCondicion.Text := '';
ddlValue.Text := '';
ddlConjuncion.Text := '';

GridView1.SelectCell(0,GridView1.RowCount - 1);
GridView1.SetFocus;
end;

procedure TfrmImprimir.Button3Click(Sender: TObject);
begin
          GridView1.DeleteRow(GridView1.SelectedRow);
end;

procedure TfrmImprimir.GridView1Click(Sender: TObject);
var i : Integer;
begin
  for i := 0 to GridView1.RowCount - 1 do
  begin
    if (GridView1.Cell[0, i].AsString = '') and (GridView1.RowCount > 1) and (GridView1.SelectedRow <> i ) Then
          GridView1.DeleteRow(i);
  end;

  Application.ProcessMessages;
end;

procedure TfrmImprimir.GridView1CellClick(Sender: TObject; ACol,
  ARow: Integer);
begin
  if (GridView1.Cells[0,GridView1.SelectedRow] <> '') and (GridView1.Cells[1,GridView1.SelectedRow] = '') then
          GridView1.Cells[1,GridView1.SelectedRow] := '=';

  Application.ProcessMessages;

  if (GridView1.Cells[0,GridView1.SelectedRow] <> '') and (GridView1.Cells[3,GridView1.SelectedRow] = '') then
          GridView1.Cells[3,GridView1.SelectedRow] := 'Y';


  Application.ProcessMessages;
end;

procedure TfrmImprimir.ddlCampoChange(Sender: TObject);
begin
if ddlCondicion.Text = '' Then ddlCondicion.Text := '=';
if ddlConjuncion.Text = '' Then ddlConjuncion.Text := 'AND';

BindValues(ddlCampo.Text);
end;

procedure TfrmImprimir.BindValues(Field : String);
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

      if Field = 'Producto' then
        begin
            SQLStr := 'SELECT Distinct Nombre FROM tblProductos ORDER BY Nombre';
            Field := 'Nombre';
        end
      else if Field = 'Cliente' then
        begin
            SQLStr := 'SELECT Distinct Clave FROM tblClientes ORDER BY Clave';
            Field := 'Clave';
        end
      else if Field = 'Orden' then
        begin
            SQLStr := 'SELECT DISTINCT SUBSTRING(ITE_Nombre,8,3) AS Orden ' +
                      'FROM tblItemTasks I ' +
                      'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                      'WHERE T.Nombre = ' + QuotedStr(lblTask.Caption) + ' ORDER BY SUBSTRING(ITE_Nombre,8,3)';
            Field := 'Orden';
        end
      else if Field = 'Empleado' then
        begin
            SQLStr := 'SELECT Distinct Nombre FROM tblEmpleados ORDER BY Nombre';
            Field := 'Nombre';
        end;


      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;


      ddlValue.Clear;
      while not Qry.Eof do begin
          ddlValue.Items.Add(VarToStr(Qry[Field]) );
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


procedure TfrmImprimir.Button4Click(Sender: TObject);
var Conn : TADOConnection;
Qry : TADOQuery;
SQLStr,sFields,sDBFields,sStatus,sSort,sWhere : String;
i,Count : integer;
slFields,slDBFields : TStringList;
begin

    Count := 0;
    for i:= 0 to gvCampos.RowCount - 1 do
      begin
            if gvCampos.Cell[2,i].AsBoolean = True then
                    Count := Count + 1;
      end;
    if Count <> 6 then
        begin
                MessageDlg('Por favor selecciona 6 campos de la lista.', mtInformation,[mbOk], 0);
                Exit;
        end;

    slFields := TStringList.Create;
    slDBFields := TStringList.Create;

    Qry := nil;
    Conn := nil;
    try
    begin
      Conn := TADOConnection.Create(nil);
      Conn.ConnectionString := gsConnString;
      Conn.LoginPrompt := False;
      Qry := TADOQuery.Create(nil);
      Qry.Connection :=Conn;

      SQLStr := 'SELECT ';

      for i := 0 to gvCampos.RowCount - 1 do
          begin
                  if gvCampos.Cell[2,i].AsBoolean = True then
                  begin
                    sDBFields := sDBFields + gvCampos.Cells[0,i] + ',';
                    sFields := sFields + gvCampos.Cells[1,i] + ',';
                  end;
          end;

      sDBFields := LeftStr(sDBFields,Length(sDBFields) - 1);
      sFields := LeftStr(sFields,Length(sFields) - 1);

      SQLStr := SQLStr + sDBFields + ' ' +
                'FROM tblItemTasks I ' +
                'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] ' +
                'INNER JOIN tblOrdenes O ON I.ITE_ID = O.ITE_ID ';

      if lblTask2.Caption = 'Todas' then
              SQLStr := SQLStr + 'WHERE LEFT(T.Nombre,LEN(' + QuotedStr(lblTask.Caption) + ')) = ' + QuotedStr(lblTask.Caption)
      else
              SQLStr := SQLStr + 'WHERE T.Nombre = ' + QuotedStr(lblTask2.Caption);

      if chkRecibido.Checked Then
        begin
           SQLStr := SQLStr + ' AND ( O.Recibido >= ' + QuotedStr(deRecibido1.Text) +
                     ' AND O.Recibido <= '+ QuotedStr(deRecibido2.Text + ' 23:59:59.99') + ')';
        end;

      if CheckBox1.Checked Then
          sStatus := sStatus + '1,';

      if CheckBox2.Checked Then
          sStatus := sStatus + '2,';

      if CheckBox3.Checked Then
          sStatus := sStatus + '3,';

      SQLStr := SQLStr + ' AND I.ITS_Status in (' +  LeftStr(sStatus,Length(sStatus) - 1) + ')';

      if chkInterna.Checked Then
        begin
           SQLStr := SQLStr + ' AND ( O.Interna >= ' + QuotedStr(deInterna1.Text) +
                     ' AND O.Interna <= '+ QuotedStr(deInterna2.Text + ' 23:59:59.99') + ')';
        end;

      if chkEntrega.Checked Then
        begin
           SQLStr := SQLStr + ' AND ( O.Entrega >= ' + QuotedStr(deEntrega1.Text) +
                     ' AND O.Entrega <= '+ QuotedStr(deEntrega2.Text + ' 23:59:59.99') + ')' ;
        end;

      if rbtAtrasadas.Checked Then
          SQLStr := SQLStr + ' AND dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo';


      for i := 0 to GridView1.RowCount - 1 do
          begin
                  if i = 0 then
                     sWhere := sWhere + ' AND '
                  else
                     sWhere := sWhere + ' ' + GridView1.Cells[3,i];

                  if GridView1.Cells[0,i] = 'Producto' then
                          sWhere := sWhere + ' O.Producto ';

                  if GridView1.Cells[0,i] = 'Cliente' then
                          sWhere := sWhere + ' LEFT(I.ITE_Nombre,3) ';

                  if GridView1.Cells[0,i] = 'Empleado' then
                          sWhere := sWhere + ' O.Nombre ';

                  if (GridView1.Cells[1,i] = '=') or (GridView1.Cells[1,i] = '<>') then
                          sWhere := sWhere + ' ' + GridView1.Cells[1,i] + ' ' +
                                    QuotedStr(GridView1.Cells[2,i]);

                  if (GridView1.Cells[1,i] = 'Like') or (GridView1.Cells[1,i] = 'Not Like') then
                          sWhere := sWhere + ' ' + GridView1.Cells[1,i] + ' ' +
                                    QuotedStr('%' + GridView1.Cells[2,i] + '%');
          end;

      if sWhere <> '' Then SQLStr := SQLStr + sWhere;

      for i := 0 to gvSort.RowCount - 1 do
          begin
                  if gvSort.Cell[2,i].AsBoolean = True then
                  begin
                    sSort := sSort + gvSort.Cells[0,i] + ',';
                  end;
          end;

      sSort := LeftStr(sSort,Length(sSort) - 1);
      if sSort <> '' Then SQLStr := SQLStr + ' ORDER BY ' + sSort;


      Qry.SQL.Clear;
      Qry.SQL.Text := SQLStr;
      Qry.Open;

      Application.Initialize;
      Application.CreateForm(TPrintReport, PrintReport);
      PrintReport.ReportTitle.Caption := 'Ordenes de Trbajo [' + lblTask.Caption + ']';
      //PrintReport.lblType.Caption := '';
      PrintReport.QRSubDetail1.DataSet := Qry;

      //Field1
      slFields.CommaText := sFields;
      slDBFields.CommaText := sDBFields;

      PrintReport.Field1.DataSet := Qry;
      PrintReport.Field1.DataField := RightStr(slDBFields[0],Length(slDBFields[0]) - 2);
      PrintReport.THeader1.Caption  := slFields[0];
      PrintReport.Header1.Caption  := slFields[0];

      {PrintReport.Field1.Width := StrToInt(gsWidth[0]);
      PrintReport.THeader1.Width := StrToInt(gsWidth[0]);
      PrintReport.Header1.Width := StrToInt(gsWidth[0]);
      }
      //Field2
      PrintReport.Field2.DataSet := Qry;
      PrintReport.Field2.DataField := RightStr(slDBFields[1],Length(slDBFields[1]) - 2);
      PrintReport.THeader2.Caption  := slFields[1];
      PrintReport.Header2.Caption  := slFields[1];

      {PrintReport.Field2.Left := StrToInt(gsWidth[0]) + 10;
      PrintReport.THeader2.Left := StrToInt(gsWidth[0]) + 10;
      PrintReport.THeader2.Left := StrToInt(gsWidth[0]) + 10;


      PrintReport.Field2.Width := StrToInt(gsWidth[1]);
      PrintReport.THeader2.Width := StrToInt(gsWidth[1]);
      PrintReport.Header2.Width := StrToInt(gsWidth[1]);
      }

      //Field3
      PrintReport.Field3.DataSet := Qry;
      PrintReport.Field3.DataField := RightStr(slDBFields[2],Length(slDBFields[2]) - 2);
      PrintReport.THeader3.Caption  := slFields[2];
      PrintReport.Header3.Caption  := slFields[2];

      {PrintReport.Field3.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10;
      PrintReport.THeader3.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10;
      PrintReport.THeader3.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10;


      PrintReport.Field3.Width := StrToInt(gsWidth[2]);
      PrintReport.THeader3.Width := StrToInt(gsWidth[2]);
      PrintReport.Header3.Width := StrToInt(gsWidth[2]);
      }

      //Field4
      PrintReport.Field4.DataSet := Qry;
      PrintReport.Field4.DataField := RightStr(slDBFields[3],Length(slDBFields[3]) - 2);
      PrintReport.THeader4.Caption  := slFields[3];
      PrintReport.Header4.Caption  := slFields[3];

      {PrintReport.Field4.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10;
      PrintReport.THeader4.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10;
      PrintReport.THeader4.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10;


      PrintReport.Field4.Width := StrToInt(gsWidth[3]);
      PrintReport.THeader4.Width := StrToInt(gsWidth[3]);
      PrintReport.Header4.Width := StrToInt(gsWidth[3]);
      }
      //Field5
      PrintReport.Field5.DataSet := Qry;
      PrintReport.Field5.DataField := RightStr(slDBFields[4],Length(slDBFields[4]) - 2);
      PrintReport.THeader5.Caption  := slFields[4];
      PrintReport.Header5.Caption  := slFields[4];

      {PrintReport.Field5.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10;
      PrintReport.THeader5.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10;
      PrintReport.THeader5.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10;

      PrintReport.Field5.Width := StrToInt(gsWidth[4]);
      PrintReport.THeader5.Width := StrToInt(gsWidth[4]);
      PrintReport.Header5.Width := StrToInt(gsWidth[4]);
      }
      //Field6
      PrintReport.Field6.DataSet := Qry;
      PrintReport.Field6.DataField := RightStr(slDBFields[5],Length(slDBFields[5]) - 2);
      PrintReport.THeader6.Caption  := slFields[5];
      PrintReport.Header6.Caption  := slFields[5];

      {PrintReport.Field6.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10  + StrToInt(gsWidth[4]) + 10;
      PrintReport.THeader6.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10  + StrToInt(gsWidth[4]) + 10;
      PrintReport.THeader6.Left := StrToInt(gsWidth[0]) + 10 + StrToInt(gsWidth[1]) + 10 + StrToInt(gsWidth[2]) + 10 + StrToInt(gsWidth[3]) + 10  + StrToInt(gsWidth[4]) + 10;

      PrintReport.Field6.Width := StrToInt(gsWidth[5]);
      PrintReport.THeader6.Width := StrToInt(gsWidth[5]);
      PrintReport.Header6.Width := StrToInt(gsWidth[5]);
      }
      PrintReport.Preview;
      PrintReport.Free;
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

procedure TfrmImprimir.CheckBox1Click(Sender: TObject);
begin
if (checkbox2.Checked = False) and (checkbox3.Checked = False) Then
        CheckBox1.Checked := True;

end;

procedure TfrmImprimir.CheckBox2Click(Sender: TObject);
begin
if (checkbox1.Checked = False) and (checkbox3.Checked = False) Then
        CheckBox2.Checked := True;

end;

procedure TfrmImprimir.CheckBox3Click(Sender: TObject);
begin
if (checkbox1.Checked = False) and (checkbox2.Checked = False) Then
        CheckBox3.Checked := True;

end;

procedure TfrmImprimir.Button5Click(Sender: TObject);
var SQLStr,sWhere,sWhere2:String;
i : Integer;
begin
    sWhere := '';
    sWhere2 := '';
    SQLStr := '';
    if chkRecibido.Checked Then
      begin
         SQLStr := SQLStr + ' AND ( O.Recibido >= ' + QuotedStr(deRecibido1.Text) +
                   ' AND O.Recibido <= '+ QuotedStr(deRecibido2.Text + ' 23:59:59.99') + ')';
      end;

    if chkInterna.Checked Then
      begin
         SQLStr := SQLStr + ' AND ( O.Interna >= ' + QuotedStr(deInterna1.Text) +
                   ' AND O.Interna <= '+ QuotedStr(deInterna2.Text + ' 23:59:59.99') + ')';
      end;

    if chkEntrega.Checked Then
      begin
         SQLStr := SQLStr + ' AND ( O.Entrega >= ' + QuotedStr(deEntrega1.Text) +
                   ' AND O.Entrega <= '+ QuotedStr(deEntrega2.Text + ' 23:59:59.99') + ')' ;
      end;

    if chkTerminado.Checked Then
      begin
         sWhere2 := sWhere2 + ' AND ( ITS_DTStop >= ' + QuotedStr(deTerminado1.Text) +
                   ' AND ITS_DTStop <= '+ QuotedStr(deTerminado2.Text + ' 23:59:59.99') + ')';
      end;

    if rbtAtrasadas.Checked Then
        SQLStr := SQLStr + ' AND dbo.GetHours(I.ITS_DTStart,GETDATE()) > T.Tiempo';


    for i := 0 to GridView1.RowCount - 1 do
        begin
                if i = 0 then
                   sWhere := sWhere + ' AND ('
                else
                   sWhere := sWhere + ' ' + GridView1.Cells[3,i - 1];

                if GridView1.Cells[0,i] = 'Producto' then
                        sWhere := sWhere + ' O.Producto ';

                if GridView1.Cells[0,i] = 'Cliente' then
                        sWhere := sWhere + ' SUBSTRING(I.ITE_Nombre,4,3) ';

                if GridView1.Cells[0,i] = 'Empleado' then
                        sWhere := sWhere + ' O.Nombre ';

                if GridView1.Cells[0,i] = 'Orden' then
                        sWhere := sWhere + ' SUBSTRING(I.ITE_Nombre,8,3) ';

                if (GridView1.Cells[1,i] = '=') or (GridView1.Cells[1,i] = '<>') then
                        sWhere := sWhere + ' ' + GridView1.Cells[1,i] + ' ' +
                                  QuotedStr(GridView1.Cells[2,i]);

                if (GridView1.Cells[1,i] = 'Like') or (GridView1.Cells[1,i] = 'Not Like') then
                        sWhere := sWhere + ' ' + GridView1.Cells[1,i] + ' ' +
                                  QuotedStr('%' + GridView1.Cells[2,i] + '%');
        end;

    if sWhere <> '' Then SQLStr := SQLStr + sWhere + ')';

    frmMain.lblQuery.Caption := SQLStr;
    frmMain.lblTerminado.Caption := sWhere2;
    frmMain.Timer1Timer(nil);
    self.Hide;
    frmMain.Enabled := True;
    setActiveWindow(frmMain.Handle);
end;

end.
