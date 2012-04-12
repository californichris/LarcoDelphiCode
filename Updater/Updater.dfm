object Form1: TForm1
  Left = 192
  Top = 135
  BorderStyle = bsSingle
  Caption = 'Updater'
  ClientHeight = 63
  ClientWidth = 391
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  WindowState = wsMaximized
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 96
    Top = 16
    Width = 148
    Height = 13
    Caption = 'Presiona el boton de actualizar.'
  end
  object Button1: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Actualizar'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 8
    Top = 40
    Width = 1001
    Height = 625
    Lines.Strings = (
      
        'CREATE    Function FormatZeros(@NUMBER INT,@LEN AS INT) RETURNS ' +
        'VARCHAR(10)'
      'AS'
      'BEGIN'
      '   DECLARE @ZEROS AS VARCHAR(10)'
      '   DECLARE @STRNUM AS VARCHAR(10)'
      ' '#9
      '   SET @ZEROS = '#39'000000000'#39#9
      '   SELECT @STRNUM = CAST(@NUMBER AS VARCHAR(10))'
      ' '#9
      #9
      '   RETURN(SELECT LEFT( @ZEROS,@LEN - LEN(@STRNUM) ) + @STRNUM  )'
      'END')
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object Button2: TButton
    Left = 168
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 2
    Visible = False
    OnClick = Button2Click
  end
  object Edit1: TEdit
    Left = 248
    Top = 8
    Width = 121
    Height = 21
    TabOrder = 3
    Text = 'Edit1'
    Visible = False
    OnKeyDown = Edit1KeyDown
  end
  object Memo2: TMemo
    Left = 24
    Top = 40
    Width = 1001
    Height = 625
    Lines.Strings = (
      
        'CREATE  FUNCTION GetAggregateValue(@PRODUCT_ID AS INT, @TAS_ORDE' +
        'R AS INT) RETURNS DECIMAL(18,2)'
      'AS '
      'BEGIN'
      ''
      'DECLARE @AGGREGATEVALUE AS DECIMAL(18,2)'
      'DECLARE @VALUE AS DECIMAL(18,2)'
      ''
      'SELECT @VALUE = A.Value'
      'FROM tblAggregateValue A'
      'INNER JOIN tblTareas T ON A.Task_ID = T.[Id]'
      'WHERE A.Product_ID = @PRODUCT_ID AND TAS_Order = @TAS_ORDER'
      ''
      ''
      'SELECT @AGGREGATEVALUE = SUM(A.Value)'
      'FROM tblAggregateValue A'
      'INNER JOIN tblTareas T ON A.Task_ID = T.[Id]'
      'WHERE A.Product_ID = @PRODUCT_ID AND TAS_Order < @TAS_ORDER'
      ''
      ''
      'RETURN'#9'@VALUE - @AGGREGATEVALUE'
      ''
      'END')
    ReadOnly = True
    TabOrder = 4
  end
  object Memo3: TMemo
    Left = 48
    Top = 40
    Width = 1001
    Height = 625
    Lines.Strings = (
      'CREATE   PROCEDURE [dbo].[Productividad_Empleado_Dinero]'
      '    @WHERE AS VARCHAR(8000)'
      'AS'
      ''
      'DECLARE @SQL AS VARCHAR(8000)'
      'DECLARE @COUNT AS INTEGER'
      ''
      
        'IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE id = OBJECT_ID' +
        '(N'#39'tempdb..#tmpProdEmp'#39'))'
      'DROP TABLE #tmpProdEmp'
      ''
      'CREATE TABLE [#tmpProdEmp] ('
      
        #9'[Empleado] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS ' +
        'NULL ,'
      #9'[Amount] DECIMAL(18,2),'
      
        #9'[ColumnName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_A' +
        'S NULL'
      ') ON [PRIMARY]'
      ''
      'SET @SQL = '#39'INSERT INTO #tmpProdEmp '#39' + '
      
        #9'   '#39'SELECT E.Nombre, COUNT(I.ITE_ID) AS Ordenes, DBO.FormatZero' +
        's(T.TAS_Order, 2) + '#39#39'Ord.'#39#39' + T.Nombre  AS Tarea '#39' + '
      #9'   '#39'FROM tblItemTasks I '#39' +'
      '           '#39'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] '#39' + '
      
        '           '#39'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nomb' +
        're '#39' +'
      #9'   '#39'LEFT OUTER JOIN tblEmpleados E ON I.USE_Login = E.[ID] '#39' + '
      #9'   '#39'WHERE '#39
      ''
      ''
      
        'PRINT(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Order' +
        #39')'
      
        'EXECUTE(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Ord' +
        'er'#39')'
      ''
      ''
      'SET @SQL = '#39'INSERT INTO #tmpProdEmp '#39' + '
      
        #9'   '#39'SELECT E.Nombre, SUM(O.Requerida) AS Piezas, DBO.FormatZero' +
        's(T.TAS_Order, 2) + '#39#39'Pie.'#39#39' + T.Nombre AS Tarea '#39' + '
      #9'   '#39'FROM tblItemTasks I '#39' +'
      '           '#39'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] '#39' + '
      
        '           '#39'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nomb' +
        're '#39' +'
      #9'   '#39'LEFT OUTER JOIN tblEmpleados E ON I.USE_Login = E.[ID] '#39' + '
      #9'   '#39'WHERE '#39
      ''
      ''
      
        'PRINT(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Order' +
        #39')'
      
        'EXECUTE(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Ord' +
        'er'#39')'
      ''
      ''
      'SET @SQL = '#39'INSERT INTO #tmpProdEmp '#39' + '
      
        #9'   '#39'SELECT E.Nombre, SUM(O.Requerida * O.Unitario * dbo.GetAggr' +
        'egateValue(P.Id, T.TAS_Order)) AS Piezas, DBO.FormatZeros(T.TAS_' +
        'Order, 2) + '#39#39'Din.'#39#39' + T.Nombre AS Tarea '#39' + '
      #9'   '#39'FROM tblItemTasks I '#39' +'
      '           '#39'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] '#39' + '
      
        '           '#39'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nomb' +
        're '#39' +'
      #9'   '#39'INNER JOIN tblProductos P ON O.Producto = P.Nombre '#39' +'
      #9'   '#39'LEFT OUTER JOIN tblEmpleados E ON I.USE_Login = E.[ID] '#39' + '
      #9'   '#39'WHERE '#39
      ''
      ''
      
        'PRINT(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Order' +
        #39')'
      
        'EXECUTE(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Ord' +
        'er'#39')'
      ''
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT '#39'Total'#39', SUM(Amount), ColumnName '
      'FROM #tmpProdEmp'
      'GROUP BY ColumnName'
      ''
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT Empleado, SUM(Amount), '#39'97Total Ord.'#39' '
      'FROM #tmpProdEmp'
      'WHERE SUBSTRING(ColumnName,3,4) = '#39'Ord.'#39
      'GROUP BY Empleado'
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT Empleado, SUM(Amount), '#39'98Total Pie.'#39' '
      'FROM #tmpProdEmp'
      'WHERE SUBSTRING(ColumnName,3,4) = '#39'Pie.'#39
      'GROUP BY Empleado'
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT Empleado, SUM(Amount), '#39'99Total Din.'#39' '
      'FROM #tmpProdEmp'
      'WHERE SUBSTRING(ColumnName,3,4) = '#39'Din.'#39
      'GROUP BY Empleado'
      ''
      '--SELECT * FROM #tmpProdEmp'
      ''
      'SELECT @COUNT = COUNT(AMOUNT) FROM #tmpProdEmp'
      ''
      'IF(@COUNT <> 0) BEGIN'
      
        #9'EXEC CROSSTAB_ORIGINAL '#39'SELECT Empleado FROM #tmpProdEmp GROUP ' +
        'BY Empleado ORDER BY Empleado'#39','#39'SUM(Amount)'#39','#39'ColumnName'#39','#39'#tmpP' +
        'rodEmp'#39','#39' ORDER BY ColumnName'#39
      'END'
      'ELSE BEGIN'
      
        #9'SELECT '#39'No hay Ordenes, para el filtro que selecciono.'#39' AS Mens' +
        'aje'
      'END'
      ''
      'DROP TABLE #tmpProdEmp')
    ReadOnly = True
    TabOrder = 5
  end
  object Memo4: TMemo
    Left = 80
    Top = 40
    Width = 969
    Height = 625
    Lines.Strings = (
      'ALTER   PROCEDURE [dbo].[Productividad_Empleado]'
      '    @WHERE AS VARCHAR(8000)'
      'AS'
      ''
      'DECLARE @SQL AS VARCHAR(8000)'
      'DECLARE @COUNT AS INTEGER'
      ''
      
        'IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE id = OBJECT_ID' +
        '(N'#39'tempdb..#tmpProdEmp'#39'))'
      'DROP TABLE #tmpProdEmp'
      ''
      'CREATE TABLE [#tmpProdEmp] ('
      
        #9'[Empleado] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS ' +
        'NULL ,'
      #9'[Amount] DECIMAL(18,2),'
      
        #9'[ColumnName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_A' +
        'S NULL'
      ') ON [PRIMARY]'
      ''
      'SET @SQL = '#39'INSERT INTO #tmpProdEmp '#39' + '
      
        #9'   '#39'SELECT E.Nombre, COUNT(I.ITE_ID) AS Ordenes, DBO.FormatZero' +
        's(T.TAS_Order, 2) + '#39#39'Ord.'#39#39' + T.Nombre  AS Tarea '#39' + '
      #9'   '#39'FROM tblItemTasks I '#39' +'
      '           '#39'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] '#39' + '
      
        '           '#39'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nomb' +
        're '#39' +'
      #9'   '#39'LEFT OUTER JOIN tblEmpleados E ON I.USE_Login = E.[ID] '#39' + '
      #9'   '#39'WHERE '#39
      ''
      ''
      
        'PRINT(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Order' +
        #39')'
      
        'EXECUTE(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Ord' +
        'er'#39')'
      ''
      ''
      'SET @SQL = '#39'INSERT INTO #tmpProdEmp '#39' + '
      
        #9'   '#39'SELECT E.Nombre, SUM(O.Requerida) AS Piezas, DBO.FormatZero' +
        's(T.TAS_Order, 2) + '#39#39'Pie.'#39#39' + T.Nombre AS Tarea '#39' + '
      #9'   '#39'FROM tblItemTasks I '#39' +'
      '           '#39'INNER JOIN tblTareas T ON I.TAS_ID = T.[ID] '#39' + '
      
        '           '#39'INNER JOIN tblOrdenes O ON I.ITE_Nombre = O.ITE_Nomb' +
        're '#39' +'
      #9'   '#39'LEFT OUTER JOIN tblEmpleados E ON I.USE_Login = E.[ID] '#39' + '
      #9'   '#39'WHERE '#39
      ''
      ''
      
        'PRINT(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Order' +
        #39')'
      
        'EXECUTE(@SQL + @WHERE + '#39' GROUP BY E.Nombre, T.Nombre, T.TAS_Ord' +
        'er'#39')'
      ''
      ''
      '--SELECT * FROM #tmpProdEmp ORDER BY Empleado, ColumnName'
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT '#39'Total'#39', SUM(Amount), ColumnName '
      'FROM #tmpProdEmp'
      'GROUP BY ColumnName'
      ''
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT Empleado, SUM(Amount), '#39'98Total Ord.'#39' '
      'FROM #tmpProdEmp'
      'WHERE SUBSTRING(ColumnName,3,4) = '#39'Ord.'#39
      'GROUP BY Empleado'
      ''
      'INSERT INTO #tmpProdEmp'
      'SELECT Empleado, SUM(Amount), '#39'99Total Pie.'#39' '
      'FROM #tmpProdEmp'
      'WHERE SUBSTRING(ColumnName,3,4) = '#39'Pie.'#39
      'GROUP BY Empleado'
      '--SELECT * FROM #tmpProdEmp'
      ''
      'SELECT @COUNT = COUNT(AMOUNT) FROM #tmpProdEmp'
      ''
      'IF(@COUNT <> 0) BEGIN'
      
        #9'EXEC CROSSTAB_ORIGINAL '#39'SELECT Empleado FROM #tmpProdEmp GROUP ' +
        'BY Empleado ORDER BY Empleado'#39','#39'SUM(Amount)'#39','#39'ColumnName'#39','#39'#tmpP' +
        'rodEmp'#39','#39' ORDER BY ColumnName'#39
      'END'
      'ELSE BEGIN'
      
        #9'SELECT '#39'No hay Ordenes, para el filtro que selecciono.'#39' AS Mens' +
        'aje'
      'END'
      ''
      'DROP TABLE #tmpProdEmp')
    ReadOnly = True
    TabOrder = 6
  end
end
