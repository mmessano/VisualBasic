'-----------------------------
'Macros to populate a spreadsheet with table and column names
Function DSN()
'Choose one of these formats and set DSN equal to the correct data source
'Need to set: {DSN=|SERVER=};UID=;PWD=;DATABASE=;
'With Pre-Defined ODBC DSN
DSN1 = "ODBC;DSN=Great Plains KK;UID=DYNSA;PWD=0129;DATABASE=TEST;"
'Direct DSN Definition
DSN2 = "ODBC;DRIVER=SQL Server;SERVER=SQLSRV;UID=sa;PWD=0129;DATABASE=TEST"
'Select the one to use
DSN = DSN2
End Function
Sub Tables()
    Q = ""
    Q = Q & " SELECT DISTINCT"
    Q = Q & "  [sysobjects].[name] AS Table_Name"
    'Q = Q & " , [systypes].[name] as Xtype"
    'Q = Q & " , [syscolumns].[xtype]"
    'Q = Q & " , [syscolumns].[xusertype]"
    'Q = Q & " , [syscolumns].[length]"
    'Q = Q & " , [cdefault]"
    'Q = Q & " , [syscolumns].[domain]"
    'Q = Q & " , [syscolumns].[collationid]"
    'Q = Q & " , [syscolumns].[type]"
    'Q = Q & " , [syscolumns].[usertype]"
    'Q = Q & " , [syscolumns].[prec]"
    'Q = Q & " , [syscolumns].[scale]"
    'Q = Q & " , [iscomputed]"
    'Q = Q & " , [isoutparam]"
    'Q = Q & " , [isnullable]"
    Q = Q & " FROM [syscolumns], [sysobjects], [systypes]"
    Q = Q & " where [syscolumns].[ID] = [sysobjects].[ID]"
    Q = Q & " and [sysobjects].[xtype] in ('V', 'U')"
    Q = Q & " and [syscolumns].[xtype]=[systypes].[xtype]"
    'Q = Q & " and [syscolumns].[id]> 100"
    'Q = Q & " order by  ParentObj,[syscolumns].[name]"

   doquery Q, "Tables"
End Sub
Sub Columns()
    Q = ""
    Q = Q & " SELECT [syscolumns].[name]"
    Q = Q & " , [sysobjects].[name] AS ParentObj"
    Q = Q & " , [systypes].[name] as Xtype"
    Q = Q & " , [syscolumns].[xtype]"
    Q = Q & " , [syscolumns].[xusertype]"
    Q = Q & " , [syscolumns].[length]"
    'Q = Q & " , [cdefault]"
    'Q = Q & " , [syscolumns].[domain]"
    'Q = Q & " , [syscolumns].[collationid]"
    Q = Q & " , [syscolumns].[type]"
    Q = Q & " , [syscolumns].[usertype]"
    Q = Q & " , [syscolumns].[prec]"
    Q = Q & " , [syscolumns].[scale]"
    'Q = Q & " , [iscomputed]"
    'Q = Q & " , [isoutparam]"
    'Q = Q & " , [isnullable]"
    Q = Q & " FROM [syscolumns], [sysobjects], [systypes]"
    Q = Q & " where [syscolumns].[ID] = [sysobjects].[ID]"
    Q = Q & " and [sysobjects].[xtype] in ('V', 'U')"
    Q = Q & " and [syscolumns].[xtype]=[systypes].[xtype]"
    'Q = Q & " and [syscolumns].[id]> 100"
    Q = Q & " order by  ParentObj,[syscolumns].[name]"

   doquery Q, "Columns"
End Sub
Sub Objects()
    Q = ""
    Q = Q & " SELECT [A].[name]"
    'Q = Q & ", [A].[id]"
    Q = Q & " , CASE [A].[xtype]"
    Q = Q & "   WHEN 'C'  Then '4 Constraint'"
    Q = Q & "   WHEN 'D'  Then '5 Default'"
    Q = Q & "   WHEN 'F'  Then '6 Foreign'"
    Q = Q & "   WHEN 'L'  Then '9 Log'"
    Q = Q & "   WHEN 'FN' Then '8 Scaler Function'"
    Q = Q & "   WHEN 'IF' Then '9 Inline Table'"
    Q = Q & "   WHEN 'P'  Then '7 Procedure'"
    Q = Q & "   WHEN 'PK' Then '1 Primary Key'"
    Q = Q & "   WHEN 'RF' Then '9 Filter'"
    Q = Q & "   WHEN 'S'  Then '9 System Table'"
    Q = Q & "   WHEN 'TF' Then '8 Table Function'"
    Q = Q & "   WHEN 'TR' Then '7 Trigger'"
    Q = Q & "   WHEN 'U'  Then '0 User Table'"
    Q = Q & "   WHEN 'UQ' Then '2 Unique'"
    Q = Q & "   WHEN 'V'  Then '3 View'"
    Q = Q & "   WHEN 'X'  Then 'Extended Proc'"
    Q = Q & "   Else '9 Unknown'"
    Q = Q & "  END As Xtype"
   'Q = Q & " ,[A].[parent_obj], [B].[name] as Parent, [A].[type]"
    Q = Q & " ,[B].[name] as Parent, [A].[type]"
    Q = Q & " FROM [sysobjects] AS A LEFT JOIN  [sysobjects] AS B"
    Q = Q & vbCrLf
    Q = Q & " ON [A].[parent_obj]=B.[id]"
    'Q = Q & " where [Xtype] in  ('U','V','P')"
    'Q = Q & " WHERE [A].[parent_obj]=B.[id]"
    Q = Q & " order by [A].[Xtype],[A].[name]"

   doquery Q, "Objects"
   
End Sub
Sub Procs()
    Q = ""
    Q = Q & " SELECT [syscolumns].[name]"
    Q = Q & " , [sysobjects].[name] AS Parent"
    Q = Q & " , [systypes].[name] as Xtype"
    Q = Q & " , [syscolumns].[xtype]"
    Q = Q & " , [syscolumns].[xusertype]"
    Q = Q & " , [syscolumns].[length]"
    Q = Q & " , [cdefault]"
    Q = Q & " , [syscolumns].[domain]"
    Q = Q & " , [syscolumns].[collationid]"
    Q = Q & " , [syscolumns].[type]"
    Q = Q & " , [syscolumns].[usertype]"
    Q = Q & " , [syscolumns].[prec]"
    Q = Q & " , [syscolumns].[scale]"
    Q = Q & " , [iscomputed]"
    Q = Q & " , [isoutparam]"
    Q = Q & " , [isnullable]"
    Q = Q & " FROM [syscolumns], [sysobjects], [systypes]"
    Q = Q & " where [syscolumns].[ID] = [sysobjects].[ID]"
    Q = Q & " and [sysobjects].[xtype] in ('P','X','IF','FN')"
    Q = Q & " and [syscolumns].[xtype]=[systypes].[xtype]"
    'Q = Q & " and [syscolumns].[id]> 100"
    Q = Q & " order by Parent,[sysobjects].[xtype]"
   doquery Q, "ProcsPrameters"
End Sub

Sub doquery(Q, N)
    X = ActiveSheet.Cells(1, 1)
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    ActiveSheet.Cells(1, 1) = X
    With ActiveSheet.QueryTables.Add(Connection:=DSN, Destination:=Range("A2"))
        .CommandText = Q
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = True
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.Name = N
End Sub