Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False
    
    Call Base_Inicial
    Call Base_Filtrada
    Call Base_Resultados

    Sheets("MACROS").Select
    Range("B7").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Base_Inicial()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE INICIAL").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BD - BASE INICIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BASE INICIAL").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AN4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

End Sub

Sub Base_Filtrada()

    Application.ScreenUpdating = False
    
    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE FILTRADA").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE INICIAL").Select
    Range("AN3").Select
    ActiveSheet.Range("$B$3:$AN$2000").AutoFilter Field:=39, Criteria1:="=Não", _
        Operator:=xlAnd
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE FILTRADA").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("BASE INICIAL").Select
    Application.CutCopyMode = False
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("B4").Select
    Sheets("BASE FILTRADA").Select
    Range("AO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("AO5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

End Sub

Sub Base_Resultados()

    Application.ScreenUpdating = False
    
    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE DE RESULTADOS").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE FILTRADA").Select
    Range("AP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BASE DE RESULTADOS").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("W3").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("W3:W1000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("R3").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("R3:R1000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("S3").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("S3:S1000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B4").Select
    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True

End Sub

Sub Arquivo_Envio()

    Application.ScreenUpdating = False
  
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C13").Value & " - Gestão de OS Abertas - Dados até dia " & Worksheets("MACROS").Range("C14").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets("QUADRO DE RESULTADOS").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE RESULTADOS").Select
    Range("B1:C1").Select
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=-2
    Sheets(Array("MACROS", "BD - ID.ÁREA", "BASE INATIVA", "BD - BASE INICIAL", _
        "ÁREA SUP. RMV", "BASE INICIAL", "BASE FILTRADA")).Select
    Sheets("BASE FILTRADA").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("MACROS", "BD - ID.ÁREA", "BASE INATIVA", "BD - BASE INICIAL", _
        "ÁREA SUP. RMV", "BASE INICIAL", "BASE FILTRADA", "TDs", "GRÁFICOS")).Select
    Sheets("GRÁFICOS").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("QUADRO DE RESULTADOS").Select
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
End Sub
