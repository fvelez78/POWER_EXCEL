Sub Oculta_hoja()
'
' Oculta_hoja Macro
'
    ActiveSheet.Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub Muestra_hoja()
'
' Muestra todas las hojas ocultas
'
    Application.ScreenUpdating = False
    For Each n In Sheets
        n.Visible = True
    Next n
    Sheets(1).Activate
    Application.ScreenUpdating = True
    

End Sub
Sub GRABA_TXT()
'
' GRABA_TXT Macro
' Macro grabada el 20/01/2010 por cofvelez
'
' Esta macro graba una hoja como archivo txt.
'
    ActiveSheet.Select
    ActiveSheet.Copy
    Application.Dialogs(xlDialogSaveAs).Show
    ActiveWorkbook.Save
    ActiveWindow.Close
End Sub


Sub VEC()
'
' VEC Macro
' Macro grabada el 29/04/2015 por cofvelez
'
' Esta macro implementa el operador VEC sobre un rango rectangular completo
'
   Dim Nf0, Nf, Nc, Nc0 As Integer
   Dim F, C, TD, Tdatos As Long
   Dim M As String
        
   Application.ScreenUpdating = False
   
    Nf0 = ActiveCell.Row
    Nc0 = ActiveCell.Column
    Selection.End(xlToRight).Select
    Nc = ActiveCell.Column
    C = Nc - Nc0 + 1
    ActiveCell.Offset(0, 1 - C).Select
    Selection.End(xlDown).Select
    Nf = ActiveCell.Row
    F = Nf - Nf0 + 1
    Tdatos = C * (F - 1) ' Calcula el total de datos asumiendo que hay una dila de titulos
    ActiveCell.Offset(1 - F, 0).Select
    TD = ActiveCell.Range(Cells(Nf0 + 1, Nc0), Cells(Nf, Nc)).Count
    
    'M = MsgBox("El rango no debe tener celdas vacias, si es asi continue", vbOKOnly, "Confirmación")
    If Tdatos = TD Then
        ActiveCell.Offset(1, 0).Select
        For i = 1 To C - 1
            ActiveCell.Offset(0, 1).Select
            Range(Selection, Selection.End(xlDown)).Select
            If i < C - 1 Then
                Range(Selection, Selection.End(xlToRight)).Select
            End If
            Selection.Cut
            ActiveCell.Offset(0, -1).Select
            Selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
            ActiveSheet.Paste
            ActiveCell.Select
        Next
    'Else
     '   M = MsgBox("El rango tiene celdas vacias, completelas y vuelva a intentarlo", vbCritical, "Confirmación")
    End If
    Selection.End(xlUp).Select
    Application.ScreenUpdating = True
End Sub

Sub CBLOQUE()
'
' Esta macro copia debajo un bloque previamente seleccionado, las veces que se indique
'

    Dim n As Integer
    Application.ScreenUpdating = False
    n = InputBox("Cuantas veces desea copiar el bloque seleccionado?", "Numero de bloques", 1)
    Selection.Copy
    For i = 1 To n
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveSheet.Paste
        ActiveCell.Select
    Next
    Selection.End(xlUp).Select
    Application.ScreenUpdating = True
End Sub

Function CONCATENAR_MASIVO(Rango As Range, delimitador As String)
Dim i As Integer, Resultado As String

Resultado = ""
    For i = 1 To Rango.Count
        Resultado = Resultado + Rango(i) & delimitador
    Next
  
CONCATENAR_MASIVO = Trim(Left(Resultado, Len(Resultado) - 1)) 'quito el último delimitador

End Function


Sub ProtegerHojas()

    'Macro para proteger todas las hojas de un fichero
    Dim wksht As Worksheet
    Dim Contraseña As String
    Contraseña = InputBox("Escriba la contraseña con la cual se protegerán las hojas", "Contraseña", 1)
    For Each wksht In ActiveWorkbook.Worksheets
    wksht.Protect Password:=Contraseña
    Next wksht
M = MsgBox("Las hojaas se han protegido", vbOKOnly, "Confirmación")
End Sub


Sub VEC_n()
'
' VEC Macro
' Macro grabada el 29/04/2015 por cofvelez
'
' Esta macro implementa el operador VEC sobre un rango rectangular completo
'
   Dim Nf0, Nf, Nc, Nc0, n As Integer
   Dim F, C, TD, Tdatos As Long
   Dim M As String
        
   Application.ScreenUpdating = False
    n = InputBox("Número de columnas a agrupar antes de apilar?", "Numero de columnas", 1)
    Nf0 = ActiveCell.Row
    Nc0 = ActiveCell.Column
    Selection.End(xlToRight).Select
    Nc = ActiveCell.Column
    C = Nc - Nc0 + 1
    ActiveCell.Offset(0, 1 - C).Select
    Selection.End(xlDown).Select
    Nf = ActiveCell.Row
    F = Nf - Nf0 + 1
    Tdatos = C * (F - 1) ' Calcula el total de datos asumiendo que hay una fila de titulos
    ActiveCell.Offset(1 - F, 0).Select
    TD = ActiveCell.Range(Cells(Nf0 + 1, Nc0), Cells(Nf, Nc)).Count
    
    'M = MsgBox("El rango no debe tener celdas vacias, si es asi continue", vbOKOnly, "Confirmación")
    If Tdatos = TD Then
        ActiveCell.Offset(1, 0).Select
        For i = 1 To (C - n) Step n
            ActiveCell.Offset(0, n).Select
            Range(Selection, Selection.End(xlDown)).Select
            If i < C - 1 Then
                Range(Selection, Selection.End(xlToRight)).Select
            End If
            Selection.Cut
            ActiveCell.Offset(0, -n).Select
            Selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
            ActiveSheet.Paste
            ActiveCell.Select
        Next
    'Else
     '   M = MsgBox("El rango tiene celdas vacias, completelas y vuelva a intentarlo", vbCritical, "Confirmación")
    End If
    Selection.End(xlUp).Select
    Application.ScreenUpdating = True
End Sub

Sub CP_Valores()
'
' CP_Valores Macro
'

Application.ScreenUpdating = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    
    
End Sub
Sub F_Limpio()
'
' F_Limpio Macro
'

Application.ScreenUpdating = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = False
    Selection.Font.Italic = False
    Selection.Font.Underline = False
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.Select
    Application.ScreenUpdating = True
End Sub