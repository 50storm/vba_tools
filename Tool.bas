Attribute VB_Name = "Tool"
Option Explicit


Sub Zoom85per()
    ActiveWindow.Zoom = 85
End Sub


Sub SelectA1()
Attribute SelectA1.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim sheetNumber As Integer
    sheetNumber = ActiveWorkbook.Worksheets.Count
    Dim i As Integer

    '左から
    'For i = 1 To sheetNumber
    '
    '    Sheets(i).Select
    '    Range("A1").Select
    '
    '
    'Next i

    '右から
    For i = sheetNumber To 1 Step -1
        Sheets(i).Select
        Range("A1").Select
    Next i


End Sub


Sub SelectA1AndSave()

    Call SelectA1
    ActiveWorkbook.Save


End Sub


Sub SpellChecker()
'
' Macroスペルチェック Macro
'

 
    Dim sheet As Object
    For Each sheet In ActiveWorkbook.Sheets
        sheet.Activate
        sheet.CheckSpelling
    Next sheet
      
End Sub

Sub SpellCheckAndSelA1()
    Call SpellChecker
    Call SelectA1
End Sub


Sub SetZoom()
    Dim sheet As Object
    Dim myNum As Integer

    
    On Error Resume Next
    myNum = Trim(Application.InputBox("Enter a number"))
    
    If CStr(myNum) = "" Then
    
    Else
    
         If IsNumeric(myNum) = True Then
        
             For Each sheet In ActiveWorkbook.Sheets
                sheet.Activate
                ActiveSheet.Range("A1").Select
                ActiveWindow.Zoom = myNum
             Next sheet
     
         End If
    End If
    
    ActiveSheet.Activate
End Sub


Sub InitSheet()
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ColumnWidth = 5
    Range("A1").Select
End Sub


Sub RemoveSheetComments()
    Cells.Select
    Selection.ClearComments
    Range("A1").Select
End Sub

Sub ClearComments()
    Selection.ClearComments
End Sub


Sub ClearColor()
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub ClearColorAndCmments()
    Call ClearComments
    
    Call ClearColor
End Sub
