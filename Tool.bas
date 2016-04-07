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
On Error GoTo closeError
    Call SelectA1
    If ActiveWorkbook.Saved = False Then
        ActiveWorkbook.Save
    Else
    End If
closeError:
    Exit Sub
    
End Sub


Sub SelectA1SaveAndClose()
On Error GoTo closeError
    Call SelectA1AndSave
    ActiveWorkbook.Close
closeError:
    Exit Sub
    
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
'    Dim strSheetName As String
'    strSheetName = ActiveSheet.Name

'    Dim index As Integer
'    index = ActiveSheet.index
'
  
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
    
'    Worksheets(index).Active
'    ActiveSheet.Activate
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

Sub CombineAndAlignLeft()
Attribute CombineAndAlignLeft.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' CombineAndAlignLeft Macro
'

'
'    Range("I15:Y15").Select
    With Selection
'        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub

Sub SquareSLine()
Attribute SquareSLine.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' SquareSLine Macro
'

'
'    Range("Z7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub SaveAndClose()
On Error GoTo closeError
    ActiveWorkbook.Save
    ActiveWorkbook.Close
closeError:
Exit Sub

End Sub
    
Sub InitScrollBar()
    Call SelectA1

    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub
