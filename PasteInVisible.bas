Attribute VB_Name = "PasteInVisible"
Option Explicit

Sub SaveAsAddIn()
 'Alt+F8 SaveAsAddIn
 Dim sName As String
 sName = SplitS(0, Application.ThisWorkbook.Name, ".") 'без расширения
 On Error Resume Next
 Application.AddIns2(sName).Installed = False 'деинсталирую
 On Error GoTo 0
 DoEvents
 'сохраняю как AddIn
 Application.ThisWorkbook.SaveAs Filename:=Application.UserLibraryPath & sName & ".xlam", FileFormat:=xlOpenXMLAddIn
 DoEvents
 On Error Resume Next
 Application.AddIns2(sName).Installed = True 'инсталирую
 On Error GoTo 0
End Sub

Sub WB_BeforeClose()
 'вызывается из Workbook_BeforeClose
 Application.OnKey "+^c"
 Application.OnKey "+^v"
 Application.OnKey "+^x"
 Application.OnKey "+^k"
End Sub

Sub WB_Open()
 'вызывается из Workbook_Open
 Application.OnKey "+^c", "SelectVisible"
 Application.OnKey "+^v", "PasteV"
 Application.OnKey "+^x", "PasteX"
 Application.OnKey "+^k", "PasteK"
End Sub

Sub SelectVisible()
 'Shift+Ctrl+C
 'связный диапазон (СД) Selection.Areas.Count=1
 'фрагментированный диапазон (ФД) Selection.Areas.Count>1
 'выделенный диапазон (ВД) Selection
 'из ВД в ЗД
 Dim rCopy As Range 'ЗД
 If Selection.Count > 1 Then
  'преобразовать ВД из СД в возможно фрагментированный группировкой или фильтрами ФД и запомнить его как (ЗД)
  Set rCopy = Selection.SpecialCells(xlVisible)
 Else
  Set rCopy = ActiveCell
 End If
 rCopy.Select 'выделить ЗД для вставки через Ctrl+D или Ctrl+R
 Selection.Copy 'пометить как после Ctrl+C
End Sub
Function min(p As Long, c As Long) As Long
 'без расширения границ
 If p <= c Then
  min = p
 Else
  min = c
 End If
End Function
Sub PasteV()
 'Shift+Ctrl+V
 'только значения из ЗД вставить в ВД
 PasteX True
End Sub
Sub PasteX(Optional val As Boolean = False)
 'Shift+Ctrl+X
 'ЗД вставить в ВД
 Dim rCopy As Range 'ЗД
 Set rCopy = GetClipboardLink
 If rCopy Is Nothing Then Set rCopy = ActiveCell
 If rCopy.Count > 1 Then Set rCopy = rCopy.SpecialCells(xlVisible)
 Dim aCalculation As XlCalculation
 aCalculation = Application.Calculation
 On Error GoTo Finally
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual
Try:
 Dim rPaste As Range
 If Selection Is Nothing Then ActiveCell.Select
 If Selection.Count > 1 Then
  'вставка ограниченна ВД - границы вставки определены размерами ВД
  Set rPaste = Selection.SpecialCells(xlVisible)
 Else
  'обычная вставка - границы вставки определены размерами ЗД
  With rCopy.Areas(rCopy.Areas.Count)
   Set rPaste = Selection.Resize(.Row + .Rows.Count - rCopy.Areas(1).Row, _
                                 .Column + .Columns.Count - rCopy.Areas(1).Column _
                                 )
   If rPaste.Count > 1 Then Set rPaste = rPaste.SpecialCells(xlVisible)
  End With
 End If
 Dim p As Long
 For p = 1 To rPaste.Areas.Count
  With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
   If val Then 'вставка только значений
    .Copy
    Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column).PasteSpecial paste:=xlPasteValues
   Else
    .Copy Destination:= _
    Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column)
   End If
  End With
 Next
Finally:
 Application.ScreenUpdating = True
 Application.Calculation = aCalculation
 Application.CutCopyMode = False
End Sub
Sub PasteK()
 'Shift+Ctrl+K
 'сравнить ключевые(не пустые) видимые поля при совпадении вставить в пустые видимые поля
 Dim rCopy As Range 'ЗД
 Set rCopy = GetClipboardLink
 If rCopy Is Nothing Then Set rCopy = ActiveCell
 If rCopy.Count > 1 Then Set rCopy = rCopy.SpecialCells(xlVisible)
 Dim aCalculation As XlCalculation
 aCalculation = Application.Calculation
 On Error GoTo Finally
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual
Try:
 Dim rPaste As Range
 If Selection Is Nothing Then ActiveCell.Select
 If Selection.Count > 1 Then
  'вставка ограниченна ВД - границы вставки определены размерами ВД
  Set rPaste = Selection.SpecialCells(xlVisible)
 Else
  'обычная вставка - границы вставки определены размерами ЗД
  With rCopy.Areas(rCopy.Areas.Count)
   Set rPaste = Selection.Resize(.Row + .Rows.Count - rCopy.Areas(1).Row, _
                                 .Column + .Columns.Count - rCopy.Areas(1).Column _
                                 )
   If rPaste.Count > 1 Then Set rPaste = rPaste.SpecialCells(xlVisible)
  End With
 End If
 Dim p As Long
 Dim r As Long
 Dim c As Long
 Dim ro As Long
 Dim co As Long
 Dim paste As Long
 Dim doPaste As Boolean
 doPaste = False
 Dim notE2E As Boolean
 For paste = 0 To 1
  For p = 1 To rPaste.Areas.Count
   ro = -Sgn(rCopy.Areas(p).Row - rPaste.Areas(p).Row) * Abs(rCopy.Areas(p).Row - rPaste.Areas(p).Row)
   co = -Sgn(rCopy.Areas(p).Column - rPaste.Areas(p).Column) * Abs(rCopy.Areas(p).Column - rPaste.Areas(p).Column)
   With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
    For r = .Row To .Row + .Rows.Count - 1
     For c = .Column To .Column + .Columns.Count - 1
      notE2E = IsEmpty(Cells(r, c).Offset(ro, co)) And Not IsEmpty(Cells(r, c))
      If paste Then
       If notE2E Then
        'Debug.Print Cells(r, c) & "@" & Cells(r, c).Address & "->" & Cells(r, c).Offset(ro, co) & "@" & Cells(r, c).Offset(ro, co).Address
        Cells(r, c).Offset(ro, co) = Cells(r, c)
       End If
      Else
       If notE2E Then
        doPaste = True
       Else
        If Not IsEmpty(Cells(r, c)) And Cells(r, c) <> Cells(r, c).Offset(ro, co) Then
         MsgBox prompt:=Cells(r, c) & "@" & Cells(r, c).Address & "<>" & Cells(r, c).Offset(ro, co) & "@" & Cells(r, c).Offset(ro, co).Address, Title:="Ключи НЕ равны"
         GoTo Msg
        End If
       End If
      End If
     Next
    Next
   End With
  Next
  If Not doPaste Then
   MsgBox prompt:="Нечего вставлять", Title:="Ключи равны"
   GoTo Msg
  End If
 Next
 GoTo Finally
Msg:
 Application.ScreenUpdating = True
 Application.Calculation = aCalculation
 rCopy.Select
 Selection.Copy
 Exit Sub
Finally:
 Application.ScreenUpdating = True
 Application.Calculation = aCalculation
 Application.CutCopyMode = False
End Sub
Function SplitS(Index As Long, Expression As String, Optional Delimiter As String = " ", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
 Dim aExpression() As String
 SplitS = ""
 aExpression = Split(Expression, Delimiter, Limit, Compare)
 If LBound(aExpression) <= Index And Index <= UBound(aExpression) Then SplitS = aExpression(Index)
End Function
