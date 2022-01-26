Attribute VB_Name = "PasteInVisible"
Option Explicit

Sub SaveAsAddIn()
 'Alt+F8 SaveAsAddIn Run
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
  'копирует связный (СД) или фрагментированный  скрытием, группировкой или фильтрацией диапазон (ФД) видимых ячеек
  Set rCopy = Selection.SpecialCells(xlVisible)
 Else
  Set rCopy = ActiveCell
 End If
 rCopy.Select 'выделить ЗД для вставки через Ctrl+D или Ctrl+R
 Selection.Copy 'в буфер обмена (БО)
End Sub

Function min(p As Long, c As Long) As Long
 If p < c Then
  min = p
 Else
  min = c
 End If
End Function

Sub PasteV()
 'Shift+Ctrl+V
 'только значения из ЗД вставляем в ВД
 PasteX True
End Sub

Sub PasteK()
 'Shift+Ctrl+K
 'ключи это не пустые видимые ячейки ВД и ЗД
 'если хоть один ключ ВД отличается от ключа в ЗД значит не вставляем ни чего
 'иначе только значения из ЗД вставляем в ВД
 PasteX True, True
End Sub

Sub PasteX(Optional val As Boolean = False, Optional key As Boolean = False)
 'Shift+Ctrl+X
 'ЗД вставить в ВД
 Dim rCopy As Range 'ЗД
 Dim rPaste As Range 'ВД
 Dim p As Long
 Dim r As Long
 Dim c As Long
 Dim ro As Long
 Dim co As Long
 Dim paste As Long
 Dim bPaste As Boolean
 Dim aPaste() As Boolean
 Dim aCalculation As XlCalculation
 Set rCopy = GetClipboardLink
 If rCopy Is Nothing Then Set rCopy = ActiveCell
 If rCopy.Count > 1 Then Set rCopy = rCopy.SpecialCells(xlVisible)
 aCalculation = Application.Calculation
 On Error GoTo Finally
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual
Try:
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
 ReDim aPaste(1 To rPaste.Areas.Count) As Boolean
 For paste = CLng(key) + 1 To 1
  For p = 1 To rPaste.Areas.Count
   'уменьшаем размеры копирования до размеров вставки
   With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
    If paste Then
     If val Then 'вставка только значений
      If Not key Or aPaste(p) Then
       .Copy
       'только не пустые для консолидации
       Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column).PasteSpecial paste:=xlPasteValues, SkipBlanks:=key
      End If
     Else
      .Copy Destination:= _
      Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column)
     End If
    Else 'сравниваем ключи
     ro = -Sgn(rCopy.Areas(p).Row - rPaste.Areas(p).Row) * Abs(rCopy.Areas(p).Row - rPaste.Areas(p).Row)
     co = -Sgn(rCopy.Areas(p).Column - rPaste.Areas(p).Column) * Abs(rCopy.Areas(p).Column - rPaste.Areas(p).Column)
     For r = .Row To .Row + .Rows.Count - 1
      For c = .Column To .Column + .Columns.Count - 1
       'куда вставлять Cells(r, c).Offset(ro, co) если не пусто значит ключ
       'откуда Cells(r, c) пустые пропускаем иначе сравниваем с ключом
       If IsEmpty(Cells(r, c)) Then 'пропускаем
       Else 'не пропускаемые
        If IsEmpty(Cells(r, c).Offset(ro, co)) Then 'не ключ
         bPaste = True 'вставляем вообще
         aPaste(p) = True 'вставляем в Areas(p)
        Else 'ключ
         If Cells(r, c) <> Cells(r, c).Offset(ro, co) Then 'они отличны
          MsgBox prompt:=Cells(r, c) & "@" & Cells(r, c).Address & "<>" & Cells(r, c).Offset(ro, co) & "@" & Cells(r, c).Offset(ro, co).Address, Title:="Ключи НЕ равны"
          GoTo Msg
         End If
        End If
       End If
      Next 'c
     Next 'r
     If Not bPaste Then
      MsgBox prompt:="Нечего вставлять", Title:="Ключи равны"
      GoTo Msg
     End If
    End If 'paste
   End With
  Next 'p
 Next 'paste
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
