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
 Dim rCopyK As Range
 Dim rPaste As Range 'ВД
 Dim rPasteK As Range
 Dim p As Long
 Dim r As Long
 Dim c As Long
 Dim ro As Long
 Dim co As Long
 Dim aCalculation As XlCalculation
 Set rCopy = GetClipboardLink
 If rCopy Is Nothing Then Set rCopy = ActiveCell
 If rCopy.Count > 1 Then Set rCopy = rCopy.SpecialCells(xlVisible)
 On Error GoTo Finally
 aCalculation = XlCalc()
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
 For p = 1 To rPaste.Areas.Count
  'уменьшаем размеры копирования до размеров вставки
  With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
   If key Then 'PasteK
    'сравниваем ключи
    ro = -Sgn(rCopy.Areas(p).Row - rPaste.Areas(p).Row) * Abs(rCopy.Areas(p).Row - rPaste.Areas(p).Row)
    co = -Sgn(rCopy.Areas(p).Column - rPaste.Areas(p).Column) * Abs(rCopy.Areas(p).Column - rPaste.Areas(p).Column)
    For r = .Row To .Row + .Rows.Count - 1
     For c = .Column To .Column + .Columns.Count - 1
      'куда вставлять Cells(r, c).Offset(ro, co) если не пусто значит ключ
      'откуда Cells(r, c) пустые пропускаем иначе сравниваем с ключом
      If IsEmpty(Cells(r, c)) Then 'пропускаем
      Else 'не пропускаемые
       If IsEmpty(Cells(r, c).Offset(ro, co)) Then 'не ключ
        If rCopyK Is Nothing Then
         Set rCopyK = Cells(r, c)
        Else
         Set rCopyK = Union(rCopyK, Cells(r, c))
        End If
        If rPasteK Is Nothing Then
         Set rPasteK = Cells(r, c).Offset(ro, co)
        Else
         Set rPasteK = Union(rPasteK, Cells(r, c).Offset(ro, co))
        End If
       Else 'ключ
        If Cells(r, c) <> Cells(r, c).Offset(ro, co) Then 'они отличны
         MsgBox prompt:=Cells(r, c) & "@" & Cells(r, c).Address & "<>" & Cells(r, c).Offset(ro, co) & "@" & Cells(r, c).Offset(ro, co).Address, Title:="Ключи НЕ равны"
         GoTo KeyD
        End If
       End If
      End If
     Next 'c
    Next 'r
   Else 'PasteX PasteV
    If val Then 'вставка только значений
     If 1 Then 'быстрей
      rPaste.Areas(p) = .Value
     Else
      .Copy
      Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column).PasteSpecial paste:=xlPasteValues
     End If
    Else
     .Copy Destination:= _
     Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column)
    End If
   End If 'key
  End With
 Next 'p
 If key Then
  If rCopyK Is Nothing Then GoTo KeyE
  For p = 1 To rCopyK.Areas.Count
   rPasteK.Areas(p) = rCopyK.Areas(p).Value
  Next
 End If
 GoTo Finally
KeyE:
 MsgBox prompt:="Нечего вставлять", Title:="Ключи равны"
KeyD:
 XlCalc aCalculation
 rCopy.Select
 Selection.Copy
 Exit Sub
Finally:
 XlCalc aCalculation
End Sub

Private Function XlCalc(Optional aCalculation As Long = 0) As XlCalculation
 Application.EnableEvents = aCalculation <> 0
 Application.ScreenUpdating = aCalculation <> 0
 If aCalculation = 0 Then
  XlCalc = Application.Calculation
  Application.Calculation = xlCalculationManual
 Else
  Application.Calculation = aCalculation
 End If
End Function

Function SplitS(Index As Long, Expression As String, Optional Delimiter As String = " ", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
 Dim aExpression() As String
 SplitS = ""
 aExpression = Split(Expression, Delimiter, Limit, Compare)
 If LBound(aExpression) <= Index And Index <= UBound(aExpression) Then SplitS = aExpression(Index)
End Function
