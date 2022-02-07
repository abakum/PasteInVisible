Attribute VB_Name = "PasteInVisible"
Option Explicit

Sub WB_BeforeClose(Optional hide_from_Macros_dialog_box As Boolean)
 'вызывается из Workbook_BeforeClose
 Application.OnKey "+^c"
 Application.OnKey "+^v"
 Application.OnKey "+^x"
 Application.OnKey "+^k"
End Sub

Sub WB_Open(Optional hide_from_Macros_dialog_box As Boolean)
 'вызывается из Workbook_Open
 Application.OnKey "+^c", "SelectVisible"
 Application.OnKey "+^v", "PasteV"
 Application.OnKey "+^x", "PasteX"
 Application.OnKey "+^k", "PasteK"
End Sub

Sub SelectVisible(Optional hide_from_Macros_dialog_box As Boolean)
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

Sub PasteV(Optional hide_from_Macros_dialog_box As Boolean)
 'Shift+Ctrl+V
 'только значения из ЗД вставляем в ВД
 PasteX True
End Sub

Sub PasteK(Optional hide_from_Macros_dialog_box As Boolean)
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
 Dim rCU As Range
 Dim rC As Range
 Dim rPaste As Range 'ВД
 Dim rPU As Range
 Dim rP As Range
 Dim p As Long
 Dim r As Long
 Dim c As Long
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
  If 1 Then 'с учётом фрагментации ЗД
   Set rPaste = Selection
   r = rPaste.Areas(1).Row - rCopy.Areas(1).Row
   c = rPaste.Areas(1).Column - rCopy.Areas(1).Column
   For p = 1 To rCopy.Areas.Count
    With rCopy.Areas(p)
     Set rPaste = Union(rPaste, Cells(.Row, .Column).Offset(r, c).Resize(.Rows.Count, .Columns.Count))
    End With
   Next 'p
  Else
   With rCopy.Areas(rCopy.Areas.Count)
    Set rPaste = Selection.Resize(.Row + .Rows.Count - rCopy.Areas(1).Row, _
                                  .Column + .Columns.Count - rCopy.Areas(1).Column)
   End With
  End If
  If rPaste.Count > 1 Then Set rPaste = rPaste.SpecialCells(xlVisible)
 End If
 For p = 1 To rPaste.Areas.Count
  'уменьшаем размеры копирования до размеров вставки
  With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(Application.min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), _
                                                               Application.min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
   If key Then 'PasteK
    'сравниваем ключи
    For r = .Row To .Row + .Rows.Count - 1
     For c = .Column To .Column + .Columns.Count - 1
      'куда вставлять если не пусто значит ключ
      Set rP = Cells(r, c).Offset(rPaste.Areas(p).Row - rCopy.Areas(p).Row, _
                                  rPaste.Areas(p).Column - rCopy.Areas(p).Column)
      Set rC = Cells(r, c) 'откуда пустые пропускаем иначе сравниваем с ключом
      If IsEmpty(rC) Then 'пропускаем
      Else 'не пропускаемые
       If IsEmpty(rP) Then 'не ключ
        If rCU Is Nothing Then
         Set rCU = rC
         Set rPU = rP
        Else
         Set rCU = Union(rCU, rC)
         Set rPU = Union(rPU, rP)
        End If
       Else 'ключ
        If rC <> rP Then 'они отличны
         MsgBox prompt:=rC & "@" & rC.Address & "<>" & rP & "@" & rP.Address, Title:="Ключи НЕ равны"
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
  If rCU Is Nothing Then GoTo KeyE
  For p = 1 To rCU.Areas.Count
   rPU.Areas(p) = rCU.Areas(p).Value
  Next
 End If
 GoTo Finally
KeyE:
 MsgBox prompt:="Нечего вставлять", Title:="Ключи равны"
KeyD:
 rCopy.Select
Finally:
 XlCalc aCalculation
End Sub

'https://stackoverflow.com/a/70890803/18055780
Sub CopyPaste(rPaste As Range, rCopy As Range, Optional val As Boolean = True)
 Dim p As Long
 Dim r As Long
 Dim c As Long
 Dim aCalculation As XlCalculation
 aCalculation = XlCalc()
 On Error GoTo Finally
Try:
 If rPaste.Count = 1 Then
  r = rPaste.Areas(1).Row - rCopy.Areas(1).Row
  c = rPaste.Areas(1).Column - rCopy.Areas(1).Column
  For p = 1 To rCopy.Areas.Count
   With rCopy.Areas(p)
    Set rPaste = Union(rPaste, Cells(.Row, .Column).Offset(r, c).Resize(.Rows.Count, .Columns.Count))
   End With
  Next 'p
 End If
 For p = 1 To rPaste.Areas.Count
  With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(Application.min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), _
                                                               Application.min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
   If val Then
    If 1 Then 'faster
     rPaste.Areas(p) = .Value
    Else
     .Copy
     Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column).PasteSpecial paste:=xlPasteValues
    End If
   Else
    .Copy Destination:= _
    Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column)
   End If 'val
  End With
 Next 'p
Finally:
 XlCalc aCalculation
End Sub

Function XlCalc(Optional aCalculation As Long = 0) As XlCalculation
 Dim bCleared As Boolean
 Dim bCutCopyMode As Boolean
 bCutCopyMode = Application.CutCopyMode
 XlCalc = Application.Calculation
 Application.EnableEvents = aCalculation <> 0
 Application.ScreenUpdating = aCalculation <> 0
 'assignment to Application.Calculation clears the clipboard
 If aCalculation = 0 Then
  bCleared = XlCalc <> xlCalculationManual
  If bCleared Then Application.Calculation = xlCalculationManual
 Else
  If aCalculation = xlCalculationAutomatic Then Application.Calculate
  bCleared = XlCalc <> aCalculation
  If bCleared Then Application.Calculation = aCalculation
 End If
 If Not bCleared Then Exit Function
 If Not bCutCopyMode Then Exit Function
 If Selection Is Nothing Then Exit Function
 Selection.Copy 'restore clipboard
End Function

'https://stackoverflow.com/a/70916088/18055780
Sub SaveAsAddIn()
 'Alt+F8 SaveAsAddIn Run
 Dim sName As String
 Dim sFilename As String
 Dim o As Object
 Application.MacroOptions _
  "SaveAsAddIn", _
  "Save ThisWorkbook as AddIn" & vbCr & _
  "Сохранить макросы в библиотеку"
 On Error GoTo Finally
Try:
 With ThisWorkbook
  sName = Split(.Name, ".")(0) 'name of ThisWorkbook without extension
  sFilename = Application.UserLibraryPath & sName & ".xlam"
  .Save
  .Worksheets.Add After:=.Worksheets(.Worksheets.Count) 'add a blank sheet at the end
  On Error Resume Next
  Application.AddIns(sName).Installed = False 'uninstall the previous version of the AddIn
  SetAttr sFilename, vbNormal
  Kill sFilename
  Application.DisplayAlerts = False
  For Each o In .Sheets  'delete all sheets except the last one
   o.Delete
  Next
  .SaveAs Filename:=sFilename, FileFormat:=xlOpenXMLAddIn 'save ThisWorkbook as AddIn
  Application.AddIns(sName).Installed = True 'install ThisWorkbook as AddIn
  .Close
  Application.DisplayAlerts = True
 End With
Finally:
End Sub
