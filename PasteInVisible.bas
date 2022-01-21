Attribute VB_Name = "PasteInVisible"
Option Explicit
Dim rCopy As Range '��

Sub SaveAsAddIn()
 'Alt+F8 SaveAsAddIn
 Dim sName As String
 sName = SplitS(0, Application.ThisWorkbook.Name, ".") '��� ����������
 On Error Resume Next
 Application.AddIns2(sName).Installed = False '������������
 On Error GoTo 0
 DoEvents
 '�������� ��� AddIn
 Application.ThisWorkbook.SaveAs Filename:=Application.UserLibraryPath & sName & ".xlam", FileFormat:=xlOpenXMLAddIn
 DoEvents
 On Error Resume Next
 Application.AddIns2(sName).Installed = True '����������
 On Error GoTo 0
End Sub

Sub WB_BeforeClose()
 '���������� �� Workbook_BeforeClose
 Application.OnKey "+^c"
 Application.OnKey "+^v"
 Application.OnKey "+^x"
End Sub

Sub WB_Open()
 '���������� �� Workbook_Open
 Application.OnKey "+^c", "SelectVisible"
 Application.OnKey "+^v", "PasteV"
 Application.OnKey "+^x", "PasteX"
End Sub

Sub SelectVisible()
 'Shift+Ctrl+C
 '������� �������� (��) Selection.Areas.Count=1
 '����������������� �������� (��) Selection.Areas.Count>1
 '���������� �������� (��) Selection
 '�� �� � ��
 If Selection.Count > 1 Then
  '������������� �� �� �� � �������� ����������������� ������������ ��� ��������� �� � ��������� ��� ��� (��)
  Set rCopy = Selection.SpecialCells(xlVisible)
 Else
  Set rCopy = ActiveCell
 End If
 rCopy.Select '�������� �� ��� ������� ����� Ctrl+D ��� Ctrl+R
 Selection.Copy '�������� ��� ����� Ctrl+C
End Sub
Function min(p As Long, c As Long) As Long
 '��� ���������� ������
 If p <= c Then
  min = p
 Else
  min = c
 End If
End Function
Sub PasteV()
 'Shift+Ctrl+V
 '������ �������� �� �� �������� � ��
 PasteX True
End Sub
Sub PasteX(Optional val As Boolean = False)
 'Shift+Ctrl+X
 '�� �������� � ��
 If rCopy Is Nothing Then
  Set rCopy = GetClipboardLink
  If rCopy Is Nothing Then
   Set rCopy = ActiveCell
  Else
   Set rCopy = rCopy.SpecialCells(xlVisible)
  End If
 End If
 Dim aCalculation As XlCalculation
 aCalculation = Application.Calculation
 On Error GoTo Finally
 Application.ScreenUpdating = False
 Application.Calculation = xlCalculationManual
Try:
 Dim rPaste As Range
 If Selection Is Nothing Then ActiveCell.Select
 If Selection.Count > 1 Then
  '������� ����������� �� - ������� ������� ���������� ��������� ��
  Set rPaste = Selection.SpecialCells(xlVisible)
 Else
  '������� ������� - ������� ������� ���������� ��������� ��
  With rCopy.Areas(rCopy.Areas.Count)
   Set rPaste = Selection.Resize(.Row + .Rows.Count - rCopy.Areas(1).Row, _
                                 .Column + .Columns.Count - rCopy.Areas(1).Column _
                                 ).SpecialCells(xlVisible)
  End With
 End If
 Dim p As Long
 For p = 1 To rPaste.Areas.Count
  With Cells(rCopy.Areas(p).Row, rCopy.Areas(p).Column).Resize(min(rCopy.Areas(p).Rows.Count, rPaste.Areas(p).Rows.Count), min(rCopy.Areas(p).Columns.Count, rPaste.Areas(p).Columns.Count))
   If val Then '������� ������ ��������
    .Copy
    Cells(rPaste.Areas(p).Row, rPaste.Areas(p).Column).PasteSpecial Paste:=xlPasteValues
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
 Set rCopy = Nothing
End Sub

Function SplitS(Index As Long, Expression As String, Optional Delimiter As String = " ", Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
 Dim aExpression() As String
 SplitS = ""
 aExpression = Split(Expression, Delimiter, Limit, Compare)
 If LBound(aExpression) <= Index And Index <= UBound(aExpression) Then SplitS = aExpression(Index)
End Function


