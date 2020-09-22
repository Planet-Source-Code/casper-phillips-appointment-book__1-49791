Attribute VB_Name = "Module2"
'Attribute VB_Name = "Module2"
Global Const asModel = 1
Global Const BLACK = &H0&
'Global Const RED = &HFF&     YOU CAN REMOVE THE (') If YOU WANT TO ADD COLORS TO YOU DATES
'Global Const GREEN = &HFF00&
'Global Const YELLOW = &HFFFF&
'Global Const BLUE = &HFF0000
'Global Const MAGENTA = &HFF00FF
'Global Const CYAN = &HFFFF00
'Global Const WHITE = &HFFFFFF
'Global Const LT_GRAY = &H808080
'Global Const DK_GRAY = &H404040

Type DayType
  TheDate As Double
  Flags As Integer
  Msg As String * 118
End Type

Global FileName As String
Global theYear As Integer
Global theMonth As Integer
Global EOL As String
Global ChangeDate As DayType
Global Months(1 To 12) As String
Global thisMonth(36) As Double
Global OldBackColor As Double
Global OldForeColor As Double


Sub ClearMonth()
  Dim Index As Integer
  For Index = 0 To 36
     apptmain.DayButton(Index).Caption = ""
     apptmain.DayButton(Index).Enabled = False
     apptmain.DayButton(Index).BackColor = &HFFFFFF
     apptmain.DayButton(Index).ForeColor = &H80000008
  Next Index

End Sub

Sub FillMonth()
   Dim StartDate As Double, aDay As DayType
   Dim theDay As Integer, RecNo As Integer
   Dim Index As Integer, I As Integer, FirstIndex As Integer
   ClearMonth

   apptmain.monthbox.Caption = Months(theMonth) + Str$(theYear)
   For I = 1 To 12
     apptmain.monthcmd(I).Checked = False
   Next I

   apptmain.monthcmd(theMonth).Checked = True
   StartDate = DateSerial(theYear, theMonth, 1)
   Open FileName For Random As #1 Len = 128
    RecNo = StartDate - DateSerial(theYear, 1, 1) + 1
    Get #1, RecNo, aDay
    theDay = Day(aDay.TheDate)
    FirstIndex = Weekday(StartDate) - 1
    Index = FirstIndex
    
    While theMonth = Month(StartDate + Index - FirstIndex)
       apptmain.DayButton(Index).Caption = Str$(theDay) + EOL + aDay.Msg
       apptmain.DayButton(Index).Enabled = True
       thisMonth(Index) = aDay.TheDate
       If aDay.TheDate = Fix(Now) Then
         apptmain.DayButton(Index).Caption = "&*" + Str$(theDay)
         'Form1.DayButton(Index).Caption = ""
         'apptmain.DayButton(Index).BackColor = YELLOW
       End If
       
       If aDay.Flags = 1 Then
         'apptmain.DayButton(Index).BackColor = CYAN
       ElseIf aDay.Flags = 2 Then
        ' apptmain.DayButton(Index).BackColor = RED
         'Form1.DayButton(Index).ForeColor = WHITE
       End If

       Index = Index + 1
       RecNo = RecNo + 1
       Get #1, RecNo, aDay
       theDay = Day(aDay.TheDate)
    Wend
    Close #1

End Sub

 Sub LookAhead(Days As Integer)
    Dim StartDate As Double, aDay As DayType
    Dim RecNo As Integer, FoundIt As Integer
    Dim I As Integer
    StartDate = Fix(Now)
    Open FileName For Random As #1 Len = 128
    RecNo = StartDate - DateSerial(theYear, 1, 1) + 1
    FoundIt = False
    For I = 1 To Days
      Get #1, RecNo, aDay
      If aDay.Flags = 2 Then
        FoundIt = True
        Beep
      End If
      RecNo = RecNo + 1
    Next I
    Close #1
    If FoundIt Then MsgBox "You Have Important dates this week"

End Sub

