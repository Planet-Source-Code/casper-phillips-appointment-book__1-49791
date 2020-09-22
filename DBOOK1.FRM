VERSION 5.00
Begin VB.Form apptmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00D0DDDF&
   BorderStyle     =   0  'None
   Caption         =   "JELS Appointment Book"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   1140
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "DBOOK1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Rightarrow 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   10320
      Picture         =   "DBOOK1.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   49
      ToolTipText     =   "Next Month"
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox Leftarrow 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1920
      Picture         =   "DBOOK1.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   48
      ToolTipText     =   "Previous Month"
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   9120
      TabIndex        =   43
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   5
      Left            =   7320
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   4
      Left            =   10920
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   47
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   46
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   45
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   44
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label CurTimeBox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   42
      Top             =   8520
      Width           =   3615
   End
   Begin VB.Label CurDateBox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   41
      Top             =   8520
      Width           =   3615
   End
   Begin VB.Label CurWeekDayBox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   40
      Top             =   8520
      Width           =   3615
   End
   Begin VB.Label monthbox 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Appointment Book"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   39
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   36
      Left            =   2040
      TabIndex        =   38
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   35
      Left            =   240
      TabIndex        =   37
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   34
      Left            =   10920
      TabIndex        =   36
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   33
      Left            =   9120
      TabIndex        =   35
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   32
      Left            =   7320
      TabIndex        =   34
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   31
      Left            =   5520
      TabIndex        =   33
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   30
      Left            =   3720
      TabIndex        =   32
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   29
      Left            =   2040
      TabIndex        =   31
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   28
      Left            =   240
      TabIndex        =   30
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   27
      Left            =   10920
      TabIndex        =   29
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   26
      Left            =   9120
      TabIndex        =   28
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   25
      Left            =   7320
      TabIndex        =   27
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   24
      Left            =   5520
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   23
      Left            =   3720
      TabIndex        =   25
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   22
      Left            =   2040
      TabIndex        =   24
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   21
      Left            =   240
      TabIndex        =   23
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   20
      Left            =   10920
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   19
      Left            =   9120
      TabIndex        =   21
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   18
      Left            =   7320
      TabIndex        =   20
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   17
      Left            =   5520
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   16
      Left            =   3720
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   15
      Left            =   2040
      TabIndex        =   17
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   14
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   13
      Left            =   10920
      TabIndex        =   15
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   12
      Left            =   9120
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   11
      Left            =   7320
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   10
      Left            =   5520
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   9
      Left            =   3720
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   8
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   6
      Left            =   10920
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   5
      Left            =   9120
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   4
      Left            =   7320
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   2
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Day"
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu Newcmd 
         Caption         =   "&New"
      End
      Begin VB.Menu OpenCmd 
         Caption         =   "&Open"
      End
      Begin VB.Menu SepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu Exitcmd 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Monthmenu 
      Caption         =   "&Month"
      Begin VB.Menu monthcmd 
         Caption         =   "&Jan"
         Index           =   1
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Feb"
         Index           =   2
      End
      Begin VB.Menu monthcmd 
         Caption         =   "Ma&r"
         Index           =   3
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Apr"
         Index           =   4
      End
      Begin VB.Menu monthcmd 
         Caption         =   "Ma&y"
         Index           =   5
      End
      Begin VB.Menu monthcmd 
         Caption         =   "J&un"
         Index           =   6
      End
      Begin VB.Menu monthcmd 
         Caption         =   "Ju&l"
         Index           =   7
      End
      Begin VB.Menu monthcmd 
         Caption         =   "Au&g"
         Index           =   8
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Sep"
         Index           =   9
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Oct"
         Index           =   10
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Nov"
         Index           =   11
      End
      Begin VB.Menu monthcmd 
         Caption         =   "&Dec"
         Index           =   12
      End
   End
End
Attribute VB_Name = "apptmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DayButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   OldBackColor = DayButton(Index).BackColor
   OldForeColor = DayButton(Index).ForeColor
  ' DayButton(Index).BackColor = BLACK
  ' DayButton(Index).ForeColor = BLACK

End Sub

Private Sub DayButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   'ClearMonth
   DayButton(Index).BackColor = OldBackColor
   DayButton(Index).ForeColor = OldForeColor

   Open FileName For Random As #1 Len = 128
     RecNo = thisMonth(Index) - DateSerial(theYear, 1, 1) + 1
     Get #1, RecNo, ChangeDate
   Close #1

    Form22.DateBox.Caption = Format$(ChangeDate.TheDate, "mm-dd-yyyy")
    Form22.Message.Text = ChangeDate.Msg
    Form22.FlagOpt(ChangeDate.Flags).Value = True
    Form22.Show 1
End Sub

Private Sub Exitcmd_Click()
   End
  
End Sub

Private Sub FillMonth()
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
         apptmain.DayButton(Index).Caption = "*" + Str$(theDay)
        ' Form1.DayButton(Index).Caption = ""
        ' apptmain.DayButton(Index).BackColor = YELLOW (you can Remove this if you want the current day to be highlited in yellow)
       End If
       
       If aDay.Flags = 1 Then
         'apptmain.DayButton(Index).BackColor = CYAN (You can remove this if youwant your routine reminders to be highlited in cyan)
       ElseIf aDay.Flags = 2 Then
         'apptmain.DayButton(Index).BackColor = RED (You can also remove this if you want your important dates highlited in red)
         'Form1.DayButton(Index).ForeColor = WHITE
       End If

       Index = Index + 1
       RecNo = RecNo + 1
       Get #1, RecNo, aDay
       theDay = Day(aDay.TheDate)
    Wend
    Close #1


End Sub

Private Sub Form_Load()
    'IF You change the code or improve it, Please send me a copy to (Casper7007@hotmail.com). Thank you.
    'To Create a new year just Click on New from the File Menu, When the Input Box Pops up Just Type in the Year That you what to create a calendar for.
    'The whole calendar year well be created for. HAVE FUN
   Dim aDay As DayType
   EOL = Chr$(13) + Chr$(10)
   Months(1) = "January"
   Months(2) = "February"
   Months(3) = "March"
   Months(4) = "April"
   Months(5) = "May"
   Months(6) = "June"
   Months(7) = "July"
   Months(8) = "August"
   Months(9) = "September"
   Months(10) = "October"
   Months(11) = "November"
   Months(12) = "December"
   
   ClearMonth

   FileName = Format$(Year(Now), "0000") + ".DBK"
   If Dir$(FileName) = "" Then Exit Sub
   Open FileName For Random As #1 Len = 128
   Get #1, 1, aDay
   theYear = Year(aDay.TheDate)
   Close #1
   theMonth = Month(Now)
   FillMonth
   LookAhead (7)
   Load Form22
   Form22.Hide
End Sub

Private Sub leftarrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Leftarrow.ForeColor = BLACK
End Sub

Private Sub leftarrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ClearMonth
  Leftarrow.ForeColor = WHITE
  theMonth = theMonth - 1
  If theMonth = 0 Then theMonth = 12
  FillMonth
End Sub

Private Sub LookAhead(Days As Integer)
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

Private Sub MonthCmd_Click(Index As Integer)
   theMonth = Index
   FillMonth
End Sub

Private Sub Newcmd_Click()
   Dim aDay As DayType, I As Integer, aYear As String
   Dim aMsg As String, Action As Integer

   theYear = 0
   If theYear < 1753 Or theYear > 2078 Then
     aYear = InputBox$("Type the year")
     If aYear = "" Then Exit Sub
     theYear = Val(aYear)
   End If

   FileName = Format$(theYear, "0000") + ".DBK"
   If Dir$(FileName) <> "" Then
     aMsg = "File already exists" + EOL + " Overwrite it ?"
     Action = MsgBox(aMsg, 257)
     If Action = 2 Then Exit Sub
   End If

   aDay.TheDate = DateSerial(theYear, 1, 1) - 1
   aDay.Flags = 0
   aDay.Msg = Space$(118)

   Open FileName For Random As #1 Len = 128
   For I = 1 To 366
     aDay.TheDate = aDay.TheDate + 1
     Put #1, I, aDay
   Next I
   Close #1

   If theYear = Year(Now) Then
     theMonth = Month(Now)
   Else
     theMonth = 1
   End If
   FillMonth
End Sub

Private Sub OpenCmd_Click()
   Dim aDay As DayType, I As Integer, aYear As String
   Dim OldFileName As String, OldYear As Integer

   OldYear = theYear
   OldFileName = FileName
   theYear = 0
   If theYear < 1753 Or theYear > 2078 Then
     aYear = InputBox$("Type the year")
     If aYear = "" Then Exit Sub
     theYear = Val(aYear)
   End If

   FileName = Format$(theYear, "0000") + ".DBK"
   If Dir$(FileName) = "" Then
     MsgBox "File doesn't exists, Use New"
     FileName = OldFileName
     theYear = OldYear
     Exit Sub
   End If


   Open FileName For Random As #1 Len = 128
      Get #1, 1, aDay
      theYear = Year(aDay.TheDate)
   Close #1

   If theYear = Year(Now) Then
     theMonth = Month(Now)
   Else
     theMonth = 1
   End If
   FillMonth

End Sub

Private Sub rightarrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Rightarrow.ForeColor = BLACK
End Sub

Private Sub rightarrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ClearMonth
  Rightarrow.ForeColor = WHITE
  theMonth = theMonth + 1
  If theMonth = 13 Then theMonth = 1
  FillMonth
End Sub

Private Sub Timer1_Timer()
    CurTimeBox.Caption = Format$(Now, "h:mm:ss AM/PM")
   ' If PrevDay <> Format$(Now, "mmmm d, yyyy") Then
       apptmain.CurDateBox.Caption = Format$(Now, "mmmm d, yyyy")
        apptmain.CurWeekDayBox.Caption = Format$(Now, "dddd")
        PrevDay = Format$(Now, "mmmm d, yyyy")

End Sub

