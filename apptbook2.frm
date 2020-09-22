VERSION 5.00
Begin VB.Form Form22 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reminder"
   ClientHeight    =   2490
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5940
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
   Icon            =   "apptbook2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2490
   ScaleWidth      =   5940
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton FlagOpt 
      Caption         =   "&Important"
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton FlagOpt 
      Caption         =   "&Routine"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton FlagOpt 
      Caption         =   "&None"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Message 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      ToolTipText     =   "Enter Text for your info. on the date you have chosen"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label DateBox 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reminder For:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancelcmd_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Open FileName For Random As #1 Len = 128
   RecNo = ChangeDate.TheDate - DateSerial(theYear, 1, 1) + 1
   Put #1, RecNo, ChangeDate
   Close #1
   
   'ClearMonth
   FillMonth
   'Form22.Hide
   Unload Me
   
End Sub

Private Sub FlagOpt_Click(Index As Integer)
    ChangeDate.Flags = Index
    If Index = 0 Then Message.Text = ""
    If Message.Visible Then Message.SetFocus
End Sub

Private Sub Message_Change()
   ChangeDate.Msg = Message.Text
   If FlagOpt(0).Value = True And Message.Text <> "" Then FlagOpt(1).Value = True
End Sub

Private Sub OKcmd_Click()
   Open FileName For Random As #1 Len = 128
   RecNo = ChangeDate.TheDate - DateSerial(theYear, 1, 1) + 1
   Put #1, RecNo, ChangeDate
   Close #1
   
   'ClearMonth
   FillMonth
   Form22.Hide
   
   

End Sub

Private Sub Option1_Click()
   ChangeDate.Flags = Index
    If Index = 0 Then Message.Text = ""
    If Message.Visible Then Message.SetFocus
End Sub

Private Sub Option2_Click()
  ChangeDate.Flags = Index
    If Index = 0 Then Message.Text = ""
    If Message.Visible Then Message.SetFocus
End Sub

Private Sub Option3_Click()
 ChangeDate.Flags = Index
    If Index = 0 Then Message.Text = ""
    If Message.Visible Then Message.SetFocus
End Sub
