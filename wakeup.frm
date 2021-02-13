VERSION 5.00
Begin VB.Form wakeup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WAKE UP!"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox screen 
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "okay, okay.. I'm up!"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   4080
      Top             =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Personalized settings:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "WAKE UP!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "wakeup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer
Dim i As Long
Dim check As Integer

Private Sub Command1_Click()
Alarm.flag = 0
Alarm.Command1.Enabled = True
Alarm.Command2.Caption = "Enable Alarm"
If Alarm.forever = 1 Then
    Alarm.mm.AutoRewind = False
    Alarm.forever = 0
End If
Alarm.mm.PlayCount = 0
Alarm.mm.FileName = ""
Alarm.mm.Stop
Alarm.wavname.Caption = ""
Alarm.Command2.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
a = -1
screen.Text = frmConfig.DisplayText.Text
check = 0
End Sub

Private Sub Timer2_Timer()
a = a * -1
If a = -1 Then
    Label1.Caption = ""
Else
    Label1.Caption = "WAKE UP!!!"
End If
End Sub
