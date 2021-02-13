VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Alarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Clock v1.0"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Alarm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Play forever"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton config 
      Caption         =   "Configure Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Set a personalized Message"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Current File to be played"
      Top             =   1920
      Width           =   2775
      Begin VB.Label wavname 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Current File to be played"
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chose a WAV file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "Chose a file to be played on Alarm activation"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable Alarm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Activate the Alarm"
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Current Time"
      Top             =   120
      Width           =   1695
   End
   Begin MSComCtl2.UpDown ud2 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox mk2 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Set the ALARM minutes"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   3
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mk1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Set the ALARM hour"
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   3
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown ud1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox min 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   "_"
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "Alarm.frx":0442
      Top             =   720
      Width           =   480
   End
   Begin MediaPlayerCtl.MediaPlayer mm 
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -200
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label6 
      Caption         =   "Example: Play file for 008 times"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Play file"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "times"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Set Minutes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Set Hour:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Menu aboutme 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hour As String
Dim minute As String
Dim total As String
Public flag As Integer
Dim a As Integer
Public MyFileName As String
Dim i As Long
Public forever As Integer


Private Sub aboutme_Click()
Call About.Show
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    forever = 1
    min.Enabled = False
    min.BackColor = RGB(170, 170, 170)
ElseIf Check1.Value = 0 Then
    forever = 0
    min.BackColor = RGB(255, 255, 255)
    min.Enabled = True
End If
End Sub

Private Sub Command1_Click()
cd1.Filter = "WAV (*.wav)|*.wav"
Call cd1.ShowOpen
MyFileName = cd1.FileName
mm.FileName = cd1.FileName
wavname.Caption = cd1.FileName
If wavname.Caption <> "" Then
    Command2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Disable Alarm" Then
    Command2.Caption = "Enable Alarm"
    flag = 0
    wavname.Caption = ""
    MyFileName = ""
    Command2.Enabled = False
    Image1.Picture = LoadPicture(App.Path & "\LIGHTOFF.ico")
    Exit Sub
ElseIf Command2.Caption = "Enable Alarm" Then
    Command2.Caption = "Disable Alarm"
    Image1.Picture = LoadPicture(App.Path & "\LIGHTON.ico")
End If
hour = mk1.Text
minute = mk2.Text
If Val(hour) < 10 And Val(hour) >= 0 Then
    hour = "0" & hour
End If
If Val(minute) < 10 And Val(minute) >= 0 Then
    minute = "0" & minute
End If
total = hour & ":" & minute
flag = 1

If Check1 = 1 Then
    forever = 1
ElseIf Check1 = 0 Then
    forever = 0
    mm.PlayCount = Val(min.Text)
End If
If MyFileName = "" Then
    MsgBox ("Please select a WAV file for your alarm!")
    flag = 0
End If
End Sub

Private Sub config_Click()
Call frmConfig.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Call About.Show
End If
End Sub

Private Sub Form_Load()
Text1.Text = Format$(Time, "hh:mm:ss")
forever = 0
a = 0
Load wakeup

With ud1
    .BuddyControl = mk1
    .Wrap = True
    .min = 0
    .Max = 23
    .Value = 0
End With

With ud2
    .BuddyControl = mk2
    .Wrap = True
    .min = 0
    .Max = 60
    .Value = 0
End With

wavname.Caption = ""
Command2.Enabled = False
flag = 0
min.Text = "999"
End Sub



Private Sub Form_Resize()
If WindowState = 1 Then
    Caption = Format(Time, "Medium Time")
Else
    Caption = "Alarm Clock"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmConfig
Unload wakeup
Unload Alarm
End Sub


Private Sub Timer1_Timer()
If WindowState = vbMinimized Then
    Caption = Format(Time, "Medium Time")
End If
Text1.Text = Format$(Time, "hh:mm:ss")
If flag = 1 Then
    If total = Mid$(Text1.Text, 1, 5) Then
        Call wakeup.Show
        If forever = 1 Then
            mm.PlayCount = 0
            mm.Play
        ElseIf forever = 0 Then
            Call mm.Play
        End If
        Command2.Enabled = False
        Command1.Enabled = False
        Call wakeup.Show
        flag = 0
    End If
End If
End Sub

Private Sub ud1_Change()
mk1.Text = ud1.Value
End Sub

Private Sub ud2_Change()
mk2.Text = ud2.Value
End Sub
