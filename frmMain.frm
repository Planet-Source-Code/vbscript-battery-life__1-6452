VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTimeLeft 
      Caption         =   "Time left on Battery"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close Program"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2055
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   1843
            MinWidth        =   1764
            TextSave        =   "8:56 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   1843
            MinWidth        =   1764
            TextSave        =   "3/7/2000"
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Is your Caps-Lock on?"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Is your Num-Lock on?"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Is your Scroll-Lock on?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
            Object.ToolTipText     =   "Insert or Overwrite text?"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBattery 
      Caption         =   "Battery Life"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin MSComctlLib.ProgressBar prgBattery 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblStatusString 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lblFull 
         Alignment       =   1  'Right Justify
         Caption         =   "100 %"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblEmpty 
         Caption         =   "Empty"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl50 
         Alignment       =   2  'Center
         Caption         =   "50 %"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3975
      End
   End
   Begin SysInfoLib.SysInfo sysBattery 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrBattery 
      Interval        =   5000
      Left            =   1920
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'   This sets up the first window caption.
    frmMain.Caption = "Battery Remaining: Checking"
End Sub

Private Sub tmrBattery_Timer()
'   calls the functions to update the display.
    Progress
    PercentageLeft
    ACStatus
    
'   The following two lines change the window caption and the status
    frmMain.Caption = "Battery Remaining: " & prgBattery.Value & "%"
    lblStatus.Caption = prgBattery.Value & "% battery power left."
End Sub

Private Function Progress()
'   This function updates the progress bar.
'   Tests first for a value, then sets the progress bar value to
'   that value and enables the progress bar.  255 is status unknown
    If sysBattery.BatteryLifePercent <> 255 Then
        prgBattery.Value = sysBattery.BatteryLifePercent
        prgBattery.Enabled = True
    Else
        prgBattery.Value = 0
        prgBattery.Enabled = False
    End If
End Function
Private Function BatteryStatus()
'   Checks the current status for the battery.  this sets the caption
'   for the  lable and the color of the text based on battery state.
    Select Case sysBattery.BatteryStatus
        Case 1 ' battery is OK
            lblStatusString.Caption = "Battery OK"
            lblStatusString.ForeColor = &H8000&
        Case 2 ' Battery getting low
            lblStatusString.Caption = "Battery Low"
            lblStatusString.ForeColor = &HFF00FF
        Case 4 ' Um, you need to plug in
            lblStatusString.Caption = "Battery Critical"
            lblStatusString.ForeColor = &HFF&
        Case 8 ' Battery charging
            lblStatusString.Caption = "Battery Charging"
            lblStatusString.ForeColor = &HFF0000
        Case 128, 255 ' Cannot get status
            lblStatusString.Caption = "No Battery Status"
            lblStatusString.ForeColor = &H0&
    End Select
End Function

Private Function ACStatus()
'   This checks to see if the laptop is on AC or battery power and
'   update caption and text color as necessary.
    Select Case sysBattery.ACStatus
        Case 0 ' Signifies that unit is on battery power
            BatteryStatus
        Case 1 ' Signifies that unit is on AC power
            lblStatusString.Caption = "Unit is on AC Power"
            lblStatusString.ForeColor = &H8000&
        Case 255 ' Cannot get status.
            lblStatusString.Caption = "AC Status Unknown"
            lblStatusString.ForeColor = &HFF&
    End Select
End Function

Private Function PercentageLeft()
'   This sets the color of the caption text based on the amount
'   of charge left on the battery.
    If prgBattery.Value < 75 And prgBattery.Value > 50 Then
        lblStatus.ForeColor = &H8000&
    ElseIf prgBattery.Value < 50 And prgBattery.Value > 25 Then
        lblStatus.ForeColor = &HFF00FF
    ElseIf prgBattery.Value < 25 Then
        lblStatus.ForeColor = &HFF&
    Else
        lblStatus.ForeColor = &HFF0000
    End If
End Function

Private Sub cmdTimeLeft_Click()
'   this gets the time left (hrs/mins) on the battery.
'   I have not been able to see this work, I am taking
'   for granted that it does.
    If sysBattery.BatteryLifeTime <> &HFFFFFFFF Then
        Dim TimeLeft As String
        Dim TimeTotal As String
        Dim temp
        temp = TimeSerial(0, 0, sysBattery.BatteryLifeTime)
        TimeLeft = Format(temp, "h:mm")
        temp = TimeSerial(0, 0, sysBattery.BatteryFullTime)
        TimeTotal = Format(temp, "h:mm")
        MsgBox TimeLeft & " time left from " & TimeTotal & " total."
    Else
        MsgBox "Cannot determine the time remaining on the battery.", _
            vbInformation, "Time Left on Battery"
    End If
End Sub

Private Sub cmdExit_Click()
'   Ends program...
    End
End Sub

