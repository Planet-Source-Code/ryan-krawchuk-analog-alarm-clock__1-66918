VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Clock"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0C06
   ScaleHeight     =   3720
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAlarm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pctTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1350
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   2040
      Width           =   1035
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Bauhaus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   650
         TabIndex        =   5
         Top             =   0
         Width           =   60
      End
      Begin VB.Label lblSec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Bauhaus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   270
      End
      Begin VB.Label lblMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Bauhaus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Bauhaus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   290
         TabIndex        =   2
         Top             =   0
         Width           =   60
      End
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Bauhaus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.Timer tmrMain 
      Interval        =   500
      Left            =   3240
      Top             =   0
   End
   Begin AgentObjectsCtl.Agent msAgent 
      Left            =   0
      Top             =   3240
   End
   Begin VB.Label lblAlarm 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   1372
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   195
      Left            =   1372
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1080
      Picture         =   "frmMain.frx":2DD08
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgAlarm 
      Height          =   240
      Left            =   2400
      Picture         =   "frmMain.frx":2DE52
      Top             =   2047
      Width           =   195
   End
   Begin VB.Shape shMid 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1740
      Shape           =   3  'Circle
      Top             =   1740
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   1080
      X2              =   1860
      Y1              =   2640
      Y2              =   1860
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   6
      X1              =   960
      X2              =   1860
      Y1              =   1200
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      X1              =   2580
      X2              =   1860
      Y1              =   1440
      Y2              =   1860
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ShapeTheForm As clsTransForm 'make a reference to the class
Dim XPos As Single
Dim YPos As Single

Dim Hr As Single
Dim realHr As Single
Dim Min As Single
Dim realMin As Single
Dim Sec As Single
Dim LastHr As Single
Dim LastMin As Single
Dim LastSec As Single

Private Const PI = 3.14159265358979

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Enum CharSize
    Small = 1
    Regular = 2
    Large = 3
End Enum
Dim Size As CharSize
Private Const SW_SHOWNORMAL = 1
Public MyAgent As IAgentCtlCharacterEx
Dim Announce As Long

Private Sub Form_Load()
    Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
    'Load the region data from a file -- much faster.  The file was '
    'first produced by setting the 'True' below to 'False' during   '
    'development.  From then on loading the data from the file is   '
    'much faster than calculating the region each time you run      '
    'your application.                                              '
    ShapeTheForm.ShapeMe frmMain, RGB(255, 255, 0), True, App.Path & "\ClockDat.dat"
    
    Me.left = Screen.Width - Me.Width
    Me.top = Screen.Height - Me.Height - 570
    DoEvents
    ModifyHands
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu frmMenu.mnuFile
    Else
        FormDrag Me
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu frmMenu.mnuOpt
End Sub

Private Sub imgAlarm_Click()
    frmAlarm.Show
End Sub

Private Sub tmrAlarm_Timer()
    Dim sndplay As String
    If blnAlarm = True Then
        If lblTime.Caption = lblAlarm.Caption Then
            sndplay = sndPlaySound(App.Path & "\DIGITAL.WAV", 1)
            blnAlarm = False
            tmrAlarm.Enabled = False
            DoEvents
            'Play Alarm Dude
            Call AlarmDude
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub tmrMain_Timer()
    Call ModifyHands
    lblTime.Caption = Format(Now, "hh:mm")
    lblHour.Caption = Format(Now, "hh")
    lblMin.Caption = Format(Now, "Nn")
    lblSec.Caption = Format(Now, "ss")
    
    If blnAlarm = True Then
        shMid.BackColor = vbRed
    Else
        shMid.BackColor = &H808080
    End If
End Sub

Private Sub ModifyHands()
    If Hr <> Hour(Now) Then Hr = Hour(Now)
    If Min <> Minute(Now) Then Min = Minute(Now)
    If Sec <> Second(Now) Then Sec = Second(Now)
    
    If realHr <> Hr + Min / 60 Then realHr = Hr + Min / 60
    If realMin <> Min + Sec / 60 Then realMin = Min + Sec / 60
    
    If LastSec <> Sec Then
        If Sec > -1 And Sec < 61 Then
            Line3.X1 = 110 * Cos(PI / 180 * (6 * Sec - 90)) + Line3.X2
            Line3.Y1 = 110 * Sin(PI / 180 * (6 * Sec - 90)) + Line3.Y2
        End If
    End If
    
    If LastMin <> realMin Then
        If realMin > -1 And realMin < 61 Then
            Line1.X1 = 100 * Cos(PI / 180 * (6 * realMin - 90)) + Line1.X2
            Line1.Y1 = 100 * Sin(PI / 180 * (6 * realMin - 90)) + Line1.Y2
        End If
    End If
    
    If LastHr <> realHr Then
        If realHr > -1 And realHr < 25 Then
            Line2.X1 = 70 * Cos(PI / 180 * (30 * realHr - 90)) + Line2.X2
            Line2.Y1 = 70 * Sin(PI / 180 * (30 * realHr - 90)) + Line2.Y2
        End If
    End If
    
    If LastHr <> realHr Then LastHr = realHr
    If LastMin <> realMin Then LastMin = realMin
    If LastSec <> Sec Then LastSec = Sec
End Sub

Public Sub AlarmDude()
    'Start the Agent
    On Error Resume Next
    Dim Id As String
    Id = "Agent"
    'load default character
    msAgent.Characters.Load (Id)
    Set MyAgent = msAgent.Characters.Item(Id)
    'Set Size
    MyAgent.Height = MyAgent.OriginalHeight
    MyAgent.Width = MyAgent.OriginalWidth
    With MyAgent.Commands
        .RemoveAll
        .Add "AdvCharOptions", "&Advanced Character Options"
    End With
    MyAgent.Height = MyAgent.OriginalHeight
    MyAgent.Width = MyAgent.OriginalWidth
    MyAgent.MoveTo ((Me.Width / 15) - MyAgent.Width) / 2 + (Me.left / 15), _
                   ((Me.Height / 15) - MyAgent.Height) / 2 + (Me.top / 15)
    MyAgent.Show
    MyAgent.Play "Listen"
    MyAgent.Speak "Master, your alarm has sounded."
    MyAgent.Play "RestPose"
    MyAgent.Hide
End Sub
