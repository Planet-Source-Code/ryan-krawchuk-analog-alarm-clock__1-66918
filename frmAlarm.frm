VERSION 5.00
Begin VB.Form frmAlarm 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Alarm"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   Picture         =   "frmAlarm.frx":0000
   ScaleHeight     =   1710
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optSet 
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   783
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optSet 
      Caption         =   "On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   783
      Width           =   735
   End
   Begin VB.TextBox txtMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "00"
      Top             =   730
      Width           =   495
   End
   Begin VB.TextBox txtHour 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "00"
      Top             =   730
      Width           =   495
   End
   Begin Clock.chameleonButton cbAlarm 
      Default         =   -1  'True
      Height          =   405
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAlarm.frx":1748A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   980
      TabIndex        =   5
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1290
      TabIndex        =   4
      Top             =   480
      Width           =   315
   End
   Begin VB.Label lblHour 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   390
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblSync 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   3820
      Picture         =   "frmAlarm.frx":174A6
      Top             =   20
      Width           =   315
   End
   Begin VB.Image imgC1 
      Height          =   315
      Left            =   120
      Picture         =   "frmAlarm.frx":17B3C
      Top             =   1320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgC2 
      Height          =   315
      Left            =   480
      Picture         =   "frmAlarm.frx":181D2
      Top             =   1320
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeTheForm As clsTransForm 'make a reference to the class

Private Sub cbAlarm_Click()
    If optSet(0).Value = True Then
        blnAlarm = True
        frmMain.lblAlarm.Caption = txtHour.Text & ":" & txtMin.Text
        frmMain.tmrAlarm.Enabled = True
    Else
        blnAlarm = False
        frmMain.tmrAlarm.Enabled = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.top = (Screen.Height - Me.Height) / 2
    Me.left = (Screen.Width - Me.Width) / 2
    
    Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
    ShapeTheForm.ShapeMe frmAlarm, RGB(255, 255, 0), True, App.Path & "\Alarm.dat"
    
    If blnAlarm = True Then
        optSet(0).Value = True
    Else
        optSet(1).Value = True
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgC1.Picture
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgC2.Picture
End Sub

Private Sub txtHour_Change()
    If Val(txtHour.Text) < 0 Or Val(txtHour.Text) > 23 Then
        txtHour.Text = "00"
        txtHour.SelStart = 0
        txtHour.SelLength = Len(txtHour.Text)
    End If
End Sub

Private Sub txtHour_GotFocus()
    txtHour.SelStart = 0
    txtHour.SelLength = Len(txtHour.Text)
End Sub

Private Sub txtHour_LostFocus()
    If Len(txtHour.Text) < 2 Then txtHour.Text = "0" & txtHour.Text
End Sub

Private Sub txtMin_Change()
    If Val(txtMin.Text) < 0 Or Val(txtMin.Text) > 59 Then
        txtMin.Text = "00"
        txtMin.SelStart = 0
        txtMin.SelLength = Len(txtMin.Text)
    End If
End Sub

Private Sub txtMin_GotFocus()
    txtMin.SelStart = 0
    txtMin.SelLength = Len(txtMin.Text)
End Sub

Private Sub txtMin_LostFocus()
    If Len(txtMin.Text) < 2 Then txtMin.Text = "0" & txtMin.Text
End Sub
