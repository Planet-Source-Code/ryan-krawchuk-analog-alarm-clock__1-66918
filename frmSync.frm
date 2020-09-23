VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSync 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Synchronization"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSync.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Clock.chameleonButton cmdCheck 
      Height          =   315
      Left            =   2970
      TabIndex        =   11
      Top             =   480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Check"
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
      MICON           =   "frmSync.frx":2DF76
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtDifference 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "txtDifference"
      ToolTipText     =   "Difference between Local time & Server time."
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtLocalTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "txtLocalTime"
      ToolTipText     =   "Local time."
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtServerTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "txtServerTime"
      ToolTipText     =   "Server time in UTC/GMT"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin Clock.chameleonButton cmdAdjust 
      Default         =   -1  'True
      Height          =   285
      Left            =   2820
      TabIndex        =   2
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Sync Now"
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
      MICON           =   "frmSync.frx":2DF92
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
      Caption         =   "Difference"
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
      TabIndex        =   9
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label lblServerTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
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
      TabIndex        =   5
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Synchronization Server"
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
      TabIndex        =   4
      Top             =   240
      Width           =   2400
   End
   Begin VB.Image imgC2 
      Height          =   315
      Left            =   3120
      Picture         =   "frmSync.frx":2DFAE
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgC1 
      Height          =   315
      Left            =   2760
      Picture         =   "frmSync.frx":2E644
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   3850
      Picture         =   "frmSync.frx":2ECDA
      Top             =   40
      Width           =   315
   End
   Begin VB.Label lblSync 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Synchronization Status"
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
      TabIndex        =   1
      Top             =   840
      Width           =   2355
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clock not synchronized yet."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "frmSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TIME_ZONE_ID_UNKNOWN = 0
Const TIME_ZONE_ID_STANDARD = 1
Const TIME_ZONE_ID_DAYLIGHT = 2

Private RemoteTime As String        'the 32bit time stamp returned by the server
Private UTCTime As Date
Private TimeDelay As Single         'the time between the acknowledgement of the
                                    'connection and the data received. we compensate
                                    'by adding half of the round trip latency.
Private ZoneFactor As Long          'Adding this to UTC time will give us loacal time

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function ntohl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long

Dim ShapeTheForm As clsTransForm 'make a reference to the class

Private Sub cmdAdjust_Click()
    Dim ST As SYSTEMTIME
   
    Timer2.Enabled = False
     
    'fill a SYSTEMTIME structure with the appropriate values
    With ST
        .wYear = Year(UTCTime)
        .wMonth = Month(UTCTime)
        .wDay = Day(UTCTime)
        .wHour = Hour(UTCTime)
        .wMinute = Minute(UTCTime)
        .wSecond = Second(UTCTime)
    End With
    Timer2.Enabled = True
    'and call the API with the new date & time
    If SetSystemTime(ST) Then
        lblStat.Caption = "PC Clock synchronised"
        cmdAdjust.Enabled = False
    Else
        lblStat.Caption = "Pc Clock not synchronised!"
    End If
End Sub

'Set ShapeTheForm = New clsTransForm 'instantiate the object from the class

Private Sub cmdCheck_Click()
    'clear the string used for incoming data
    RemoteTime = Empty
   
    cmdCheck.Enabled = False
    Timer2.Enabled = False
    txtServerTime.Text = ""
    txtDifference.Text = "Calculating time difference..."
    ZoneFactor = 60 * AdjustTimeForTimeZone
   
    'connect
    With Socket
        If .State <> sckClosed Then .Close
        .RemoteHost = cboServers.Text
        .RemotePort = 37  'port 37 is the timserver port
        .Connect
    End With
End Sub

Private Sub Form_Load()
    Me.top = (Screen.Height - Me.Height) / 2
    Me.left = (Screen.Width - Me.Width) / 2
    Set ShapeTheForm = New clsTransForm 'instantiate the object from the class
    ShapeTheForm.ShapeMe frmSync, RGB(255, 255, 0), False, App.Path & "\Sync.dat"
    With cboServers
        .AddItem "time.nist.gov"
        .AddItem "time.windows.com"
        .AddItem "time-a.timefreq.bldrdoc.gov"
        .AddItem "time-b.timefreq.bldrdoc.gov"
        .AddItem "time-c.timefreq.bldrdoc.gov"
        .AddItem "utcnist.colorado.edu"
        .AddItem "time-nw.nist.gov"
        .AddItem "nist1.nyc.certifiedtime.com"
        .AddItem "nist1.dc.certifiedtime.com"
        .AddItem "nist1.sjc.certifiedtime.com"
        .AddItem "nist1.datum.com"
        .AddItem "ntp2.cmc.ec.gc.ca"
        .AddItem "ntps1-0.uni-erlangen.de"
        .AddItem "ntps1-1.uni-erlangen.de"
        .AddItem "ntps1-2.uni-erlangen.de"
        .AddItem "ntps1-0.cs.tu-berlin.de"
        .AddItem "time.ien.it"
        .AddItem "ptbtime1.ptb.de"
        .AddItem "ptbtime2.ptb.de"
        .ListIndex = 0
    End With
    txtServerTime.Text = ""
    txtLocalTime.Text = Format(Now, "hh:mm:ss") + "  " + Format(Now, "ddd mmm d, yyyy")
    txtDifference.Text = ""
    lblStat.Caption = ""
    Timer1.Enabled = True
    cmdAdjust.Enabled = False
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

Private Sub Socket_Close()
    Dim NTPTime As Double
    Dim LocalTime As Date
    Dim dwSecondsSince1990 As Long
    Dim Difference As Long
    
    Socket.Close
    RemoteTime = Trim(RemoteTime)
    If Len(RemoteTime) = 4 Then
        Timer2.Enabled = True
        'since the data was returned in a string,
        'format it back into a numeric value
        NTPTime = Asc(left$(RemoteTime, 1)) * (256 ^ 3) + _
                  Asc(Mid$(RemoteTime, 2, 1)) * (256 ^ 2) + _
                  Asc(Mid$(RemoteTime, 3, 1)) * (256 ^ 1) + _
                  Asc(right$(RemoteTime, 1))
        
        'calculate round trip delay
        TimeDelay = (Timer - TimeDelay)
        'and create a valid date based on the seconds since January 1, 1990
        dwSecondsSince1990 = NTPTime - 2840140800# + CDbl(TimeDelay)
        UTCTime = DateAdd("s", CDbl(dwSecondsSince1990), #1/1/1990#)
        'convert UTC time to local time and get the difference
        LocalTime = DateAdd("s", CDbl(ZoneFactor), UTCTime)
        Difference = DateDiff("s", Now, LocalTime)
        
        cmdAdjust.Enabled = True
        
        If Difference < 0 Then
            txtDifference.Text = "PC Clock " + CStr(-Difference) + " sec ahead."
        ElseIf Difference > 0 Then
            txtDifference.Text = "PC Clock " + CStr(Difference) + " sec behind."
        Else
            txtDifference.Text = "Clock Sync'd."
        End If
        lblStat.Caption = "Server successfully contacted."
    Else
        txtServerTime.Text = ""
        lblStat.Caption = "Time received not valid."
        Timer2.Enabled = False
        cmdAdjust.Enabled = False
    End If
    cmdCheck.Enabled = True
End Sub

Private Sub Socket_Connect()
    lblStat.Caption = "Connected to Server"
    TimeDelay = Timer
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    lblStat = "Server is sending data"
    Socket.GetData sData, vbString
    RemoteTime = RemoteTime & sData
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStat.Caption = Description
    cmdCheck.Enabled = True
End Sub

Private Sub Timer1_Timer()
    txtLocalTime.Text = Format(Now, "hh:mm:ss") + "  " + Format(Now, "ddd mmm d, yyyy")
End Sub

Private Sub Timer2_Timer()
    txtServerTime.Text = Format(UTCTime, "hh:mm:ss") + "  " + Format(UTCTime, "ddd mmm d, yyyy")
    UTCTime = DateAdd("s", CDbl(1), UTCTime)
End Sub

Private Function AdjustTimeForTimeZone() As Long
'Returns the amount of adjustment in seconds necessary from UTC time for the
'current system by checking the system's time zone and daylight savings properties
    Dim TZI As TIME_ZONE_INFORMATION
    Dim RetVal As Long
    Dim ZoneCorrection As Long
    
    RetVal = GetTimeZoneInformation(TZI)
    ZoneCorrection = TZI.Bias
    If RetVal = TIME_ZONE_ID_STANDARD Then
        ZoneCorrection = ZoneCorrection + TZI.StandardBias
    ElseIf RetVal = TIME_ZONE_ID_DAYLIGHT Then
            ZoneCorrection = ZoneCorrection + TZI.DaylightBias
    Else
        MsgBox "Unable to get zone information.", vbExclamation, "Error"
    End If
    AdjustTimeForTimeZone = -ZoneCorrection     'correction in minutes
End Function

