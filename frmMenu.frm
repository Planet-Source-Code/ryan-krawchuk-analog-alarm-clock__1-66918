VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAlarm 
         Caption         =   "&Alarm"
      End
      Begin VB.Menu mnuFileSync 
         Caption         =   "&Synchronize"
      End
      Begin VB.Menu mnuFileH1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptChange 
         Caption         =   "&Change Character"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuFileAlarm_Click()
    frmAlarm.Show
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileSync_Click()
    frmSync.Show
End Sub

Private Sub mnuOptChange_Click()
    On Error Resume Next
    frmMain.msAgent.ShowDefaultCharacterProperties
End Sub
