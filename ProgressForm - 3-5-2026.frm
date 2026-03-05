VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Processing FTZ Duty Calculations"
   ClientHeight    =   3432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8988.001
   OleObjectBlob   =   "ProgressForm - 3-5-2026.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.lblStatus.Caption = "Initializing..."
    Me.lblPhase.Caption = ""
    Me.lblEntry.Caption = ""
    Me.ProgressBar1.Width = 0
End Sub

Public Property Let ProgressValue(value As Integer)
    If value < 0 Then value = 0
    If value > 100 Then value = 100
    Me.ProgressBar1.Width = (frameProgress.Width - 4) * (value / 100)
End Property

Public Property Get ProgressValue() As Integer
    ProgressValue = (Me.ProgressBar1.Width / (frameProgress.Width - 4)) * 100
End Property
