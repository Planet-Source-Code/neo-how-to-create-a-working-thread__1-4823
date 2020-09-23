VERSION 5.00
Begin VB.Form frmThread 
   Caption         =   "Demo Of How To Work With Threads"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Demo of How To Work With Threads"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1575
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "frmThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

Dim myThreadTop As New clsThreads, myThreadBottom As New clsThreads

On Error Resume Next
With myThreadTop
    .Initialize AddressOf FlickerTop
    .Enabled = True
End With
With myThreadBottom
    .Initialize AddressOf FlickerBottom
    .Enabled = True
End With

MsgBox "Let's wait and see what happens..."

Set myThreadTop = Nothing
Set myThreadBottom = Nothing

End Sub
