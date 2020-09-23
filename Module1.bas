Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub FlickerTop()

Static BgColor As Long
Dim lTick As Long, lCounter As Long

On Error Resume Next
For lCounter = 0 To 5999
    If BgColor <> &HFF& Then BgColor = &HFF& Else BgColor = &HFF00&
    frmThread.Picture1.BackColor = BgColor
    lTick = GetTickCount
    While GetTickCount - lTick < 1250
    Wend
Next

End Sub
Public Sub FlickerBottom()

Static BgColor As Long
Dim lTick As Long, lCounter As Long

On Error Resume Next
For lCounter = 0 To 5999
    If BgColor <> &HFFFF& Then BgColor = &HFFFF& Else BgColor = &HFF0000
    frmThread.Picture2.BackColor = BgColor
    lTick = GetTickCount
    While GetTickCount - lTick < 500
    Wend
Next

End Sub
