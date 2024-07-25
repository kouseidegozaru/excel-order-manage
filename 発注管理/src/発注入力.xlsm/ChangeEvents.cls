VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangeEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ws As Worksheet
Private targetRange As range

Public Sub Initialize(targetWs As Worksheet)
    If targetWs Is Nothing Then
        MsgBox "Invalid worksheet provided.", vbCritical
        Exit Sub
    End If
    Set ws = targetWs
End Sub

Public Sub SetRange(targetRng As range)
    If targetRng Is Nothing Then
        MsgBox "Invalid range provided.", vbCritical
        Exit Sub
    End If
    Set targetRange = targetRng
End Sub

Public Function IsChanged(strRange As String) As Boolean
    IsChanged = False
    If ws Is Nothing Then
        MsgBox "Worksheet is not initialized.", vbCritical
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    If Not Intersect(targetRange, ws.range(strRange)) Is Nothing Then
        IsChanged = True
    End If
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    IsChanged = False
End Function

