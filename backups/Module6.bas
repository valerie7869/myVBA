Attribute VB_Name = "Module6"
Option Explicit

Sub Wrong()
  Dim i As Integer, a As Integer, b As Integer
  On Error GoTo Ignore
  For i = 1 To 2
    a = i / b
Ignore:
   Next
End Sub

Sub Right()
  Dim i As Integer, a As Integer, b As Integer
  On Error GoTo Errorhandler
  For i = 1 To 2
    a = i / b
  Next
  Exit Sub

Errorhandler:
  Select Case Err.Number
    Case 11
      'Ignore this error
      Resume Next
    Case Else
      Debug.Print "Source     : " & Err.Source
      Debug.Print "Error      : " & Err.Number
      Debug.Print "Description: " & Err.Description
      If MsgBox("Error " & Err.Number & ": " & vbNewLine & vbNewLine & _
          Err.Description & vbNewLine & vbNewLine & _
          "Enter debug mode?", vbOKCancel + vbDefaultButton2, Err.Source) = vbOK Then
        Stop 'Press F8 twice
        Resume
      End If
  End Select
End Sub

