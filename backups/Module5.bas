Attribute VB_Name = "Module5"
Option Explicit
Sub test()
'Call FuzzyMatch

End Sub

Sub FuzzyMatch()
'ByVal string1 As String, _
'                    ByVal string2 As String, _
'                    min_percentage As Long)

Dim i As Long, j As Long
Dim string1, string2 As String
Dim string1_length As Long
Dim string2_length As Long
Dim distance() As Long, result As Long
Dim min_percentage As Long
Dim fuzzy As String

fuzzy = Empty

string1 = Range("A1").Value
string2 = Range("B1").Value

string1_length = Len(string1)
string2_length = Len(string2)

min_percentage = 50
' Check if not too long
If string1_length >= string2_length * (min_percentage / 100) Then
    ' Check if not too short
    If string1_length <= string2_length * ((200 - min_percentage) / 100) Then

        ReDim distance(string1_length, string2_length)
        For i = 0 To string1_length: distance(i, 0) = i: Next
        For j = 0 To string2_length: distance(0, j) = j: Next

        For i = 1 To string1_length
            For j = 1 To string2_length
                If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
                    distance(i, j) = distance(i - 1, j - 1)
                Else
                    distance(i, j) = Application.WorksheetFunction.Min _
                    (distance(i - 1, j) + 1, _
                     distance(i, j - 1) + 1, _
                     distance(i - 1, j - 1) + 1)
                End If
            Next
        Next
        result = distance(string1_length, string2_length) 'The distance
    End If
End If

Select Case result
    Case Is = 0
        fuzzy = "Perfect Match (0)"
    Case Is <> 0
        fuzzy = (CLng((100 - ((result / string1_length) * 100)))) & _
                 "% (" & result & ")" 'Convert to percentage
    Case Else
        fuzzy = "Not a match"
End Select
                 
                 
'If result <> 0 Then
'    fuzzy = (CLng((100 - ((result / string1_length) * 100)))) & _
'                 "% (" & result & ")" 'Convert to percentage
'Else
'    fuzzy = "Not a match"
'End If

End Sub


Sub SearchReplace()
Attribute SearchReplace.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Mreplacetest Macro
'

'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Replace What:="<br>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
    Selection.Replace What:="<p>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
    Selection.Replace What:="</p>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
    Selection.Replace What:="</em>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
    Selection.Replace What:="<em>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
    Selection.Replace What:="</br>", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False
End Sub


