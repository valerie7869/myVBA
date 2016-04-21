Attribute VB_Name = "MenuCode"
Option Explicit

'*****Do not change the code in this module*******

Sub WBCreatePopUp()
    Dim MenuSheet As Worksheet
    Dim MenuItem As Object
    Dim SubMenuItem As CommandBarButton
    Dim Row As Integer
    Dim MenuLevel, NextLevel, MacroName, Caption, Divider, FaceId

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Location for menu data
    Set MenuSheet = ThisWorkbook.Sheets("MenuSheet")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''

    '   Make sure the menus aren't duplicated
    Call WBRemovePopUp

    '   Initialize the row counter
    Row = 5

    '   Add the menu, menu items and submenu items using
    '   data stored on MenuSheet

    ' First we create a PopUp menu with the name of the value in B2
    With Application.CommandBars.Add(ThisWorkbook.Sheets("MenuSheet"). _
                                     Range("B2").Value, msoBarPopup, False, True)

        Do Until IsEmpty(MenuSheet.Cells(Row, 1))
            With MenuSheet
                MenuLevel = .Cells(Row, 1)
                Caption = .Cells(Row, 2)
                MacroName = .Cells(Row, 3)
                Divider = .Cells(Row, 4)
                FaceId = .Cells(Row, 5)
                NextLevel = .Cells(Row + 1, 1)
            End With

            Select Case MenuLevel
            Case 2    ' A Menu Item
                If NextLevel = 3 Then
                    Set MenuItem = .Controls.Add(Type:=msoControlPopup)
                Else
                    Set MenuItem = .Controls.Add(Type:=msoControlButton)
                    MenuItem.OnAction = ThisWorkbook.Name & "!" & MacroName
                End If
                MenuItem.Caption = Caption
                If FaceId <> "" Then MenuItem.FaceId = FaceId
                If Divider Then MenuItem.BeginGroup = True

            Case 3    ' A SubMenu Item
                Set SubMenuItem = MenuItem.Controls.Add(Type:=msoControlButton)
                SubMenuItem.Caption = Caption
                SubMenuItem.OnAction = ThisWorkbook.Name & "!" & MacroName
                If FaceId <> "" Then SubMenuItem.FaceId = FaceId
                If Divider Then SubMenuItem.BeginGroup = True
            End Select
            Row = Row + 1
        Loop
    End With
End Sub

Sub RDBDisplayPopUp()
Attribute RDBDisplayPopUp.VB_ProcData.VB_Invoke_Func = "l\n14"
    On Error Resume Next
    Application.CommandBars(ThisWorkbook.Sheets("MenuSheet").Range("B2").Value).ShowPopup
    On Error GoTo 0
End Sub

Sub EditMenu()
    ThisWorkbook.IsAddin = False
End Sub

Sub WBRemovePopUp()
    On Error Resume Next
    Application.CommandBars(ThisWorkbook.Sheets("MenuSheet").Range("B2").Value).Delete
    On Error GoTo 0
End Sub

Sub RefreshMenu()
    Call WBCreatePopUp
    Call RDBDisplayPopUp
    'MsgBox "Click on the button in the Tools menu to see if your menu is correct.", vbOKOnly, "Favorite Macro Menu"
End Sub

Sub SaveAddin()
    Call WBCreatePopUp
    Range("A1").Select
    ThisWorkbook.IsAddin = True
    ThisWorkbook.Save
End Sub

Sub CancelAddin()
    ThisWorkbook.IsAddin = True
    ThisWorkbook.Saved = True
End Sub

Sub Macro999()
Attribute Macro999.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro999 Macro
'

'
    Sheets("Summary").Select
    Sheets("Summary").Copy Before:=Workbooks( _
        "dms---oem_loas11370-oemprojects_license-analysis (5).xls").Sheets(1)
End Sub
