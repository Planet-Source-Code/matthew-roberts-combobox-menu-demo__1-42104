VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Menu ComboBox Demo"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuTest 
      Caption         =   "&Reporting"
   End
   Begin VB.Menu mnuTest2 
      Caption         =   "&Reporting2"
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuTest_Click()
      Dim Menu As New frmCustomMenu
    
    With Menu
        '   Position the menu to be at the location clicked
        '   Will have to be adjusted for each menu
        .Left = Me.Left + 50
        .Top = Me.Top + 700
        '   Open the menu form modally
        .Show 1
        '   Retrieve the value selected on the menu form
        MsgBox "Value Selected: " & Menu.SelectedValue
    End With
End Sub

Private Sub mnuTest2_Click()
       Dim Menu As New frmCustomMenu
    
    With Menu
        '   Position the menu to be at the location clicked
        '   Will have to be adjusted for each menu
        .Left = Me.Left + 1000
        .Top = Me.Top + 700
        '   Open the menu form modally
        .Show 1
        '   Retrieve the value selected on the menu form
        MsgBox "Value Selected: " & Menu.SelectedValue
    End With
    

End Sub

Private Sub ShowMenu()


End Sub



