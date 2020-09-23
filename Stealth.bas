Attribute VB_Name = "Stealth"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'This module is to show you how to make a form run
'invisible.  A lot of people say "Oh, just use
'App.TaskVisible = False", but that doesn't work
'on non-ole projects.  That is what this is for.
'In order to do this, you have to use the functions
'below.  HideExWindow will hide the form from
'the TaskBar, and the TaskList.  ShowExWindow
'will show the form.  Feel free to use this
'Module in any of your projects.
'Good Luck!  And Happy Coding!
'-Tucker Nance, BurnerGuy@aol.com


'EX.
'Private Sub cmdMakeStealth_Click()
'HideExWindow (txtClass.Text)
'End Sub

'Private Sub cmdMakeVisible_Click()
'ShowExWindow (txtClass.Text)
'End Sub



Sub HideExWindow(Class)
HideW% = FindWindow(Class, vbNullString)
Call ShowWindow(HideW%, 0)
End Sub
