VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmServer 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eCas Database Backup Server"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   ControlBox      =   0   'False
   Icon            =   "Server.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5880
      ScaleHeight     =   825
      ScaleWidth      =   3225
      TabIndex        =   15
      Top             =   5640
      Width           =   3255
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Developed By : Bharat Jagtap"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   2865
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bharat_jagtap26@yahoo.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1125
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Minimize"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Configure BackUp Server"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   13
      Top             =   135
      Width           =   3075
   End
   Begin VB.TextBox TxtSaveFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "C:\"
      Top             =   6210
      Visible         =   0   'False
      Width           =   5610
   End
   Begin VB.ListBox LstLog 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "Server.frx":08CA
      Left            =   135
      List            =   "Server.frx":08CC
      TabIndex        =   1
      Top             =   585
      Width           =   10320
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   6750
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSWinsockLib.Winsock WindowsSocket 
      Left            =   675
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2543
   End
   Begin MSComctlLib.ProgressBar ProgressTransfer 
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   5040
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   9810
      Picture         =   "Server.frx":08CE
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   4230
      Width           =   510
   End
   Begin VB.CommandButton CmdSendFile 
      Caption         =   "Transfer File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8370
      TabIndex        =   6
      Top             =   4545
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Stop Backup Server"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6840
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bytes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   14
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JBackup Server [MS SQL Server] "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4935
      TabIndex        =   8
      Top             =   135
      Width           =   5370
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status : Waiting......."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   5445
      Width           =   2250
   End
   Begin VB.Label lblpacketsize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2048"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1485
      TabIndex        =   4
      Top             =   5760
      Width           =   540
   End
   Begin VB.Label lblpacket 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packet Size :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   5760
      Width           =   1380
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Client
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Dim AcceptTransfer As Byte
Dim PacketState As Byte
Dim CanceledTransfer As Boolean
Dim PacketWait As Boolean
'common
Dim KickoffHost As Boolean

'Server
Dim FileToSaveName As String
Dim FileToSaveSize As Long
Dim FileNumber As Integer
Dim FileDataRecievedSoFar As Long
Dim FileInTransfer As Boolean
 

 

Private Sub CmdSendFile_Click()
Call TransferFile(Trim(TxtSaveFile))
'Call TransferFile("D:\a.zip")
End Sub

Private Sub Command1_Click()
    If gdbConnection.State Then
        gdbConnection.Close
    End If

Me.Hide
End Sub

'Private Sub Command2_Click()
'    If Command2.Caption = "&Stop Backup Server" Then
'        Command2.Caption = "&Start Backup Server"
'        WriteLog "Backup Server Stopped"
'        WindowsSocket.Close
'    Else
'        Command2.Caption = "&Stop Backup Server"
'        WriteLog "Backup Server Started"
'        WriteLog "Listening.."
'        WindowsSocket.Listen
'    End If
'End Sub

Private Sub Command3_Click()
WindowsSocket.Close
CloseConnection
End
End Sub

Private Sub Command4_Click()
Form1.Show vbModal
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call IsApplicationAlreadyInstenciated
    
    If Trim(GetSetting("ABSOLUTE NetWare", "NetSoft", "DBServer")) = "" Then
        Form1.Show vbModal
    End If
    
    'end get ip
    
    WindowsSocket.Listen
    WriteLog "listening"
    
      SaveSetting HKEY_LOCAL_MACHINE, "SOFTWARE\ABSOLUTE\Windows\CurrentVersion\Run", App.EXEName, App.Path & "\" & App.EXEName & ".exe"
    'HKEY_LOCAL_MACHINE , "Software\ABSOLUTE\Windows\CurrentVersion\Run", Zero, KEY_ALL_ACCESS, Hkey
    '''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''
    'We set up Picture1 to accept callback data
    'in it's MouseMove procedure.
    NotifyIcon.cbSize = Len(NotifyIcon)
    NotifyIcon.hWnd = Picture1.hWnd
    NotifyIcon.uID = 1&
    NotifyIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    NotifyIcon.uCallbackMessage = WM_MOUSEMOVE
    
    'Now, we set up the icon and tool tip message
    NotifyIcon.hIcon = Picture1.Picture
    NotifyIcon.szTip = "JBackup Server" & Chr$(0)
    
    'Lastly, we add the icon
    Shell_NotifyIcon NIM_ADD, NotifyIcon
    
    ''''''''''''''''''''''''''''''.
    frmLogin.Show
    frmLogin.Hide
    CloseConnection
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrHand
    If Hex(x) = "1E0F" Then
       'Use "1E3C" for right-click
       'MsgBox (" SQL BackUp.."), 48, ("JBackup  BackUp Utility")
     
       Me.Show
       Form_Load
    End If
    CloseConnection
    Exit Sub

ErrHand:
  If InStr(1, Err.Description, "modal") > 1 Then
  Else
  MsgBox Err.Description
  End If
End Sub
Private Sub WindowsSocket_Close()
    On Error GoTo ErrHand:
    WindowsSocket.Close
    DoEvents
    
    FileDataRecievedSoFar = 0
    FileInTransfer = False
    
    WriteLog "Status: Disconnecting from Client..."
    WriteLog "Status: Listening for connection...."

    
    DoEvents
    
    WindowsSocket.Listen
    DoEvents
    Close #FileNumber
    CloseConnection
    Exit Sub
ErrHand:
    WriteLog "err: " & Err.Description
End Sub

Private Sub WindowsSocket_ConnectionRequest(ByVal requestID As Long)

    On Error Resume Next
    If WindowsSocket.State <> sckClosed Then WindowsSocket.Close
    DoEvents
    LstLog.Clear
    WindowsSocket.Accept requestID
    DoEvents
    
    LstLog.Clear
    WriteLog "Connection Request Accepted : " & WindowsSocket.RemoteHostIP & ":" & WindowsSocket.RemotePort
    WindowsSocket.SendData "Received:accept"

    
    
End Sub

Public Sub WriteLog(ByVal strMsg As String)
    Dim strMsg1 As String
    strMsg1 = Now() & " : " & strMsg
    LstLog.AddItem strMsg1
    LstLog.Text = strMsg1
End Sub


Private Sub WindowsSocket_DataArrival(ByVal bytesTotal As Long)
    
    Dim DataReceived As String
    Dim strArr() As String
    WindowsSocket.GetData DataReceived
    
    DoEvents
    
    If FileInTransfer = True Then
        FileDataRecievedSoFar = FileDataRecievedSoFar + Len(DataReceived)
        If FileDataRecievedSoFar >= FileToSaveSize Then
            Put #FileNumber, , CStr(Left(DataReceived, Len(DataReceived) - (FileDataRecievedSoFar - FileToSaveSize)))
            lblstatus = "Status: 100%, File transfer complete"
            ProgressTransfer.value = ProgressTransfer.Max
            FileDataRecievedSoFar = 0
            FileInTransfer = False
            
            DoEvents
            
            WindowsSocket.Close
            DoEvents
            
            WriteLog "Status: Disconnecting from Client..."
            WriteLog "Status: Listening for connection...."


            WindowsSocket.Listen
            
            Close #FileNumber
        Else
            Put #FileNumber, , CStr(DataReceived)
            lblstatus = "Status: " & CInt((FileDataRecievedSoFar / FileToSaveSize) * 100) & "%, Received " & FormatKB(FileDataRecievedSoFar) & " of " & FormatKB(FileToSaveSize)
            ProgressTransfer.value = (FileDataRecievedSoFar / FileToSaveSize) * 100
            DoEvents
        End If
    Else
        If LCase(Left(DataReceived, 8)) = LCase("Request#") Then
            Select Case LCase(Mid(DataReceived, 10, 6))
                Case LCase("PACKET")
                    WindowsSocket.SendData "Received:PACKET:" & lblpacketsize.Caption
            End Select
        ElseIf LCase(Left(DataReceived, 8)) = LCase("File~~~#") Then
            Select Case LCase(Mid(DataReceived, 10, 6))
                Case "name~#"
                    FileToSaveName = Right(DataReceived, Len(DataReceived) - 16)
                Case "size~#"
                    FileToSaveSize = CLng(Right(DataReceived, Len(DataReceived) - 16))
                    
                    FileNumber = FreeFile
                    Open TxtSaveFile & FileToSaveName For Binary Access Write As #FileNumber
                    
                    WindowsSocket.SendData "Received:PEND~#:1"
                    WriteLog "Receiving file..." & TxtSaveFile & FileToSaveName
                     
                     
                    FileInTransfer = True
                   
            End Select
        End If
    End If
    DoEvents
''cODE fOR cLIENT
     If LCase(Left(DataReceived, 8)) = LCase("Received") Then
        Select Case LCase(Mid(DataReceived, 10, 6))
            Case LCase("PACKET")
                LookupPacket = CInt(Right(DataReceived, Len(DataReceived) - 16))
                If LookupPacket < CInt(lblpacketsize) Then
                    WriteLog "The packet size your requesting has been denied by the server. The packet size is now set to optimum size. When ready please click on 'Start Transfer' once again."
                    PacketState = 0
                Else
                    PacketState = 1
                End If
            Case LCase("PEND~#")
                AcceptTransfer = CInt(Right(DataReceived, Len(DataReceived) - 16))
            Case LCase("denied")
                WriteLog "The RemoteHost denied access to connect."

                WriteLog "Status: Listening for connection...."
                WindowsSocket.Listen

                WindowsSocket.Close
                DoEvents
                WindowsSocket.Listen
            Case LCase("kickH#")
                KickoffHost = True
        End Select
    End If
    
    If UCase(Left(DataReceived, 10)) = UCase("TAKEBACKUP") Then
    strArr = Split(DataReceived, ",")
        g_sServerName = strArr(1)
        g_sPassword = strArr(2)
        g_sDatabase = strArr(3)
        Call frmLogin.Backup
        TxtSaveFile = Mid(App.Path, 1, 2) & TxtSaveFile
        CmdSendFile_Click
    End If
End Sub
Private Function SendFile(FileName As String, CheckSize As Long)
    
    Dim FileNumber As Integer
    Dim FileBinary As String
    Dim BlockSize As Integer
    Dim SentSize As Long
        
    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
        
        WriteLog "Sending file..." & FileName
        
        BlockSize = CInt(lblpacketsize)
        FileBinary = Space(BlockSize)
                
        Do
            PacketWait = True
        
            Get #FileNumber, , FileBinary
            SentSize = SentSize + Len(FileBinary)
            If WindowsSocket.State <> 7 Then GoTo CheckError
            If SentSize > CheckSize Then
                ProgressTransfer.value = ProgressTransfer.Max
                WriteLog "Status: 100% Complete. " & FormatKB(CheckSize) & " Sent."
                lblstatus = "Status: 100% Complete. " & FormatKB(CheckSize) & " Sent."
                WindowsSocket.SendData Mid(FileBinary, 1, Len(FileBinary) - (SentSize - CheckSize))
            Else
                 ProgressTransfer.value = (SentSize / CheckSize) * ProgressTransfer.Max
                 lblstatus = "Status: " & CInt((SentSize / CheckSize) * ProgressTransfer.Max) & "% Complete, " & FormatKB(CheckSize - SentSize) & " Remaining..."
                WindowsSocket.SendData FileBinary
            End If
                        
            Do
                DoEvents
            Loop Until PacketWait = False
                        
            DoEvents: Loop Until EOF(FileNumber)
            
            SendFile = True
            
    Close #FileNumber
    Exit Function
    
CheckError:

    Close #FileNumber

End Function
Private Sub TransferFile(ByVal strFileName As String)
    On Error Resume Next
    If WindowsSocket.State <> 7 Then
        WindowsSocket.Close
        DoEvents
        WindowsSocket.Listen
        Exit Sub
    End If
    WindowsSocket.SendData "Request#:PACKET"
    PacketState = 2
    Do Until PacketState <> 2
        DoEvents
    Loop
    If PacketState = 0 Then Exit Sub
    
    If PacketState = 1 Then
        AcceptTransfer = 2
        WriteLog "Sending file attributes..."
        PacketWait = True
       
        WindowsSocket.SendData "File~~~#:name~#:" & GetFileName(strFileName)
        
        Do
            DoEvents
        Loop Until PacketWait = False
        
        PacketWait = True
        
        WindowsSocket.SendData "File~~~#:size~#:" & FileLen(strFileName)
        
        PacketWait = True
        
        WriteLog "Pending file request clearance..."
        
        Do Until AcceptTransfer <> 2
            DoEvents
        Loop
        
        If AcceptTransfer = 0 Then
            WriteLog "File transfer was denied by the server."
        Else
            If SendFile(strFileName, FileLen(strFileName)) = False Then
                If CanceledTransfer = False Then
                    If KickoffHost = True Then
                        WriteLog "File transfer failed. You was kicked off the server."
                    Else
                        WriteLog "File transfer failed, lost peer connection."
                    End If
                    WindowsSocket.Close
                    WindowsSocket.Close
                
                WriteLog "Status: Listening for connection...."
                WindowsSocket.Listen

                
                Else
                    WriteLog "File transfer canceled by the Host."
                End If
                CanceledTransfer = False
                KickoffHost = False
                WriteLog "Status: Not connected"
                
            Else
                WriteLog "File transfer complete."
                WindowsSocket.Close
                
                Call deleteBkfile
                DoEvents
                WriteLog "Status: Disconnecting from Client..."
                WriteLog "Status: Listening for connection...."
                WindowsSocket.Listen
            End If
        End If
        
    End If

End Sub

Function GetFileName(Path As String) As String

    For FindSep = 1 To Len(Path)
        If Mid(Path, Len(Path) - (FindSep - 1), 1) = "\" Or Mid(Path, Len(Path) - (FindSep - 1), 1) = "/" Then
            GetFileName = Right(Path, FindSep - 1)
            Exit Function
        End If
    Next FindSep
End Function

Private Sub WindowsSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    FileDataRecievedSoFar = 0
    FileInTransfer = False
    
    WindowsSocket.Close
    DoEvents
    WriteLog "Status: Disconnecting from Client..."
    WriteLog "Status: Listening for connection...."
    

    DoEvents
    
    WindowsSocket.Listen
    
    Close #FileNumber
End Sub
Public Function FormatKB(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim result As String
    Buffer = Space$(255)
    result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    If InStr(result, vbNullChar) > 1 Then FormatKB = Left$(result, InStr(result, vbNullChar) - 1)
End Function
Private Sub WindowsSocket_SendComplete()
    PacketWait = False
End Sub

Public Sub deleteBkfile()
    On Error Resume Next
    Dim fso As New FileSystemObject
    fso.DeleteFile TxtSaveFile
End Sub
Public Sub GetServerWindowsName()
On Error GoTo ErrHand
 
Dim objBk As New clsBackUp
    CodeModule.strServerWindowsName = objBk.GetServerWindowsName(txtpwd)
    CloseConnection
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Public Sub CloseConnection()
On Error Resume Next
 If gdbConnection.State Then
        gdbConnection.Close
 End If
End Sub
