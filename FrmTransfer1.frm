VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmTransfer1 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eCas Database Backup Scheduler....."
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   ControlBox      =   0   'False
   FillColor       =   &H00800000&
   Icon            =   "FrmTransfer1.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2640
      TabIndex        =   7
      Text            =   "Critical%data@very*userfull~05"
      Top             =   6360
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8505
      TabIndex        =   6
      Top             =   315
      Width           =   1410
   End
   Begin VB.CommandButton cmdTakeBackup 
      Caption         =   "Take Backup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   315
      Width           =   1770
   End
   Begin VB.TextBox txtServerIP 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   315
      Width           =   1590
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
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9825
   End
   Begin MSComctlLib.ProgressBar PrgTransferStatus 
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   5040
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar PrgTransfer 
      Height          =   435
      Left            =   135
      TabIndex        =   2
      Top             =   5445
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock WindowsSocket 
      Left            =   120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup File Password :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   6398
      Width           =   2355
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Status : Waiting..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   6030
      Width           =   2475
   End
End
Attribute VB_Name = "FrmTransfer1"
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
Dim iTryToConnect As Integer



Public Sub WriteLog(ByVal strMsg As String)
    Dim strMsg1 As String
    strMsg1 = Now() & " : " & strMsg
    LstLog.AddItem strMsg1
    LstLog.Text = strMsg1
End Sub
'ASk server for Connection
'Client Side Function
Private Sub SeekConnectionWithServer(ByVal strServerAddress As String)

    On Error GoTo CheckError
    CntPort = 2543
    WindowsSocket.Close
    DoEvents
    WindowsSocket.Connect strServerAddress, CntPort

    WriteLog "Status: Pending " & strServerAddress & ":" & CntPort

    Do Until WindowsSocket.State <> 6
        DoEvents
    Loop

    If WindowsSocket.State = 7 Then
        WriteLog "Status: Connected " & strServerAddress & ":" & CntPort
    Else
        GoTo CheckError
    End If

    Exit Sub

CheckError:
    
    WindowsSocket.Close
    WriteLog "JBackup  Backup Server is Buzy or Service Not Started Please Try Again.."
    DoEvents
    
    
    cmdclose.Enabled = True
    cmdTakeBackup.Enabled = True
      lblstatus = ""
   
End Sub
 'Sends File
 'Server as well Client side Function

 Private Sub TransferFile(ByVal strFileName As String)
 On Error GoTo ErrorHandler

    
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
                    DoEvents
                Else
                    WriteLog "File transfer canceled by the Host."
                End If
                CanceledTransfer = False
                KickoffHost = False
                WriteLog "Status: Not connected"

            Else
                WriteLog "File transfer complete."
                WriteLog "Status: Not connected"
                WindowsSocket.Close
                DoEvents
            End If
        End If

    End If

    
Exit Sub
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents
End Sub


Private Sub cmdClose_Click()
On Error GoTo ErrorHandler
    WindowsSocket.Close
    DoEvents
    If gdbConnection.State Then
        gdbConnection.Close
    End If

    Unload Me
Exit Sub
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
End Sub

Private Sub cmdTakeBackup_Click()
On Error GoTo ErrorHandler
        bTakeBackup = False
        cmdTakeBackup.Enabled = False
        cmdclose.Enabled = False
        
        If Trim(txtServerIP.Text) <> "" Then
            SeekConnectionWithServer txtServerIP.Text
        End If
Exit Sub
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents

End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler

    If bTakeBackup Then
    iTryToConnect = 1
        
        Call cmdTakeBackup_Click
        
    End If
Exit Sub
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True

End Sub

 

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    WindowsSocket.Close
    DoEvents
    bTakeBackup = False
    frmLogin.Timer1.Enabled = True
End Sub

Private Sub WindowsSocket_Connect()
    WriteLog "Connected"
    WindowsSocket.SendData "TAKEBACKUP," & Trim(g_sServerName) & "," & Trim(g_sPassword) & "," & Trim(g_sDatabase)
    cmdclose.Enabled = False
End Sub
Private Sub WindowsSocket_Close()
    WriteLog "Connetion Closed"
    cmdclose.Enabled = True
    WindowsSocket.Close
End Sub
Private Sub WindowsSocket_DataArrival(ByVal bytesTotal As Long)
On Error GoTo ErrorHandler
    
'ClentCode

    Dim DataReceived As String
    Dim LookupPacket As Integer

    WindowsSocket.GetData DataReceived

    If LCase(Left(DataReceived, 8)) = LCase("Received") Then
        Select Case LCase(Mid(DataReceived, 10, 6))
            Case LCase("PACKET")
                LookupPacket = CInt(Right(DataReceived, Len(DataReceived) - 16))
                If LookupPacket < CInt(lblpacketsize) Then
                    WriteLog "The packet size your requesting has been denied by the server. The packet size is now set to optimum size. When ready please click on 'Start Transfer' once again."
                    PacketScroll.value = LookupPacket
                    PacketState = 0
                Else
                    PacketState = 1
                End If
            Case LCase("PEND~#")
                AcceptTransfer = CInt(Right(DataReceived, Len(DataReceived) - 16))
            Case LCase("denied")
                WriteLog "The server denied access to connect."
              
                WindowsSocket.Close
                DoEvents
            Case LCase("accept")
                
            Case LCase("kickH#")
                KickoffHost = True
        End Select
    End If
'ServerCode to Recive File
'===================================================================================================
    If FileInTransfer = True Then
        FileDataRecievedSoFar = FileDataRecievedSoFar + Len(DataReceived)
        If FileDataRecievedSoFar >= FileToSaveSize Then
            Put #FileNumber, , CStr(Left(DataReceived, Len(DataReceived) - (FileDataRecievedSoFar - FileToSaveSize)))

            WriteLog "Transfer Status: 100%, File transfer complete"
            lblstatus = "Transfer Status: 100%, File transfer complete"
            WriteLog "Backup File Name :[" & g_strBackupPath & "\" & FileToSaveName & "]"
            PrgTransfer.value = PrgTransfer.Max

            FileDataRecievedSoFar = 0
            FileInTransfer = False
            
            WindowsSocket.Close
            
            DoEvents
            Close #FileNumber
            cmdclose.Enabled = True
         Else
            Put #FileNumber, , CStr(DataReceived)
            lblstatus = "Status: " & CInt((FileDataRecievedSoFar / FileToSaveSize) * 100) & "%, Received " & FormatKB(FileDataRecievedSoFar) & " of " & FormatKB(FileToSaveSize)
            PrgTransfer.value = (FileDataRecievedSoFar / FileToSaveSize) * 100
        End If
    Else
        If LCase(Left(DataReceived, 8)) = LCase("Request#") Then
            Select Case LCase(Mid(DataReceived, 10, 6))
                Case LCase("PACKET")
                    WindowsSocket.SendData "Received:PACKET:" & 2048
            End Select
        ElseIf LCase(Left(DataReceived, 8)) = LCase("File~~~#") Then
            Select Case LCase(Mid(DataReceived, 10, 6))
                Case "name~#"
                    FileToSaveName = Right(DataReceived, Len(DataReceived) - 16)
                Case "size~#"
                    FileToSaveSize = CLng(Right(DataReceived, Len(DataReceived) - 16))


                        FileNumber = FreeFile
                        Open g_strBackupPath & "\" & FileToSaveName For Binary Access Write As #FileNumber

                        WindowsSocket.SendData "Received:PEND~#:1"
                        WriteLog "Receiving file..." & TxtSaveFile & FileToSaveName
                        WriteLog "Packet size: " & 2048

                        FileInTransfer = True

            End Select
        End If
    End If
'===================================================================================================


Exit Sub
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents
End Sub

Function GetFileName(Path As String) As String
On Error GoTo ErrorHandler

    For FindSep = 1 To Len(Path)
        If Mid(Path, Len(Path) - (FindSep - 1), 1) = "\" Or Mid(Path, Len(Path) - (FindSep - 1), 1) = "/" Then
            GetFileName = Right(Path, FindSep - 1)
            Exit Function
        End If
    Next FindSep

On Error GoTo ErrorHandler

    
Exit Function
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents

End Function

Public Function FormatKB(ByVal Amount As Long) As String
On Error GoTo ErrorHandler

    

    Dim Buffer As String
    Dim result As String
    Buffer = Space$(255)
    result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    If InStr(result, vbNullChar) > 1 Then FormatKB = Left$(result, InStr(result, vbNullChar) - 1)
    
    On Error GoTo ErrorHandler

    
Exit Function
ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents

End Function

Private Function SendFile(FileName As String, CheckSize As Long)

On Error GoTo ErrorHandler

    

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
            If WindowsSocket.State <> 7 Then GoTo ErrorHandler
            If SentSize > CheckSize Then
                ProgressTransfer.value = ProgressTransfer.Max
                WriteLog "Status: 100% Complete. " & FormatKB(CheckSize) & " Sent."
                WindowsSocket.SendData Mid(FileBinary, 1, Len(FileBinary) - (SentSize - CheckSize))
            Else
                ProgressTransfer.value = (SentSize / CheckSize) * ProgressTransfer.Max
                WriteLog "Status: " & CInt((SentSize / CheckSize) * ProgressTransfer.Max) & "% Complete, " & FormatKB(CheckSize - SentSize) & " Remaining..."
                WindowsSocket.SendData FileBinary
            End If

            Do
                DoEvents
            Loop Until PacketWait = False

            DoEvents: Loop Until EOF(FileNumber)

            SendFile = True

    Close #FileNumber
    Exit Function

ErrorHandler:
    WriteLog "ERROR:" & Err.Description
    cmdclose.Enabled = True
    WindowsSocket.Close
    DoEvents
    Close #FileNumber

End Function


Private Sub WindowsSocket_SendComplete()
    PacketWait = False
End Sub

Private Sub WindowsSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    FileDataRecievedSoFar = 0
    FileInTransfer = False
    WriteLog "Error No. -261080 : Socket Error !!" & Description
    'WindowsSocket.Close
    
    DoEvents
    Close #FileNumber
End Sub


Public Sub TickTick(ByVal lngTicks As Long)
On Error GoTo errHandler
    Dim i As Long
    Dim strTick As String
    For i = 0 To lngTicks * 1000 Step 37
        strTick = Str(i)
        lblstatus = "Attempting for Connection with Server:"
        DoEvents
    Next
errHandler:
End Sub
