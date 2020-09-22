VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Configuration"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3690
      TabIndex        =   6
      Top             =   1710
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   5
      Top             =   1665
      Width           =   915
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   5040
      Begin VB.TextBox txtservername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtpwd 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1275
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "This#is(for)system^admin~05"
         Top             =   840
         Width           =   3570
      End
      Begin VB.Label lblpwd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   885
         Width           =   975
      End
      Begin VB.Label lblserver 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
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
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'getIP
 Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

' Return the local host's name.
Private Function LocalHostName() As String
Dim hostname As String * 256

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        LocalHostName = "<Error>"
    Else
        LocalHostName = Trim$(hostname)
    End If
End Function
Private Sub InitializeSockets()
Dim WSAD As WSADATA
Dim iReturn As Integer
Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If

    'iMaxSockets is not used in winsock 2. So the following check is only
    'necessary for winsock 1. If winsock 2 is requested,
    'the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If

End Sub

Private Sub CleanupSockets()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If

End Sub
Private Function IPAddressFromHostName(ByVal hostname As String) As String
Dim hostent_addr As Long
Dim host As HOSTENT
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String
Dim result As String

    hostent_addr = gethostbyname(hostname)
    If hostent_addr = 0 Then
        IPAddressFromHostName = "<error>"
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    ' Get multiple pieces of the IP address
    ' if machine is multi-homed.
    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next

        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        result = result & ip_address & vbCrLf
        ip_address = ""

        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)

    ' Remove the last vbCrLf.
    If Len(result) > 0 Then result = Left$(result, Len(result) - Len(vbCrLf))

    IPAddressFromHostName = result
End Function
Public Sub CreateTableAppParam()
On Error GoTo ErrHand
    Dim strSql As String
    
    AdoConnection1
    strSql = "if not exists (select * from dbo.sysobjects  " _
        & " where id = object_id(N'[dbo].[AppParam]') and  " _
        & " OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
        & " CREATE TABLE [AppParam] ( " _
        & " [ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL , " _
        & " [AppKey] [varchar] (50)  NULL , " _
        & " [AppValue] [varchar] (50) NULL , " _
        & " [CreatedBy] [varchar] (50) NULL , " _
        & " [Version] [varchar] (50) NULL , " _
        & " [Purpose] [varchar] (50) NULL , " _
        & " [CreatedDate] [datetime] NULL , " _
        & " [UpdatedDate] [datetime] NULL " _
        & " ) ON [PRIMARY] "
    
    gdbConnection.Execute strSql
    
    strSql = " if not exists (select * from dbo.sysobjects  " _
            & " where id = object_id(N'[dbo].[AppParam]') and  " _
            & " OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
            & " ALTER TABLE [dbo].[AppParam] WITH NOCHECK ADD " _
            & " CONSTRAINT [DF_AppParam_CreatedDate] DEFAULT (getdate()) FOR [CreatedDate], " _
            & " CONSTRAINT [DF_AppParam_UpdatedDate] DEFAULT (getdate()) FOR [UpdatedDate] "
    
    gdbConnection.Execute strSql
    gdbConnection.Close
Exit Sub
ErrHand:
    MsgBox Err.Description

End Sub

Public Sub SetServerIP()
On Error GoTo ErrHand
Dim strSql As String
    
    AdoConnection1
    CreateTableAppParam
    AdoConnection1
    strSql = " Delete FROM AppParam WHERE (AppKey = 'ServerIP')"
    gdbConnection.Execute strSql
    
    strSql = " if not exists (SELECT  AppKey FROM AppParam WHERE (AppKey = 'ServerIP'))" _
            & " INSERT INTO AppParam (AppValue, AppKey, CreatedBy, Version, Purpose) " _
            & " VALUES  ('" & IPAddressFromHostName(LocalHostName()) & "', 'ServerIP', 'Bharat Jagtap', '1.1.6', 'JB SQL Backup Scheduler')"
    
    gdbConnection.Execute strSql
    gdbConnection.Close
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

''end get ip
 

Public Function AdoConnection1() As Boolean
On Error GoTo ErrHnd
   ' gdbConnection.Close
    Set gdbConnection = New ADODB.Connection
    If gdbConnection.State = adStateOpen Then gdbConnection.Close
        Set gdbConnection = New ADODB.Connection
        If Len(Trim(g_sPassword)) <> 0 Then
              gdbConnection.Open "Provider=SQLOLEDB.1;Password='" & g_sPassword & "';Persist Security Info=True;User ID=sa;Initial Catalog=CAS;Data Source=" & txtservername
        Else
              gdbConnection.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=MASTER;Data Source=" & g_sServerName
        End If
    
    If gdbConnection.State = 1 Then
        AdoConnection1 = True
    Else
        AdoConnection1 = False
    End If
    Exit Function
ErrHnd:
    AdoConnection1 = False
    If InStr(1, Err.Description, "SQL Server does not exist", vbTextCompare) > 1 Then
        MsgBox "SQL Server does not exist or access denied.", vbCritical
    Else
        MsgBox Err.Description
    End If
    On Error Resume Next
End Function

Private Sub Command1_Click()
txtpwd = txtpwd.Tag
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHand
    Dim strLocalHost As String
    Dim strRemoteHost As String
    
    
    If Trim(txtservername) = "" Then
        MsgBox "Please Enter the Server Name"
        txtservername.SetFocus
        Exit Sub
    End If
    
    g_sServerName = Trim(txtservername)
    g_sPassword = Trim(txtpwd)
    g_sDatabase = "Master"
    
    strLocalHost = Trim$(UCase(LocalHostName()))
    strRemoteHost = GetServerWindowsName(g_sPassword)
    Text1.Text = strLocalHost
    strLocalHost = Trim$(Text1.Text)
    
    
    
    If UCase(strLocalHost) <> UCase("<Error>") And UCase(strRemoteHost) <> UCase("<Error>") Then
        If UCase(GetServerWindowsName(txtpwd)) <> strLocalHost Then
            MsgBox "Database Server is Hosted on Another Computer." & vbCrLf & "The Database Server And JBackup  Backup Server Should Be Hosted on Same Computer", vbCritical
            txtservername.SetFocus
            Exit Sub
        End If
        
        AdoConnection1
        SetServerIP
        SaveSetting "ABSOLUTE NetWare", "NetSoft", "DBServer", "Config"
    Else
       ' MsgBox "Unexpected Error Occured.."
    End If
    If gdbConnection.State Then
        gdbConnection.Close
    End If
    Unload Me
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
'get ip
    ' Initialize the sockets library.
    InitializeSockets
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean up the sockets library.
    CleanupSockets
End Sub
Public Static Function GetServerWindowsName(ByVal strPwd) As String
On Error GoTo ErrHand
    Dim rstServerName As New ADODB.Recordset
    Dim oSQLServer As New sqldmo.SQLserver
    Dim hostname As String
    
    Dim sServerName As String
    If AdoConnection1 Then
    rstServerName.Open "SELECT @@SERVERNAME ", gdbConnection, adOpenStatic, adLockReadOnly
    sServerName = rstServerName.Fields(0)
    
    oSQLServer.Connect UCase(sServerName), "SA", UCase(g_sPassword)
    hostname = oSQLServer.NetName
    GetServerWindowsName = hostname
    oSQLServer.Disconnect
    Set rstServerName = Nothing
    Exit Function
    Else
        GetServerWindowsName = "<ERROR>"
    End If
    Exit Function
ErrHand:
    MsgBox Err.Description
End Function
