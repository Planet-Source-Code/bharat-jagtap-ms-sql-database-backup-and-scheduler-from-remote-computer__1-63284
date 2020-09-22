VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSQLDMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'
' This simple class will help You implementing an SQL Server (and MSDE)
' backup/restore, the class uses the object library SQLDMO to interact
' with the native SQLsrv backup/restore functions so You'll need to
' reference that object library.
'
' It must be noticed that (if backing up on a file), the specified path
' is interpreted as "local" to the machine where the SQL server resides
' unless the path is specified in UNC notation, so if from a client You
' specify "c:\sqlbk\somefile.bak", the file will be created into the
' "sqlbk" folder of the machine running the SQL server and *not* on the
' client machine, while using "\\client\sqlbk\somefile.bak" will achieve
' the needed result
'

' device
Private Const BACKUP_DEVICE = "DMO_BKP_DEV"

' properties
Private msSQLserver As String ' Server name/address
Private msSQLuser As String ' user ID
Private msSQLpassword As String ' Password
Private msDataBase As String ' db name
Private msBackupFile As String ' backup file (pathname)
Private msBackupName As String ' backup set name
Private msBackupDescription As String ' backup set description
Private mbReplaceDatabase As Boolean ' true=creates/recreates db

Private mcolServers As Collection ' SQL server list
Private mcolDataBases As Collection ' database list

' SQL DMO
Private WithEvents moSRV As sqldmo.SQLserver
Attribute moSRV.VB_VarHelpID = -1
Private WithEvents moBAK As sqldmo.Backup
Attribute moBAK.VB_VarHelpID = -1
Private WithEvents moRST As sqldmo.Restore
Attribute moRST.VB_VarHelpID = -1
Private moDEV As sqldmo.BackupDevice

' events
Public Event ServerMessage(ByVal sMessage As String)
Public Event BackupProgress(ByVal sMessage As String, ByVal lPercent As Long)
Public Event BackupCompleted(ByVal sMessage As String)
Public Event RestoreProgress(ByVal sMessage As String, ByVal lPercent As Long)
Public Event RestoreCompleted(ByVal sMessage As String)
Public Event InsertMedia(ByVal sMessage As String)
Public Event Failure(ByVal sSource As String, ByVal lCode As Long, ByVal sMessage As String)

' init
Private Sub Class_Initialize()
Set mcolServers = New Collection
Set mcolDataBases = New Collection
End Sub

' reset
Private Sub Class_Terminate()
On Local Error Resume Next
moBAK.Abort
Set moBAK = Nothing
moSRV.Disconnect
Set moSRV = Nothing
Set mcolDataBases = Nothing
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::
':: Properties
'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::

' SQL server
Public Property Let SQLserver(ByVal sName As String)
msSQLserver = sName
End Property

' user id
Public Property Let SQLuser(ByVal sUser As String)
msSQLuser = sUser
End Property

' password
Public Property Let SQLpassword(ByVal sPass As String)
msSQLpassword = sPass
End Property

' database
Public Property Let DataBase(ByVal sName As String)
msDataBase = sName
End Property

' backup file path\name
Public Property Let BackupFile(ByVal sPathName As String)
msBackupFile = sPathName
End Property

' backup set name
Public Property Let BackupName(ByVal sName As String)
msBackupName = sName
End Property

' backup set description
Public Property Let BackupDescription(ByVal sDescr As String)
msBackupDescription = sDescr
End Property

' Let restore create/recreate DB
Public Property Let ReplaceDatabase(ByVal bYesNo As Boolean)
mbReplaceDatabase = bYesNo
End Property

' Number of known servers
Public Property Get ServerCount() As Long
On Local Error Resume Next
ServerCount = mcolServers.Count
End Property

' N-th server name
Public Property Get ServerName(ByVal lIndex As Long) As String
On Local Error Resume Next
ServerName = mcolServers(lIndex)
End Property

' Number of existing DBs
Public Property Get DataBaseCount() As Long
On Local Error Resume Next
DataBaseCount = mcolDataBases.Count
End Property

' Name of N-th DB
Public Property Get DataBaseName(ByVal lIndex As Long) As String
On Local Error Resume Next
DataBaseName = mcolDataBases(lIndex)
End Property


'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::
':: Methods
'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::

' Loads servers list
Public Function LoadServerList() As Long
Dim oSL As NameList, iSRV As Integer

On Local Error GoTo Catch
Set mcolServers = New Collection
Set oSL = sqldmo.ListAvailableSQLServers()
For iSRV = 1 To oSL.Count
mcolServers.Add oSL(iSRV)
Next iSRV

BailOut:
On Local Error Resume Next
Set oSL = Nothing
LoadServerList = mcolServers.Count
Exit Function

Catch:
RaiseEvent Failure(Err.Source, Err.Number, Err.Description)
Set mcolServers = New Collection
Resume BailOut
End Function

' Loads DB list for a given server
Public Function LoadDBlist() As Long
Dim iDB As Long

On Local Error GoTo Catch
LoadDBlist = 0
Set mcolDataBases = New Collection
If PingServer() = False Then
Exit Function
End If

Set moSRV = New sqldmo.SQLserver
moSRV.Connect msSQLserver, msSQLuser, msSQLpassword
moSRV.Databases.Refresh
For iDB = 1 To moSRV.Databases.Count
mcolDataBases.Add moSRV.Databases(iDB).Name
Next iDB

BailOut:
On Local Error Resume Next
moSRV.Disconnect
LoadDBlist = mcolDataBases.Count
Exit Function

Catch:
RaiseEvent Failure(Err.Source, Err.Number, Err.Description)
Set mcolDataBases = New Collection
Resume BailOut
End Function

' Backup
Public Function Backup() As Boolean
Dim bRet As Boolean, sDev As String

On Local Error GoTo Catch
bRet = False

' Connect...
If PingServer() = False Then
Exit Function
End If
moSRV.Connect msSQLserver, msSQLuser, msSQLpassword

' Remove device
On Local Error Resume Next
sDev = BACKUP_DEVICE
moSRV.BackupDevices.Remove sDev

' Add device
On Local Error GoTo Catch
Set moDEV = New sqldmo.BackupDevice
With moDEV
.Name = sDev
.Type = SQLDMODevice_DiskDump
.PhysicalLocation = msBackupFile
End With
moSRV.BackupDevices.Add moDEV

' Run backup
If Len(msBackupName) < 1 Then
msBackupName = "BKP" & Year(Now) & Right("00" & Month(Now), 2) & Right("00" & Day(Now), 2)
End If
If Len(msBackupDescription) < 1 Then
msBackupDescription = "Backup " & msDataBase & " (" & Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ")"
End If
Set moBAK = New sqldmo.Backup
With moBAK
.DataBase = msDataBase
.Devices = sDev
.TruncateLog = SQLDMOBackup_Log_Truncate
.BackupSetName = msBackupName
.BackupSetDescription = msBackupDescription
.PercentCompleteNotification = 10
.SQLBackup moSRV
End With
bRet = True

BailOut:
On Local Error Resume Next
' cleanup
If bRet = False Then
moBAK.Abort
End If
Set moBAK = Nothing
moSRV.BackupDevices.Remove sDev
moSRV.Disconnect
Set moSRV = Nothing
Backup = bRet
Exit Function

Catch:
' error !
bRet = False
RaiseEvent Failure(Err.Source, Err.Number, Err.Description)
Resume BailOut
End Function

' Restore
Public Function Restore() As Boolean
Dim bRet As Boolean, sDev As String

On Local Error GoTo Catch
bRet = False

' Connect to server
If PingServer() = False Then
Exit Function
End If
moSRV.Connect msSQLserver, msSQLuser, msSQLpassword

' Remove device
On Local Error Resume Next
sDev = BACKUP_DEVICE
moSRV.BackupDevices.Remove sDev

' Add device
On Local Error GoTo Catch
Set moDEV = New sqldmo.BackupDevice
With moDEV
.Name = sDev
.Type = SQLDMODevice_DiskDump
.PhysicalLocation = msBackupFile
End With
moSRV.BackupDevices.Add moDEV

' Restore
Set moRST = New sqldmo.Restore
With moRST
.DataBase = msDataBase
.Devices = sDev
.ReplaceDatabase = mbReplaceDatabase
.PercentCompleteNotification = 10
.SQLRestore moSRV
End With
bRet = True

BailOut:
On Local Error Resume Next
' cleanup
If bRet = False Then
moRST.Abort
End If
Set moRST = Nothing
moSRV.BackupDevices.Remove sDev
moSRV.Disconnect
Set moSRV = Nothing
Restore = bRet
Exit Function

Catch:
' error !
bRet = False
RaiseEvent Failure(Err.Source, Err.Number, Err.Description)
Resume BailOut
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::
':: Service function (class internal use)
'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::

' creates an instance of the "SQLServer" object and verifies
' if that server is reachable using the "SQLping" function
Private Function PingServer() As Boolean
Dim vVer As SQLDMO_SQL_VER
Dim bRet As Boolean

On Local Error Resume Next
bRet = False
Set moSRV = New sqldmo.SQLserver
vVer = moSRV.PingSQLServerVersion(msSQLserver, msSQLuser, msSQLpassword)
If vVer <> SQLDMOSQLVer_Unknown Then
bRet = True
Else
Set moSRV = Nothing
End If
If bRet = False Then
RaiseEvent Failure("SQLbackup", -1, "Can't " & Chr(34) & "ping" & Chr(34) & " specified server (" & msSQLserver & ")")
End If
PingServer = bRet
End Function

'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::
':: Events from DMO object, reflected to class consumers
'::::::::::::::::::::::::::::::::::::::::::::::::: ::::::::::::::::::::::

' backup complete
Private Sub moBAK_Complete(ByVal Message As String)
RaiseEvent BackupCompleted(Message)
End Sub

' new media
Private Sub moBAK_NextMedia(ByVal Message As String)
RaiseEvent InsertMedia(Message)
End Sub

' % complete
Private Sub moBAK_PercentComplete(ByVal Message As String, ByVal Percent As Long)
RaiseEvent BackupProgress(Message, Percent)
End Sub

' restore complete
Private Sub moRST_Complete(ByVal Message As String)
RaiseEvent RestoreCompleted(Message)
End Sub

' new media
Private Sub moRST_NextMedia(ByVal Message As String)
RaiseEvent InsertMedia(Message)
End Sub

' % complete
Private Sub moRST_PercentComplete(ByVal Message As String, ByVal Percent As Long)
RaiseEvent RestoreProgress(Message, Percent)
End Sub

' command
Private Sub moSRV_CommandSent(ByVal SQLCommand As String)
RaiseEvent ServerMessage("Command sent: " & SQLCommand)
End Sub

' connection broken
Private Function moSRV_ConnectionBroken(ByVal Message As String) As Boolean
RaiseEvent Failure("SQLDMO", 0, "Connection Broken: " & Message)
End Function

' query timeout
Private Function moSRV_QueryTimeout(ByVal Message As String) As Boolean
RaiseEvent Failure("SQLDMO", 0, "Query Timeout: " & Message)
End Function

' login error
Private Sub moSRV_RemoteLoginFailed(ByVal Severity As Long, ByVal MessageNumber As Long, ByVal MessageState As Long, ByVal Message As String)
RaiseEvent Failure("SQLDMO", MessageNumber, "Remote Login Failed: " & Message)
End Sub

' message/warning/error
Private Sub moSRV_ServerMessage(ByVal Severity As Long, ByVal MessageNumber As Long, ByVal MessageState As Long, ByVal Message As String)
RaiseEvent ServerMessage(Message)
End Sub

