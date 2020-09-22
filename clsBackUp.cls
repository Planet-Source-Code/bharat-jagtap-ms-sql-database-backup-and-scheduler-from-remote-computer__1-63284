VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const BACKUP_DEVICE = "DMO_BKP_DEV"
Public Function backupDB(sDBaseName As String, sPath As String, gsUserName As String, gsPaswword As String) As Boolean
'Author:  sangram
'Purpose: Database Back
On Error GoTo ErrHnd
    Dim rstServerName As New ADODB.Recordset
    Dim oSQLServer As New sqldmo.SQLserver
    Dim oBackup As New sqldmo.Backup
    Dim sServerName As String
    AdoConnection
    rstServerName.Open "SELECT @@SERVERNAME ", gdbConnection, adOpenStatic, adLockReadOnly
    sServerName = rstServerName.Fields(0)
    Set rstServerName = Nothing
    oSQLServer.Connect sServerName, gsUserName, gsPaswword
    oBackup.Action = SQLDMOBackup_Database
    oBackup.DataBase = sDBaseName
    oBackup.BackupSetDescription = "Complete Backup of " & sDBaseName
    'oBackup.Files = "\\pritesh" & "\c:\" & "e26c10a80stmp.bak" 'gettempp
    
    oBackup.Files = GetTempPathName & "e26c10a80stmp.bak"  'gettempp
    oBackup.SQLBackup oSQLServer
    oSQLServer.Disconnect
    backupDB = True
    Exit Function
ErrHnd:
    backupDB = False
End Function

Public Static Function GetServerWindowsName(ByVal strPwd) As String
On Error GoTo ErrHand
    Dim rstServerName As New ADODB.Recordset
    Dim oSQLServer As New sqldmo.SQLserver
    
    Dim sServerName As String
    AdoConnection
    rstServerName.Open "SELECT @@SERVERNAME ", gdbConnection, adOpenStatic, adLockReadOnly
    sServerName = rstServerName.Fields(0)
    
    oSQLServer.Connect UCase(sServerName), "SA", g_sPassword
     
     GetServerWindowsName = oSQLServer.NetName
    Set rstServerName = Nothing
    Exit Function

ErrHand:
    MsgBox Err.Description
End Function
Private Function GetTempPathName() As String
    Dim sbuffer As String
    Dim lRet As Long
    
    sbuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sbuffer)
    
    If lRet > 0 Then
        sbuffer = Left$(sbuffer, lRet)
    End If
    GetTempPathName = sbuffer
    
End Function
