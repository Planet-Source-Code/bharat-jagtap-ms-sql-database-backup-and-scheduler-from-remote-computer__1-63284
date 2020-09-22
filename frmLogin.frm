VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7350
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   3225
      TabIndex        =   33
      Top             =   5640
      Width           =   3255
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
         TabIndex        =   36
         Top             =   120
         Width           =   975
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
         TabIndex        =   35
         Top             =   600
         Width           =   2505
      End
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
         TabIndex        =   34
         Top             =   360
         Width           =   2865
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800000&
      Caption         =   "Backup Location"
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
      Height          =   660
      Left            =   3435
      TabIndex        =   28
      Top             =   1995
      Width           =   5040
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "&Browse"
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
         Left            =   4005
         TabIndex        =   30
         Top             =   270
         Width           =   960
      End
      Begin VB.TextBox txtBackupPath 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Backup Schedule"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   3420
      TabIndex        =   15
      Top             =   2760
      Width           =   5070
      Begin VB.ListBox LstWeekday 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "frmLogin.frx":000C
         Left            =   1560
         List            =   "frmLogin.frx":0025
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPTime 
         Height          =   330
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22740994
         CurrentDate     =   38651
      End
      Begin VB.PictureBox PicHourly 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1155
         ScaleHeight     =   495
         ScaleWidth      =   2700
         TabIndex        =   23
         Top             =   1680
         Width           =   2700
         Begin VB.ComboBox CmbHrs 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmLogin.frx":0069
            Left            =   915
            List            =   "frmLogin.frx":00B2
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hour(s)"
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
            Left            =   1740
            TabIndex        =   26
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Every"
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
            Left            =   240
            TabIndex        =   24
            Top             =   180
            Width           =   570
         End
      End
      Begin VB.OptionButton OptHourly 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Hourly"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton optdaily 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Daily"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton OptMonthly 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Monthly"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton OptWeekly 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Weekly"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1065
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   3675
         TabIndex        =   27
         Top             =   1200
         Width           =   3675
      End
      Begin MSComCtl2.MonthView mvwMonth 
         Height          =   2370
         Left            =   1155
         TabIndex        =   18
         Top             =   1230
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   8388608
         BackColor       =   8388608
         Appearance      =   1
         StartOfWeek     =   22740993
         CurrentDate     =   38651
      End
   End
   Begin VB.CommandButton cmdBackupNow 
      Caption         =   "&Take Backup Now !!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      TabIndex        =   14
      Top             =   6600
      Width           =   2400
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800000&
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
      Height          =   750
      Left            =   3420
      TabIndex        =   10
      Top             =   6555
      Width           =   5040
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1305
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton cmdclose 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800000&
      Caption         =   "Backup Type"
      Enabled         =   0   'False
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
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   1995
      Visible         =   0   'False
      Width           =   4905
      Begin VB.OptionButton optComplete 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Complete"
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
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptDifferential 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Differential"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
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
      Height          =   1575
      Left            =   3435
      TabIndex        =   2
      Top             =   75
      Width           =   5040
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1290
         TabIndex        =   31
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtservername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtpwd 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1275
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
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
         Left            =   240
         TabIndex        =   32
         Top             =   1125
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
         TabIndex        =   7
         Top             =   360
         Width           =   855
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
         TabIndex        =   6
         Top             =   765
         Width           =   975
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3645
      TabIndex        =   1
      Top             =   405
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   -45
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cdlgOpen 
      Left            =   405
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Database Backup"
      Filter          =   "*.Bak"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3510
      Picture         =   "frmLogin.frx":0109
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   5880
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   1995
      Left            =   1395
      Picture         =   "frmLogin.frx":0413
      Stretch         =   -1  'True
      Top             =   585
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   315
      Picture         =   "frmLogin.frx":057E
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Database Backup Scheduler"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   450
      TabIndex        =   13
      Top             =   3915
      Width           =   2715
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private WithEvents s As ClsSQLDMO
Attribute s.VB_VarHelpID = -1
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private mstrFileToWrite As String
Private SetServerName As String 'servername
Private SetPwd As String 'password
Private SetType As String 'type i.e. complete or differential
Private SetFrequency As String 'frequency i.e. daily, weekly,monthly
Private SetPath As String 'default backup path
Private SetDate As String
Private SetWeekDays As String
'Private ObjCipher As New clsCipher
Private fso As New FileSystemObject
Private strFrqType As String
Private strFrqDay As String
Private strFrqHH As String
Private strFrqMM As String
Private strFrqWeekDays As String
Public strFrqBackupPath As String
Private strFrequency As String
Public strServerIP As String
Public strZipFileName As String
Private strHrTime As String






Private Function LargeIntegerToDouble(low_part As Long, high_part As Long) As Double
Dim result As Double

On Error GoTo ErrHand
    result = high_part
    If high_part < 0 Then result = result + 2 ^ 32
    result = result * 2 ^ 32

    result = result + low_part
    If low_part < 0 Then result = result + 2 ^ 32

    LargeIntegerToDouble = result
    Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Public Function Backup() As Boolean
On Error GoTo ErrHand
    Dim db As clsBackUp
    Dim fileEntry() As String
        Dim vbCompare As String
            strBackupFileName = "\" & g_sDatabase & "_" & "BK_" & Format(Now, "dd_mm_yy_hh_mm_ss")
            Screen.MousePointer = 11
            mstrFileToWrite = InPath & strBackupFileName
            strZipFileName = mstrFileToWrite & ".zip"
            mstrFileToWrite = mstrFileToWrite & ".bak"
            gsBackUpVariable = 2
            
            Set db = New clsBackUp
            If db.backupDB(g_sDatabase, strBackupFileName & ".bak", "sa", g_sPassword) = True Then
                Call FileBackupZip
                'MsgBox "Backup has completed successfully.", vbInformation + vbOKOnly
                Backup = True
                
                
            Else
                'MsgBox "Backup has not completed successfully.", vbInformation + vbOKOnly
                DeleteSetting "ABSOLUTE NetWare", "NetSoft"
                Backup = False
            End If
            
            Set db = Nothing
            Screen.MousePointer = 0
            CloseConnection
        Exit Function
ErrHand:
    FrmServer.WriteLog Err.Number & " " & Err.Description
    Screen.MousePointer = 0
   
End Function

Public Function GetTempPathName() As String
On Error GoTo ErrHand
    Dim sbuffer As String
    Dim lRet As Long
    
    sbuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sbuffer)
    
    If lRet > 0 Then
        sbuffer = Left$(sbuffer, lRet)
    End If
    GetTempPathName = sbuffer
    Exit Function

ErrHand:
    FrmServer.WriteLog Err.Description
End Function

Public Function FileBackupZip() As Integer
On Error GoTo errHandler

Dim m_cZ As clsZip
Dim fout As New FileSystemObject
Dim mintret As Double
Dim mstrdrivepath As String
Dim strTmpUserFile As String
Dim strTmpBackFile As String
Dim strTmpPath As String
    Set m_cZ = New clsZip
    mstrdrivepath = InPath & strBackupFileName
     FrmServer.TxtSaveFile = mstrdrivepath & ".zip"
           '*********Zip a file creation*********
     strTmpPath = GetTempPathName
            strTmpBackFile = strTmpPath & "\e26c10a80stmp.bak"
            strTmpUserFile = strTmpPath & "\" & fout.GetFileName(mstrFileToWrite)
            '*********Zip a file creation*********
            If fout.FileExists(strTmpUserFile) Then
                fout.DeleteFile (strTmpUserFile)
            End If
            
            fout.MoveFile strTmpBackFile, strTmpUserFile
                   '*********Zip a file creation*********
    
            With m_cZ
                .Encrypt = True
                .ZipFile = mstrdrivepath & ".zip"
                .StoreFolderNames = False
                .ClearFileSpecs
                .AddFileSpec strTmpUserFile
                .Zip
                'for Error Msg
                If (.Success) Then
                    FileBackupZip = 1
                Else
                    FileBackupZip = 2
                End If
            End With
            
        DoEvents
            fout.DeleteFile strTmpUserFile
            Set fout = Nothing
            CloseConnection
        Exit Function

errHandler:
    FrmServer.WriteLog Err.Number & " " & Err.Description

End Function

Public Sub ShowBusyScreen(nBusyCaption As String)
    frmBusy.lblWait.Caption = nBusyCaption
    frmBusy.Show
    DoEvents
End Sub

Public Sub UnloadBusyScreen()
    Unload frmBusy
End Sub

Private Sub cmdBackupNow_Click()
    If Trim(g_sServerName) = "" Then
        MsgBox "Server Name Can Not Be Empty!"
        Exit Sub
    End If
    
Call FrmTransfer1.TickTick(1000)
Call RequestServer4BackUp
End Sub

 

Private Sub CmdBrowse_Click()
On Error GoTo ErrHand
    FX_MSchart.Show vbModal
    txtBackupPath = strFrqBackupPath
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
    
    If gdbConnection.State Then
        gdbConnection.Close
    End If

Me.Hide
End Sub

Private Sub CmdOk_Click()
Dim strEncryFrq As String
Dim strEncryServer As String
Dim strEncryPassword As String
Dim strEncryBackupType As String
Dim strEncryFrequency As String
Dim strEncryPath As String
Dim strEncryDate As String
Dim strEncryWeekdays As String
Dim strEncryWinSvr As String
Dim strEncryBackupPath As String

Timer1.Enabled = False
On Error GoTo ErrHand
    
    If Trim(txtservername) = "" Then
        MsgBox "Server Name Can Not Be Empty!"
        Exit Sub
    End If
    
    
    If Trim(txtBackupPath) = "" Then
        MsgBox "Please Set the Backup Folder!"
        txtBackupPath.SetFocus
        Exit Sub
    End If
    
    If OptHourly Then
        If CmbHrs.Text = "" Then
            MsgBox "Please Set the Hours Frequency for Backup.", vbCritical
            CmbHrs.SetFocus
            Exit Sub
        End If
    End If
    
    MousePointer = vbHourglass
    
    g_sServerName = txtservername
    g_sDatabase = txtDatabase
    g_sPassword = txtpwd
    
    Call CreateSchedule
    
    If AdoConnection = True Then
           Call GetServerWindowsName
           Call GetServerIP

          strEncryServer = Trim(UCase(g_sServerName))
          strEncryPassword = Trim(UCase(g_sPassword))
          strEncryBackupType = Trim(UCase("Complete"))
          strEncryFrequency = Trim(UCase(tempFrequency))
          strEncryPath = Trim(UCase(InPath))
          strEncryDate = Trim(UCase(date))
          strEncryFrq = Trim(UCase(strFrequency))
          strEncryWinSvr = Trim(UCase(CodeModule.strServerWindowsName))
          strEncryWeekdays = Trim(UCase(strFrqWeekDays))
          strEncryBackupPath = Trim(UCase(strFrqBackupPath))
          
          
          'savesetting
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NetWare", strEncryServer 'servername
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "BizNet", strEncryPassword 'password
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NetShape", strEncryBackupType ' backup type
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "MacroNet", strEncryFrequency ' backup frequency d/w/m
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NetWin", strEncryPath 'backup path
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NetDomain", strEncryDate 'start date
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NEO", strEncryFrq  'BkupSchedule
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "NTService", strEncryWinSvr  'winserver
          'weekdays
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "IO", strEncryWeekdays  'weekdays
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "BKP", strEncryBackupPath  'backup path
          SaveSetting "ABSOLUTE NetWare", "NetSoft", "DB", g_sDatabase 'DB
          
          Timer1.Enabled = True
    End If
    MsgBox "Schedule Saved Successfully!!", vbInformation
    MousePointer = vbDefault
    Exit Sub
ErrHand:
    MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
Form_Load
End Sub

Private Sub Form_GotFocus()
Form_Load
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand
    
 Call IsApplicationAlreadyInstenciated
 
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\ABSOLUTE\Windows\CurrentVersion\Run", App.EXEName, App.Path & "\" & App.EXEName & ".exe"
 'We set up Picture1 to accept callback data
'in it's MouseMove procedure.
If strHrTime = "" Then
    strHrTime = FormatDateTime(Now(), vbShortTime)
End If

NotifyIcon.cbSize = Len(NotifyIcon)
NotifyIcon.hWnd = Picture1.hWnd
NotifyIcon.uID = 1&
NotifyIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
NotifyIcon.uCallbackMessage = WM_MOUSEMOVE

'Now, we set up the icon and tool tip message
NotifyIcon.hIcon = Picture1.Picture
NotifyIcon.szTip = "JBackup Scheduler" & Chr$(0)

'Lastly, we add the icon
Shell_NotifyIcon NIM_ADD, NotifyIcon

''''''''''''''''''''''''''''''
Timer1.Enabled = True
    
    ''''''''''''''''''''''''''''''
    optComplete.value = True
    optdaily.value = True
    g_sServerName = GetSetting("ABSOLUTE NetWare", "NetSoft", "Netware")
    g_sPassword = GetSetting("ABSOLUTE NetWare", "NetSoft", "BizNet")
    g_sDatabase = GetSetting("ABSOLUTE NetWare", "NetSoft", "DB")
   ' SetType = GetSetting("ABSOLUTE NetWare", "NetSoft", "NetShape")
   ' SetFrequency = GetSetting("ABSOLUTE NetWare", "NetSoft", "MacroNet")
    SetPath = GetSetting("ABSOLUTE NetWare", "NetSoft", "NetWin")
   ' SetDate = GetSetting("ABSOLUTE NetWare", "NetSoft", "NetDomain")
    strFrqWeekDays = GetSetting("ABSOLUTE NetWare", "NetSoft", "IO")
    strFrqBackupPath = GetSetting("ABSOLUTE NetWare", "NetSoft", "BKP")
    strFrequency = GetSetting("ABSOLUTE NetWare", "NetSoft", "NEO")
    CodeModule.strServerWindowsName = GetSetting("ABSOLUTE NetWare", "NetSoft", "NTService")
    ''''''''''''''''
    'enable timer only if reg values are set
    Call GetSchedule
    InitComponents
    CloseConnection
    Exit Sub

ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, NotifyIcon

End Sub

Private Sub optdaily_Click()
    LstWeekday.Visible = False
    mvwMonth.Visible = False
    DTPTime.Visible = True
    PicHourly.Visible = False
End Sub

Private Sub OptHourly_Click()
    LstWeekday.Visible = False
    mvwMonth.Visible = False
    DTPTime.Visible = False
    PicHourly.Visible = True
End Sub

Private Sub OptMonthly_Click()
    LstWeekday.Visible = False
    mvwMonth.Visible = True
    DTPTime.Visible = True
    PicHourly.Visible = False
End Sub

Private Sub OptWeekly_Click()
    LstWeekday.Visible = True
    mvwMonth.Visible = False
    DTPTime.Visible = True
    PicHourly.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrHand
    If Hex(x) = "1E0F" Then
       'Use "1E3C" for right-click
       'MsgBox ("JBackup  SQL BackUp.."), 48, ("JBackup  BackUp Utility")
     
       Me.Show
       Form_Load
    End If
    Exit Sub

ErrHand:
  If InStr(1, Err.Description, "modal") > 1 Then
  Else
  MsgBox Err.Description
  End If
End Sub

Private Sub s_ServerMessage(ByVal sMessage As String)
MsgBox sMessage
End Sub

Public Sub SetBackUpPath()
Dim bytes_avail As LARGE_INTEGER
Dim bytes_total As LARGE_INTEGER
Dim bytes_free As LARGE_INTEGER
Dim dbl_total As Double
Dim dbl_free As Double
Dim i As Integer

On Error GoTo ErrHand
    For i = 0 To Drive1.ListCount
      If Drive1.List(Drive1.ListIndex) <> "a:" Then
         GetDiskFreeSpaceEx Drive1.List(Drive1.ListIndex), bytes_avail, bytes_total, bytes_free
    
        ' Convert values into Doubles.
        dbl_total = LargeIntegerToDouble(bytes_total.lowpart, bytes_total.highpart)
        dbl_free = (LargeIntegerToDouble(bytes_free.lowpart, bytes_free.highpart) / 1024) / 1024
          'chk that atleast 1 mb free space is needed
          If dbl_free >= 50 Then
           If fso.FolderExists(Drive1.List(Drive1.ListIndex) & "\_JBackup BackUp") = False Then
             fso.CreateFolder (Drive1.List(Drive1.ListIndex) & "\_JBackup BackUp")
             InPath = Drive1.List(Drive1.ListIndex) & "\_JBackup BackUp" 'backuppath
           Else
             InPath = Drive1.List(Drive1.ListIndex) & "\_JBackup BackUp" 'backuppath
             Exit For
           End If
            Exit For
          End If
      End If
    Next i
Exit Sub

ErrHand:
    MsgBox Err.Description

End Sub
' Return a formatted string representing the number of
' bytes.
Private Function FormatBytes(ByVal num_bytes As Double) As _
    String
Const ONE_KB As Double = 1024
Const ONE_MB As Double = ONE_KB * 1024
Const ONE_GB As Double = ONE_MB * 1024
Const ONE_TB As Double = ONE_GB * 1024
Const ONE_PB As Double = ONE_TB * 1024
Const ONE_EB As Double = ONE_PB * 1024
Const ONE_ZB As Double = ONE_EB * 1024
Const ONE_YB As Double = ONE_ZB * 1024
Dim value As Double
Dim txt As String

On Error GoTo ErrHand
    ' See how big the value is.
    If num_bytes <= 999 Then
        ' Format in bytes.
        FormatBytes = Format$(num_bytes, "0") & " bytes"
    ElseIf num_bytes <= ONE_KB * 999 Then
        ' Format in KB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_KB) & " KB"
    ElseIf num_bytes <= ONE_MB * 999 Then
        ' Format in MB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_MB) & " MB"
    ElseIf num_bytes <= ONE_GB * 999 Then
        ' Format in GB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_GB) & " GB"
    ElseIf num_bytes <= ONE_TB * 999 Then
        ' Format in TB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_TB) & " TB"
    ElseIf num_bytes <= ONE_PB * 999 Then
        ' Format in PB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_PB) & " PB"
    ElseIf num_bytes <= ONE_EB * 999 Then
        ' Format in EB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_EB) & " EB"
    ElseIf num_bytes <= ONE_ZB * 999 Then
        ' Format in ZB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_ZB) & " ZB"
    Else
        ' Format in YB.
        FormatBytes = ThreeNonZeroDigits(num_bytes / _
            ONE_YB) & " YB"
    End If
    Exit Function
    
ErrHand:
        MsgBox Err.Description
End Function

' Return the value formatted to include at most three
' non-zero digits and at most two digits after the decimal
' point. Examples:
'         1
'       123
'        12.3
'         1.23
'         0.12
Private Function ThreeNonZeroDigits(ByVal value As Double) _
    As String
On Error GoTo ErrHand
    If value >= 100 Then
        ' No digits after the decimal.
        ThreeNonZeroDigits = Format$(CInt(value))
    ElseIf value >= 10 Then
        ' One digit after the decimal.
        ThreeNonZeroDigits = Format$(value, "0.0")
    Else
        ' Two digits after the decimal.
        ThreeNonZeroDigits = Format$(value, "0.00")
    End If
    Exit Function
    
ErrHand:
    MsgBox Err.Description
End Function


Public Function CreateSchedule() As String
On Error GoTo ErrHand
    If optdaily.value = True Then
        strFrqType = "D"
        strFrqDay = 1
        strFrqHH = Hour(DTPTime)
        strFrqMM = Minute(DTPTime)
        
    ElseIf OptHourly.value = True Then
        strFrqType = "H"
        strFrqDay = Val(CmbHrs.Text)
    
    ElseIf OptWeekly.value = True Then
        strFrqType = "W"
        strFrqDay = Getweekday(LstWeekday.Text)
        strFrqHH = Hour(DTPTime)
        strFrqMM = Minute(DTPTime)
        strFrqWeekDays = getWeekDaysList()
    ElseIf OptMonthly.value = True Then
        strFrqType = "M"
        strFrqDay = Day(mvwMonth)
        strFrqHH = Hour(DTPTime)
        strFrqMM = Minute(DTPTime)
    End If
    
    strFrequency = strFrqType & "," & _
                    strFrqDay & "," & _
                    strFrqHH & "," & _
                    strFrqMM
    strFrqBackupPath = Trim$(txtBackupPath)
    CreateSchedule = strFrequency
Exit Function

ErrHand:
    MsgBox Err.Description
    
End Function

Public Function Getweekday(ByVal strWKdy As String) As Integer
On Error GoTo ErrHand
    Select Case UCase(strWKdy)
        Case "SUNDAY"
            Getweekday = 1
        Case "MONDAY"
            Getweekday = 2
        Case "TUESDAY"
            Getweekday = 3
        Case "WEDNESDAY"
            Getweekday = 4
        Case "THURSDAY"
            Getweekday = 5
        Case "FRIDAY"
            Getweekday = 6
        Case "SATURDAY"
            Getweekday = 7
    End Select
Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Public Sub InitComponents()
On Error GoTo ErrHand
    Dim strArr() As String
    If g_sServerName <> "" Then
        If strFrequency <> "" Then
            strArr = Split(strFrequency, ",")
        End If
        
        txtservername = g_sServerName
        txtpwd = g_sPassword
        txtDatabase = g_sDatabase
            
        If SetType = UCase("Complete") Then
            optComplete = True
        Else
            OptDifferential = True
        End If
        
        If strArr(0) = "D" Then
            optdaily = True
            'LstWeekday = WeekdayName(strArr(2))
            DTPTime = strArr(2) & ":" & strArr(3)
            mvwMonth = date
        ElseIf strArr(0) = "W" Then
            OptWeekly = True
           ' LstWeekday.Text = WeekdayName(strArr(1))
            If Trim(strFrqWeekDays) <> 0 Then
                Call FillWeekList(strFrqWeekDays)
            End If
            DTPTime = strArr(2) & ":" & strArr(3)
            mvwMonth = date
        ElseIf strArr(0) = "M" Then
            OptMonthly = True
            mvwMonth = strArr(1) & "/" & Format(date, "MMM/yyyy")
            DTPTime = strArr(2) & ":" & strArr(3)
        ElseIf strArr(0) = "H" Then
            OptHourly = True
            mvwMonth = IIf(strArr(1) = 0, 1, strArr(1)) & "/" & Format(date, "MMM/yyyy")
            DTPTime = Val(strArr(1)) & ":" & Val(strArr(3))
            CmbHrs = IIf(Val(strArr(1)) = 0, 1, Val(strArr(1)))
        End If
    End If
    txtBackupPath = strFrqBackupPath
    g_strBackupPath = txtBackupPath
    Exit Sub

ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub FillWeekList(ByVal strWeekdays As String)
Dim i As Integer
Dim j As Integer
Dim iNoOfDays As Integer
Dim strArrWeekdays() As String


strArrWeekdays = Split(strWeekdays, ",")
iNoOfDays = UBound(strArrWeekdays)

    For j = 0 To iNoOfDays
        For i = 0 To 6
            If UCase(LstWeekday.List(i)) = UCase(strArrWeekdays(j)) Then
                LstWeekday.Selected(i) = True
            End If
        Next i
    Next j
End Sub
Private Function getWeekDaysList() As String
Dim i As Integer
Dim strWkdays  As String

        For i = 0 To LstWeekday.ListCount - 1
            If LstWeekday.Selected(i) Then
                strWkdays = strWkdays & LstWeekday.List(i) & ","
            End If
        Next i
        
        strWkdays = Mid(strWkdays, 1, Len(strWkdays) - 1)
        getWeekDaysList = strWkdays
End Function

Private Sub Timer1_Timer()
 
    If IsScheduledNow = True Then
            Call RequestServer4BackUp
    End If
End Sub

Public Function GetSchedule() As String

Dim strArr() As String

On Error GoTo ErrHand
    If strFrequency <> "" Then
        strArr = Split(strFrequency, ",")
        strFrqType = strArr(0)
        strFrqDay = strArr(1)
        strFrqHH = strArr(2)
        strFrqMM = strArr(3)
    End If
    Exit Function
    
ErrHand:
    MsgBox Err.Description
End Function

Public Function IsScheduledNow() As Boolean
On Error GoTo ErrHand
    Dim strTemp As String
    
    If strFrqType = "D" Then
        If (strFrqHH & ":" & strFrqMM) = (Hour(Now()) & ":" & Minute(Now())) Then
            IsScheduledNow = True
        End If
        
    ElseIf strFrqType = "W" And _
        isOneOfTheScheduledWeekDay() And _
        (strFrqHH & ":" & strFrqMM) = (Hour(Now()) & ":" & Minute(Now())) Then
        
        If (strFrqHH & ":" & strFrqMM) = (Hour(Now()) & ":" & Minute(Now())) Then
            IsScheduledNow = True
        End If
        
    ElseIf strFrqType = "M" And _
        strFrqDay = Day(Now) And _
        (strFrqHH & ":" & strFrqMM) = (Hour(Now()) & ":" & Minute(Now())) Then
    
        If (strFrqHH & ":" & strFrqMM) = (Hour(Now()) & ":" & Minute(Now())) Then
            IsScheduledNow = True
        End If
    
    ElseIf strFrqType = "H" Then
        strTemp = DateAdd("h", Val(strFrqDay), CDate(strHrTime))
        If (Hour(strTemp) & ":" & Minute(strTemp)) = (Hour(Now()) & ":" & Minute(Now())) Then
            strHrTime = DateAdd("n", Val(strFrqDay), CDate(strHrTime))
            IsScheduledNow = True
        End If
    End If
    
    Exit Function

ErrHand:
    MsgBox Err.Description
End Function

Public Sub RequestServer4BackUp()
On Error GoTo ErrHand
Dim varFreeSpaceSz As Variant
    
    
        If Trim(g_strBackupPath) <> "" Then
            varFreeSpaceSz = CDec(getFreeSpaceOnDriveInMB(Mid(g_strBackupPath, 1, 1)))
        Else
            varFreeSpaceSz = CDec(getFreeSpaceOnDriveInMB())
        End If
        
        If CDec(varFreeSpaceSz) < CDec(1) Then
            FX_MSchart.LblDiskSpace.Visible = True
            FX_MSchart.Show
            Exit Sub
        End If
        
    
    
    Timer1.Enabled = False
    bTakeBackup = True
    
    Call GetServerIP
    FrmTransfer1.Show vbModal
    CloseConnection
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Public Sub GetServerWindowsName()
On Error GoTo ErrHand
 
Dim objBk As New clsBackUp
    CodeModule.strServerWindowsName = objBk.GetServerWindowsName(g_sPassword)
    CloseConnection
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Public Sub GetServerIP()
'On Error GoTo ErrHand
    Dim rstServerName As New ADODB.Recordset
    
    AdoConnection
    rstServerName.Open "SELECT ISNULL(AppValue, '') AS ServerIP  FROM  AppParam Where AppKey = 'ServerIP'", gdbConnection, adOpenStatic, adLockReadOnly
    strServerIP = rstServerName.Fields(0)
    FrmTransfer1.txtServerIP = strServerIP
    Set rstServerName = Nothing
    CloseConnection
End Sub


Public Sub CloseConnection()
On Error Resume Next
 If gdbConnection.State Then
        gdbConnection.Close
 End If
End Sub

Public Function isOneOfTheScheduledWeekDay() As Boolean
On Error Resume Next
Dim strArrTemp() As String
Dim i As Integer

    strArrTemp = Split(strFrqWeekDays, ",")
     
    For i = 0 To UBound(strArrTemp)
        'MsgBox Getweekday(strArrTemp(i)) & "  " & Weekday(date)
        If Getweekday(strArrTemp(i)) = Weekday(date) Then
            
            isOneOfTheScheduledWeekDay = True
            Exit Function
        Else
            isOneOfTheScheduledWeekDay = False
        End If
    Next
     
End Function

Private Function CMilliseconds(ByVal iHrs As Integer) As Double
    CMilliseconds = iHrs * 60 * 60 * 10000
End Function
