VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FX_MSchart 
   BackColor       =   &H00800000&
   Caption         =   "Configure Backup Location"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Legend"
      Height          =   495
      Left            =   362
      TabIndex        =   18
      Top             =   4440
      Width           =   4300
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Free Space"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Used Space"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   150
         Width           =   1125
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2160
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   120
         Width           =   255
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3255
      Left            =   360
      OleObjectBlob   =   "FX_MSChart.frx":0000
      TabIndex        =   0
      Top             =   1200
      Width           =   4305
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Close"
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
      Left            =   7650
      TabIndex        =   16
      Top             =   5880
      Width           =   1050
   End
   Begin VB.CommandButton CmdOk 
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
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   5880
      Width           =   1140
   End
   Begin VB.TextBox txtPath 
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
      Height          =   315
      Left            =   315
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   525
      Width           =   8385
   End
   Begin VB.DirListBox dirList 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4800
      TabIndex        =   2
      Top             =   1920
      Width           =   3885
   End
   Begin VB.DriveListBox drvList 
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
      Left            =   4815
      TabIndex        =   1
      Top             =   1485
      Width           =   3885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   1725
      Left            =   855
      TabIndex        =   3
      Top             =   1440
      Width           =   2085
      Begin VB.CommandButton cmdAddData 
         Caption         =   "Add Demo Data"
         Height          =   855
         Left            =   0
         Picture         =   "FX_MSChart.frx":24B8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddTitle 
         Caption         =   "Add Title"
         Height          =   495
         Left            =   450
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddLegend 
         Caption         =   "Add Legend"
         Height          =   495
         Left            =   180
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.Frame chartTypeReq 
         Height          =   1455
         Left            =   225
         TabIndex        =   5
         Top             =   1260
         Width           =   2235
         Begin VB.OptionButton graphType1 
            Caption         =   "2D Pie"
            Height          =   1245
            Index           =   6
            Left            =   1125
            Picture         =   "FX_MSChart.frx":313A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   180
            Width           =   990
         End
         Begin VB.OptionButton graphType1 
            Caption         =   "3D Bar"
            Height          =   1245
            Index           =   5
            Left            =   135
            Picture         =   "FX_MSChart.frx":4EDC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdPrintGraph 
         Caption         =   "Print"
         Height          =   735
         Left            =   0
         Picture         =   "FX_MSChart.frx":717E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1125
         Width           =   1095
      End
   End
   Begin VB.Label LblDiskSpace 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Disk Space Low Can Not Perform Backup Operation!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   17
      Tag             =   "Disk Space Low Can Not Perform Backup Operation!!"
      Top             =   5400
      Visible         =   0   'False
      Width           =   7620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "[All Figures in the Graph are in GB]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   315
      TabIndex        =   15
      Top             =   975
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Configure BackUp Location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   315
      TabIndex        =   14
      Top             =   75
      Width           =   5100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Atleast 1 GB Free Space Should be Available On Backup Drive "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   315
      TabIndex        =   11
      Top             =   5145
      Width           =   6795
   End
End
Attribute VB_Name = "FX_MSchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private ObjCipher As New clsCipher
Dim tglLegend As Boolean, tglTitle As Boolean

Private Sub cmdAddData_Click()

On Error Resume Next

'  This subroutine shows you how to setup and pass data to the chart object
'  imax is the maximum of columns that you want to display as demo data
'  datascale will adjust the actual values to suit a different axis scale

   Dim i As Integer
   Dim fso   As New FileSystemObject
   Dim x() As Variant
   Dim iRow As Integer
   Dim imax As Integer
   Dim iCnt As Integer
   
   For i = 1 To fso.Drives.Count - 1
      ' MsgBox fso.Drives.Item(Chr(66 + i)).DriveLetter & "  = " & fso.Drives.Item(Chr(66 + i)).DriveType
      
       If fso.Drives.Item(Chr(66 + i)).DriveType = Fixed Then
            imax = imax + 1
       End If
   Next
    
  ReDim x(1 To imax + 1, 1 To 3)
      
  x(1, 1) = "Drives"
  x(1, 2) = "Used Space:"
  x(1, 3) = "Free Space:"
  
  dataScale = 1
  
   For i = 1 To fso.Drives.Count - 1
      ' MsgBox fso.Drives.Item(Chr(66 + i)).DriveLetter
       If fso.Drives.Item(Chr(66 + i)).DriveType = Fixed Then
            iCnt = iCnt + 1
            x(iCnt + 1, 1) = fso.Drives.Item(Chr(66 + i)).DriveLetter & " Drive"
            x(iCnt + 1, 2) = FX_MSChartBas.getUsedSpaceOnDriveInMB(fso.Drives.Item(Chr(66 + i)).DriveLetter)
            x(iCnt + 1, 3) = FX_MSChartBas.getFreeSpaceOnDriveInMB(fso.Drives.Item(Chr(66 + i)).DriveLetter)
       End If
   Next
  
 

addChartData:

' Reset the chart back to default to avoid any surprises
  MSChart1.ToDefaults
    
 
  Call addDataArray(MSChart1, x(), True)

  With MSChart1
       .chartType = VtChSeriesType3dBar
       .Plot.Projection = VtProjectionTypeOblique
       
       .Stacking = True
  End With
'  cmdAddLegend_Click
End Sub



Private Sub cmdAddLegend_Click()

On Error GoTo ErrorHandler

  tglLegend = Not tglLegend
  
  Call AddLegend(MSChart1, tglLegend)

Exit Sub
ErrorHandler:
    MsgBox Err.Description
    

End Sub

Private Sub cmdAddTitle_Click()
  tglTitle = Not tglTitle

  Call AddTitle(MSChart1, "Disk Space", tglTitle)
End Sub



Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()


Dim strEncryBackupPath As String
Dim varFreeSpaceinMB As Variant
On Error GoTo ErrorHandler


    If Trim(txtPath) = "" Then
        MsgBox "Please Select the Directory!"
        Exit Sub
    End If
    
    varFreeSpaceinMB = getFreeSpaceOnDriveInMB(Mid(drvList.Drive, 1, 1))
    
    If CDec(varFreeSpaceinMB) < 1024 Then
        MsgBox "Disk Space is Less than 1 GB." & vbCrLf _
            & "Atleast 1 GB Free Space Should be Available On Backup Drive!!", vbCritical
        Exit Sub
    End If

    strEncryBackupPath = Trim(UCase(txtPath))
    SaveSetting "ABSOLUTE NetWare", "NetSoft", "BKP", strEncryBackupPath  'backup path
    frmLogin.txtBackupPath = txtPath
    frmLogin.strFrqBackupPath = txtPath
    g_strBackupPath = txtPath
    Unload Me


Exit Sub
ErrorHandler:
    MsgBox Err.Description
    

End Sub

Private Sub cmdPrintGraph1_Click()
On Error GoTo ErrorHandler

' First turn of all the controls that do not have a Tag property of Print

  Call PrintVisibility(False)

' Now print the objects that are still visible

  PrintForm

exitForm:
' And then turn all the other objects on again
  Call PrintVisibility(True)

Exit Sub
ErrorHandler:

  MsgBox "The form can't be printed."
  GoTo exitForm
  
End Sub
Private Sub PrintVisibility(visState As Boolean)
'  Hide or show all objects with a tag of Print
 
 Dim i As Integer

'  First turn off the other controls

 For i = 0 To Me.Count - 1
   If Me(i).Tag <> "Print" Then
     Me(i).Visible = visState
   End If
 Next

End Sub

 

Public Function getFreeSpaceOnDriveInMB(Optional ByVal strDriveLetter As String = "C") As String
On Error GoTo ErrHand
    Dim varMB As Variant
    Dim varCapacity As Variant
    Dim varUsed As Variant
    Dim fso As FileSystemObject
    
    
    varMB = CDec(1024) * CDec(1024)
    Set fso = New FileSystemObject

        
    Set oDrive = fso.GetDrive(strDriveLetter)
    varCapacity = Format(CDec(oDrive.TotalSize) / (varMB), "#############.00")
    varCapacity = Format(CDec(oDrive.TotalSize) / (varMB), "#############.00")
    
    getFreeSpaceOnDriveInMB = Format(CDec(oDrive.FreeSpace) / (varMB), "#############.00")

Exit Function
ErrHand:
    getFreeSpaceOnDriveInMB = "<ERROR>"
End Function

Private Sub Form_Load()
On Error Resume Next
Dim varFreeSpaceSz As Variant

    cmdAddData_Click
    drvList.Drive = frmLogin.txtBackupPath
    dirList.Path = frmLogin.txtBackupPath
    If Trim(txtPath) = "" Then
    txtPath = dirList.Path
    End If
    
    
    
    If Trim(g_strBackupPath) <> "" Then
        varFreeSpaceSz = CDec(getFreeSpaceOnDriveInMB(Mid(g_strBackupPath, 1, 1))) / CDec(1024)
    Else
        varFreeSpaceSz = CDec(getFreeSpaceOnDriveInMB()) / CDec(1024)
    End If
    
    If CDec(varFreeSpaceSz) < CDec(1) Then
        LblDiskSpace.Visible = True
    End If
    
End Sub

Private Sub graphType1_Click(Index As Integer)
On Error Resume Next
  Dim graphInt As Integer, chartStr As String
       
  With MSChart1
  
  chartStr = graphType1(Index).Caption
  
  Select Case chartStr
 
      Case Is = "2D Area"
      
         .chartType = VtChChartType2dArea
         .Stacking = True
         
      Case Is = "2D Bar"
      
         .chartType = VtChChartType2dBar
         .Stacking = False
  
     Case Is = "3D Bar"
  
       .chartType = VtChSeriesType3dBar
       .Plot.Projection = VtProjectionTypeOblique
       .Stacking = True
    
     Case Is = "2D Stack"
  
         .chartType = VtChChartType2dBar
         .Stacking = True
      
     Case Is = "2D Line"
     
       .chartType = VtChChartType2dLine
       .Stacking = False
     
     Case Is = "2D Pie"
     
       .chartType = VtChChartType2dPie
       .Stacking = False
  
     Case Else
      
          MsgBox "Chart Type Not Supported"
          
  End Select
  End With
End Sub


Private Sub drvList_Change()
On Error GoTo errHandler
    
    dirList.Path = drvList.Drive
    Exit Sub

errHandler:
    MsgBox Err.Description
    drvList.Drive = dirList.Path
End Sub


Private Sub dirList_Change()
On Error GoTo errHandler
    txtPath.Text = dirList.Path
    
    Exit Sub
errHandler:
    MsgBox Err.Description

End Sub

