Attribute VB_Name = "mZip"
Option Explicit
      
'Dim clsSapObject As New clsSAPinterface

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' argv
Private Type ZIPnames
    s(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 4096) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

' Store the callback functions
Private Type ZIPUSERFUNCTIONS
    lPtrPrint As Long          ' Pointer to application's print routine
    lptrComment As Long        ' Pointer to application's comment routine
    lptrPassword As Long       ' Pointer to application's password routine.
    lptrService As Long        ' callback function designed to be used for allowing the
                               ' app to process Windows messages, or cancelling the operation
                               ' as well as giving option of progress.  If this function returns
                               ' non-zero, it will terminate what it is doing.  It provides the app
                               ' with the name of the archive member it has just processed, as well
                               ' as the original size.
End Type

Public Type ZPOPT
  date           As String ' US Date (8 Bytes Long) "12/31/98"?
  szRootDir      As String ' Root Directory Pathname (Up To 256 Bytes Long)
  szTempDir      As String ' Temp Directory Pathname (Up To 256 Bytes Long)
  fTemp          As Long   ' 1 If Temp dir Wanted, Else 0
  fSuffix        As Long   ' Include Suffixes (Not Yet Implemented!)
  fEncrypt       As Long   ' 1 If Encryption Wanted, Else 0
  fSystem        As Long   ' 1 To Include System/Hidden Files, Else 0
  fVolume        As Long   ' 1 If Storing Volume Label, Else 0
  fExtra         As Long   ' 1 If Excluding Extra Attributes, Else 0
  fNoDirEntries  As Long   ' 1 If Ignoring Directory Entries, Else 0
  fExcludeDate   As Long   ' 1 If Excluding Files Earlier Than Specified Date, Else 0
  fIncludeDate   As Long   ' 1 If Including Files Earlier Than Specified Date, Else 0
  fVerbose       As Long   ' 1 If Full Messages Wanted, Else 0
  fQuiet         As Long   ' 1 If Minimum Messages Wanted, Else 0
  fCRLF_LF       As Long   ' 1 If Translate CR/LF To LF, Else 0
  fLF_CRLF       As Long   ' 1 If Translate LF To CR/LF, Else 0
  fJunkDir       As Long   ' 1 If Junking Directory Names, Else 0
  fGrow          As Long   ' 1 If Allow Appending To Zip File, Else 0
  fForce         As Long   ' 1 If Making Entries Using DOS File Names, Else 0
  fMove          As Long   ' 1 If Deleting Files Added Or Updated, Else 0
  fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
  fUpdate        As Long   ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
  fFreshen       As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
  fJunkSFX       As Long   ' 1 If Junking SFX Prefix, Else 0
  fLatestTime    As Long   ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
  fComment       As Long   ' 1 If Putting Comment In Zip File, Else 0
  fOffsets       As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
  fPrivilege     As Long   ' 1 If Not Saving Privileges, Else 0
  fEncryption    As Long   ' Read Only Property!!!
  fRecurse       As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
  fRepair        As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
  flevel         As Byte   ' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type

'This assumes zip32.dll is in your \windows\system directory!
Private Declare Function ZpInit Lib "vbzip11.dll" (ByRef tUserFn As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "vbzip11.dll" (ByRef tOpts As ZPOPT) As Long ' Set Zip options
Private Declare Function ZpGetOptions Lib "vbzip11.dll" () As ZPOPT ' used to check encryption flag only
Private Declare Function ZpArchive Lib "vbzip11.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action

' Object for callbacks:
Private m_cZip As clsZip
Private m_bCancel As Boolean
Private m_bPasswordConfirm As Boolean

Private Function plAddressOf(ByVal lPtr As Long) As Long
   ' VB Bug workaround fn
   plAddressOf = lPtr
End Function

Public Function VBZip( _
      cZipObject As clsZip, _
      tZPOPT As ZPOPT, _
      sFileSpecs() As String, _
      iFileCount As Long _
   ) As Long
Dim tUser As ZIPUSERFUNCTIONS
Dim lR As Long
Dim i As Long
Dim sZipFile As String
Dim tZipName As ZIPnames

   m_bCancel = False
   m_bPasswordConfirm = False
   Set m_cZip = cZipObject

   If Not Len(Trim$(m_cZip.BasePath)) = 0 Then
      ChDir m_cZip.BasePath
   End If

   ' Set address of callback functions
   tUser.lPtrPrint = plAddressOf(AddressOf ZipPrintCallback)
   tUser.lptrPassword = plAddressOf(AddressOf ZipPasswordCallback)
   'tUser.lptrService = plAddressOf(AddressOf ZipServiceCallback)
   lR = ZpInit(tUser)

   ' Set options
   lR = ZpSetOptions(tZPOPT)
   
   ' Go for it!
   For i = 1 To iFileCount
      tZipName.s(i - 1) = sFileSpecs(i)
   Next i
   tZipName.s(i) = vbNullChar
   
   sZipFile = cZipObject.ZipFile
   lR = ZpArchive(iFileCount, sZipFile, tZipName)
   
   Debug.Print lR
   VBZip = lR

End Function

Private Function ZipServiceCallback( _
      ByRef mname As CBChar, _
      ByVal x As Long _
   ) As Long
Dim iPos As Long
Dim sInfo As String
Dim bCancel As Boolean
    
'-- Always Put This In Callback Routines!
On Error Resume Next
    
   ' Check we've got a message:
   If x > 1 And x < 1024 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), mname, x
      ' Convert to VB string:
      sInfo = StrConv(b, vbUnicode)
      iPos = InStr(sInfo, vbNullChar)
      If iPos > 0 Then
         sInfo = Left$(sInfo, iPos - 1)
      End If
      m_cZip.Service sInfo, bCancel
      If bCancel Then
         m_bCancel = True
         ZipServiceCallback = 1
      Else
         ZipServiceCallback = 0
      End If
   End If
   
End Function

Private Function ZipPrintCallback( _
      ByRef fname As CBChar, _
      ByVal x As Long _
   ) As Long
Dim iPos As Long
Dim sFIle As String
   On Error Resume Next
   
   ' Check we've got a message:
   If x > 1 And x < 1024 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), fname, x
      ' Convert to VB string:
      sFIle = StrConv(b, vbUnicode)
      If iPos > 0 Then
         sFIle = Left$(sFIle, iPos - 1)
      End If
      ' Tell the caller about it
      m_cZip.ProgressReport sFIle
   End If
   ZipPrintCallback = 0
   
End Function

Private Function ZipCommentCallback( _
      ByRef comm As CBChar _
   ) As Integer
Dim bCancel As Boolean
Dim sComment As String
Dim b() As Byte
Dim lSize As Long
Dim i As Long

On Error Resume Next
   ' not supported always return \0
   comm.ch(0) = vbNullString
   ZipCommentCallback = 1
   
   If m_bCancel Then
      Exit Function
   End If
     
   ' Cancel out if no useful comment:
   If bCancel Or Len(sComment) = 0 Then
      m_bCancel = True
      Exit Function
   End If
   
   ' Put password into return parameter:
   lSize = Len(sComment)
   b = StrConv(sComment, vbFromUnicode)
   i = 0
   lSize = UBound(b)
   If (lSize > 254) Then lSize = 254
   For i = 0 To lSize
      comm.ch(i) = b(i)
   Next i
   comm.ch(lSize + 1) = 0
   ZipCommentCallback = 0
   
End Function

Private Function ZipPasswordCallback( _
      ByRef pwd As CBCh, _
      ByVal maxPasswordLength As Long, _
      ByRef s2 As CBCh, _
      ByRef Name As CBCh _
   ) As Integer

Dim bCancel As Boolean
Dim sPassword As String
Dim b() As Byte
Dim lSize As Long
Dim i As Long

On Error Resume Next

   ' The default:
   ZipPasswordCallback = 1

   If m_bCancel Then
      Exit Function
   End If
   ''This is added by shahid on 12 apr 2005 for security password
   If gsBackUpVariable = 2 Or gsBackUpVariable = 2 Then
        sPassword = "Critical%data@very*userfull~05"
   ElseIf gsBackUpVariable = 1 Then
        sPassword = "This#is(for)system^admin~05"
   ElseIf gsBackUpVariable = 3 Or gsBackUpVariable = 3 Then
        sPassword = "This*will!save&my@work^05"
   End If
   ' Ask for password:
   m_cZip.PasswordRequest sPassword, maxPasswordLength, m_bPasswordConfirm, bCancel


   ' Put password into return parameter:
   lSize = Len(sPassword)
   b = StrConv(sPassword, vbFromUnicode)
   For i = 0 To maxPasswordLength - 1
      If (i <= UBound(b)) Then
         pwd.ch(i) = b(i)
      Else
         pwd.ch(i) = 0
      End If
   Next i

   ' Ask UnZip to process it:
   m_bPasswordConfirm = False
   ZipPasswordCallback = 0

End Function

