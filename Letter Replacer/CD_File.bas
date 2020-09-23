Attribute VB_Name = "CD_File"
'Replaces the CommonDialog control.
Option Explicit

Dim iAction As Integer
Dim lAPIReturn As Long
Dim bCancelError As Boolean
Dim sDefaultExt As String
Dim sDialogTitle As String
Dim lExtendedError As Long
Dim sFileName As String
Dim sFileTitle As String
Dim sFilter As String
Dim iFilterIndex As Integer
Dim lFlags As Long
Dim lHelpCommand As Long
Dim sHelpContext As String
Dim sHelpFile As String
Dim sHelpKey As String
Dim sInitDir As String
Dim lMax As Long
Dim lMaxFileSize As Long
Dim lMin As Long
Dim objObject As Object

Dim lhWndOwner As Long

Public Enum DlgFileFlags
   OFN_ALLOWMULTISELECT = &H200
   OFN_CREATEPROMPT = &H2000 = &H80
   OFN_EXPLORER = &H80000
   OFN_EXTENSIONDIFFERENT = &H400
   OFN_FILEMUSTEXIST = &H1000
   OFN_HIDEREADONLY = &H4
   OFN_NameS = &H200000
   OFN_NOCHANGEDIR = &H8
   OFN_NODEREFERENCELINKS = &H100000
   OFN_NONameS = &H40000
   OFN_NONETWORKBUTTON = &H20000
   OFN_NOREADONLYRETURN = &H8000
   OFN_NOTESTFILECREATE = &H10000
   OFN_NOVALIDATE = &H100
   OFN_OVERWRITEPROMPT = &H2
   OFN_PATHMUSTEXIST = &H800
   OFN_READONLY = &H1
   OFN_SHOWHELP = &H10
End Enum

Private Const CLSCD_NOACTION = 0
Private Const CLSCD_SHOWOPEN = 1
Private Const CLSCD_SHOWSAVE = 2
Private Const CLSCD_USERCANCELED = 0
Private Const CLSCD_USERSELECTED = 1

Private Const CLSCD_MAXFILESIZE = 128
Private Const CLSCD_ERRNUMUSRCANCEL = 32755
Private Const CLSCD_ERRDESUSRCANCEL = "Cancel was selected."
Private Const CLSCD_ERRNUMUSRBUFFER = 32756
Private Const CLSCD_ERRDESUSRBUFFER = "Buffer to small"

Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_FILENAMECODES = &H3000
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_SUBCLASSFAILURE = &H3001

Private Type tOPENFILENAME
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As DlgFileFlags
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" (pOpenfilename As tOPENFILENAME) As Long
Private Declare Function GetSaveFileNameA Lib "comdlg32.dll" (pOpenfilename As tOPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

' Read Only
Public Property Get Action() As Integer
   Action = iAction
End Property

' Read Only
Public Property Get APIReturn() As Long
   APIReturn = lAPIReturn
End Property

' Read/Write
Public Property Get CancelError() As Boolean
   CancelError = bCancelError
End Property
Public Property Let CancelError(vNewValue As Boolean)
   bCancelError = vNewValue
End Property


' Read/Write
Public Property Get DefaultExt() As String
   DefaultExt = sDefaultExt
End Property
Public Property Let DefaultExt(vNewValue As String)
   sDefaultExt = vNewValue
End Property

' Read/Write
Public Property Get DialogTitle() As String
   DialogTitle = sDialogTitle
End Property
Public Property Let DialogTitle(vNewValue As String)
   sDialogTitle = vNewValue
End Property

' Read Only
Public Property Get ExtendedError() As Long
   ExtendedError = lExtendedError
End Property

' Read/Write
Public Property Get FileName() As String
   FileName = sFileName
End Property
Public Property Let FileName(vNewValue As String)
   sFileName = vNewValue
End Property

' Read/Write
Public Property Get FileTitle() As String
   FileTitle = sFileTitle
End Property
Public Property Let FileTitle(vNewValue As String)
   sFileTitle = vNewValue
End Property

' Read/Write
Public Property Get filter() As String
   filter = sFilter
End Property
Public Property Let filter(vNewValue As String)
   sFilter = vNewValue
End Property

' Read/Write
Public Property Get FilterIndex() As Integer
   FilterIndex = iFilterIndex
End Property
Public Property Let FilterIndex(vNewValue As Integer)
   iFilterIndex = vNewValue
End Property

' Read/Write
Public Property Get Flags() As Long
   Flags = lFlags
End Property
Public Property Let Flags(vNewValue As Long)
   lFlags = vNewValue
End Property

' Read/Write
Public Property Get hWndOwner() As Long
   hWndOwner = lhWndOwner
End Property
Public Property Let hWndOwner(vNewValue As Long)
   lhWndOwner = vNewValue
End Property

' Read/Write
Public Property Get HelpCommand() As Long
   HelpCommand = lHelpCommand
End Property
Public Property Let HelpCommand(vNewValue As Long)
   lHelpCommand = vNewValue
End Property

' Read/Write
Public Property Get HelpContext() As String
   HelpContext = sHelpContext
End Property
Public Property Let HelpContext(vNewValue As String)
   sHelpContext = vNewValue
End Property

' Read/Write
Public Property Get HelpFile() As String
   HelpFile = sHelpFile
End Property
Public Property Let HelpFile(vNewValue As String)
   sHelpFile = vNewValue
End Property

' Read/Write
Public Property Get HelpKey() As String
   HelpKey = sHelpKey
End Property
Public Property Let HelpKey(vNewValue As String)
   sHelpKey = vNewValue
End Property

' Read/Write
Public Property Get InitDir() As String
   InitDir = sInitDir
End Property
Public Property Let InitDir(vNewValue As String)
   sInitDir = vNewValue
End Property

' Read/Write
Public Property Get MaxFileSize() As Long
   MaxFileSize = lMaxFileSize
End Property
Public Property Let MaxFileSize(vNewValue As Long)
   lMaxFileSize = vNewValue
End Property

'  Read Only
Public Property Get Object() As Object
   Object = objObject
End Property
'Provide the ShowOpen method.
Public Sub ShowOpen()
   ShowFileDialog (CLSCD_SHOWOPEN)
End Sub

'Provide the ShowSave method.
Public Sub ShowSave()
   ShowFileDialog (CLSCD_SHOWSAVE)
End Sub

Private Sub ShowFileDialog(ByVal iAction As Integer)
   Dim vOpenFile As tOPENFILENAME
   Dim lMaxSize As Long
   Dim sFileNameBuff As String
   Dim sFileTitleBuff As String
   
   On Error GoTo ShowFileDialogError
   iAction = iAction  'Action property
   lAPIReturn = 0  'APIReturn property
   lExtendedError = 0  'ExtendedError property
   If lMaxFileSize > 0 Then
      lMaxSize = lMaxFileSize
      Else
      lMaxSize = CLSCD_MAXFILESIZE
      End If
   
   vOpenFile.hWndOwner = lhWndOwner
   vOpenFile.lpstrFile = sFileName & Space(lMaxSize - Len(sFileName) - 1) & vbNullChar
   vOpenFile.nMaxFile = lMaxSize
   vOpenFile.lpstrDefExt = sDefaultExt
   vOpenFile.lpstrFileTitle = Space(lMaxSize - 1) & vbNullChar
   vOpenFile.nMaxFileTitle = lMaxSize
   vOpenFile.lpstrFilter = sAPIFilter(sFilter)
   vOpenFile.nFilterIndex = iFilterIndex
   vOpenFile.Flags = lFlags 'And Not (OFN_ALLOWMULTISELECT)
   vOpenFile.lpstrInitialDir = sInitDir
   vOpenFile.lpstrTitle = sDialogTitle
   vOpenFile.lStructSize = Len(vOpenFile)
   
   Select Case iAction
      Case CLSCD_SHOWOPEN
         lAPIReturn = GetOpenFileNameA(vOpenFile)
      Case CLSCD_SHOWSAVE
         lAPIReturn = GetSaveFileNameA(vOpenFile)
      Case Else   'unknown action
         Exit Sub
   End Select
   
   If lAPIReturn = CLSCD_USERSELECTED Then
      sFileName = sLeftOfNull(vOpenFile.lpstrFile)
      sFileTitle = sLeftOfNull(vOpenFile.lpstrFileTitle)
      Else
      lExtendedError = CommDlgExtendedError
      If lExtendedError = FNERR_BUFFERTOOSMALL Then
         On Error GoTo 0
         Err.Raise Number:=CLSCD_ERRNUMUSRBUFFER, Description:=CLSCD_ERRDESUSRBUFFER
         Exit Sub
         Else
         If bCancelError = True Then
            On Error GoTo 0
            Err.Raise Number:=CLSCD_ERRNUMUSRCANCEL, Description:=CLSCD_ERRDESUSRCANCEL
            Exit Sub
            End If
         End If
      End If
   Exit Sub
   
ShowFileDialogError:
   Exit Sub
   
End Sub

Private Function sAPIFilter(ByVal filter As String) As String
   Dim I As Long
   Dim C As String * 1
   Dim NullFilter As String
   
   For I = 1 To Len(filter)
      C = Mid(filter, I, 1)
      If C = "|" Then
         NullFilter = NullFilter & Chr(0)
         Else
         NullFilter = NullFilter & C
         End If
   Next I
   While Right(NullFilter, 2) <> Chr(0) & Chr(0)
      NullFilter = NullFilter & Chr(0)
   Wend
   sAPIFilter = NullFilter
End Function

Private Function sLeftOfNull(ByVal txt As String)
   Dim I As Long, P As Long
   Dim ntxt As String, k As String * 1
      
   P = InStr(txt, Chr(0) & Chr(0))
   If P > 0 Then
      For I = 1 To P - 1
         k = Mid(txt, I, 1)
         If k = Chr(0) Then ntxt = ntxt & " " Else ntxt = ntxt & k
      Next I
      Else
      ntxt = Left(txt, InStr(txt, Chr(0)) - 1)
      End If
   sLeftOfNull = ntxt
End Function

