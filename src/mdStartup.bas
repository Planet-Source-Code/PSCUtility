Attribute VB_Name = "mdStartup"
Option Explicit
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION               As String = "1.0"

Private m_oOpt                      As Object

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim lExitCode       As Long
    
    On Error GoTo EH
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde And lExitCode <> -1 Then
        Call ExitProcess(lExitCode)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Function Process(vArgs As Variant) As Long
    Dim vKey            As Variant
    Dim lIdx            As Long
    
    Set m_oOpt = GetOpt(vArgs, "list:-list:l")
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo folder:f help:h:?")
        vKey = Split(vKey, ":")
        For lIdx = 0 To UBound(vKey)
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, lIdx))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, lIdx))
            End If
        Next
    Next
    If Not C_Bool(m_oOpt.Item("--nologo")) Then
        ConsoleError App.ProductName & " v" & STR_VERSION & vbCrLf & Replace(App.LegalCopyright, "©", "(c)") & vbCrLf & vbCrLf
    End If
    If C_Bool(m_oOpt.Item("--help")) Then
        ConsolePrint "Usage: " & App.EXEName & ".exe [options...]" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -l, --list FOLDER   list all zip archives contents in FOLDER" & vbCrLf
        GoTo QH
    End If
    If LenB(m_oOpt.Item("-l")) <> 0 Then
        pvListArchives m_oOpt.Item("-l")
        GoTo QH
    End If
QH:
End Function

Private Function pvListArchives(sPath As String) As Boolean
    Const FUNC_NAME     As String = "pvListArchives"
    Const STR_PSC_README As String = "@PSC_ReadMe_"
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim oArchive        As cZipArchive
    Dim sName           As String
    Dim sExt            As String
    Dim oInfo           As Object
    Dim rs              As Recordset
    
    On Error GoTo EH
    Set oInfo = Nothing
    For Each vElem In EnumFiles(sPath)
        Set oArchive = New cZipArchive
        If oArchive.OpenArchive(vElem) Then
            For lIdx = 0 To oArchive.FileCount - 1
                sName = oArchive.FileInfo(lIdx)(0)
                If Left$(sName, Len(STR_PSC_README)) <> STR_PSC_README And Right$(sName, 1) <> "\" Then
                    ConsolePrint GetFileName(vElem) & "#" & sName & vbCrLf
                    sExt = GetFileExt(sName)
                    If LenB(sExt) > 0 Then
                        JsonItem(oInfo, "." & sExt) = JsonItem(oInfo, "." & sExt) + 1
                        If sExt <> GetFileExt2(sName) Then
                            sExt = GetFileExt2(sName)
                            JsonItem(oInfo, "." & sExt) = JsonItem(oInfo, "." & sExt) + 1
                        End If
                    End If
                End If
            Next
        End If
    Next
    Set rs = New ADODB.Recordset
    rs.Fields.Append "Ext", adVarWChar, 1000
    rs.Fields.Append "Count", adInteger
    rs.Open
    For Each vElem In JsonKeys(oInfo)
        rs.AddNew Array(0, 1), Array(vElem, JsonItem(oInfo, vElem))
    Next
    rs.Sort = "Count DESC, Ext"
    Do While Not rs.EOF
        ConsolePrint rs!Count.Value & vbTab & LCase$(rs!Ext.Value) & vbCrLf
        rs.MoveNext
    Loop
    '--- success
    pvListArchives = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Object
    Dim oRetVal         As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            .Item("-" & vElem) = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        GoTo Continue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Continue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If lIdx < LBound(vData) Then
            '--- lIdx = -1 for last element
            lIdx = UBound(vData) + 1 + lIdx
        End If
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = C_Str(vData(lIdx))
        End If
    End If
QH:
End Property

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Public Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Public Function EnumFiles( _
            sPath As String, _
            Optional FileMask As String, _
            Optional RetVal As Collection) As Collection
    Dim sFile           As String
    
    If RetVal Is Nothing Then
        Set RetVal = New Collection
    End If
    sFile = Dir(PathCombine(sPath, Zn(FileMask, "*.*")), vbDirectory)
    Do While LenB(sFile) <> 0
        If sFile <> "." And sFile <> ".." And (sFile Like FileMask Or Zn(FileMask, "*.*") = "*.*") Then
            sFile = PathCombine(sPath, sFile)
'            If Not SearchCollection(RetVal, sFile) Then
                RetVal.Add sFile, sFile
'            End If
        End If
        sFile = vbNullString
        sFile = Dir
    Loop
    Set EnumFiles = RetVal
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Function GetFileName(ByVal sPath As String) As String
    GetFileName = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function

Public Function GetFileExt(ByVal sPath As String) As String
    If InStrRev(sPath, ".") > InStrRev(sPath, "\") Then
        GetFileExt = Mid$(sPath, InStrRev(sPath, ".") + 1)
    End If
End Function

Public Function GetFileExt2(ByVal sPath As String) As String
    If InStrRev(sPath, ".") > InStrRev(sPath, "\") Then
        GetFileExt2 = Mid$(sPath, InStr(InStrRev(sPath, "\") + 1, sPath, ".") + 1)
    End If
End Function

