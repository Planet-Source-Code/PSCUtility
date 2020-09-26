Attribute VB_Name = "mdGlobals"
Option Explicit

'=========================================================================
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL                         As Long = 2
'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

'--- formats
Public Const FORMAT_TIME_ONLY           As String = "hh:nn:ss"
Public Const FORMAT_DATETIME_LOG        As String = "yyyy.MM.dd hh:nn:ss"
Public Const FORMAT_DATETIME_ISO        As String = "yyyy\-mm\-dd hh:nn:ss"
Public Const FORMAT_BASE_2              As String = "0.00"
Public Const FORMAT_BASE_3              As String = "0.000"
Private Const STR_PREFIX_ERROR          As String = "[Error] "
Private Const STR_PREFIX_WARNING        As String = "[Warning] "

'=========================================================================
' Error handling
'=========================================================================

'Private Sub PrintError(sFunction As String)
'    #If USE_DEBUG_LOG <> 0 Then
'        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
'    #Else
'        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
'    #End If
'End Sub

'=========================================================================
' Functions
'=========================================================================

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

Public Function Nz(vValue As Variant, Optional IfValueIsNull As String = vbNullString) As String
    Nz = IIf(IsNull(vValue), IfValueIsNull, C_Str(vValue))
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function PathCombine(ByVal sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Function GetFileName(ByVal sPath As String) As String
    sPath = Mid$(sPath, InStrRev(sPath, "\") + 1)
    GetFileName = Mid$(sPath, InStrRev(sPath, "/") + 1)
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

Public Function MkPath(sPath As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(sPath)
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1)) Then
                Exit Function
            End If
        End If
        If CreateDirectory(StrPtr(sPath), 0) = 0 Then
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim sPrefix         As String
    
    If App.LogMode = vbLogToNT Then
        App.LogEvent sText, Clamp(eType, 0, vbLogEventTypeInformation)
    Else
        sPrefix = Format$(Now, FORMAT_TIME_ONLY) & Right$(Format$(TimerEx, FORMAT_BASE_3), 4) & ": "
        Select Case eType
        Case vbLogEventTypeError
            sPrefix = sPrefix & STR_PREFIX_ERROR
        Case vbLogEventTypeWarning
            sPrefix = sPrefix & STR_PREFIX_WARNING
        End Select
        sPrefix = sPrefix & IIf(Len(sText) > 200, Left$(sText, 200) & "...", sText) & vbCrLf
        If eType = vbLogEventTypeError Then
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sPrefix
        Else
            ConsolePrint sPrefix
        End If
    End If
End Sub

Public Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal lMin As Long = -2147483647, _
            Optional ByVal lMax As Long = 2147483647) As Long
    Select Case lValue
    Case lMin To lMax
        Clamp = lValue
    Case Is < lMin
        Clamp = lMin
    Case Is > lMax
        Clamp = lMax
    End Select
End Function

Public Function SearchCollection(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Boolean
    Dim vItem           As Variant
    
    If oCol Is Nothing Then
        GoTo QH
    ElseIf pvCallCollectionItem(oCol, Index, vItem) < 0 Then
        GoTo QH
    End If
    If IsObject(vItem) Then
        Set RetVal = vItem
    Else
        RetVal = vItem
    End If
    '--- success
    SearchCollection = True
QH:
End Function

Private Function pvCallCollectionItem(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM As Long = 7
    
    pvPatchMethodTrampoline AddressOf mdGlobals.pvCallCollectionItem, IDX_COLLECTION_ITEM
    pvCallCollectionItem = pvCallCollectionItem(oCol, Index, RetVal)
End Function

Private Function pvPatchMethodTrampoline(ByVal Pfn As Long, ByVal lMethodIdx As Long) As Boolean
    Dim bInIDE          As Boolean

    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(Pfn, ByVal Pfn + &H16, 4)
    Else
        Call VirtualProtect(Pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal Pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (Pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * 4, 4)
    '--- success
    pvPatchMethodTrampoline = True
End Function

Public Function preg_replace(find_re As String, sText As String, Optional sReplace As String) As String
    preg_replace = pvInitRegExp(find_re).Replace(sText, sReplace)
End Function

Public Function preg_match(find_re As String, sText As String, Optional Matches As Variant, Optional Indexes As Variant) As Long
    Dim lIdx            As Long
    
    With pvInitRegExp(find_re).Execute(sText)
        preg_match = .Count
        If Not IsMissing(Matches) Then
            If .Count = 0 Then
                Matches = Split(vbNullString)
            ElseIf .Count = 1 Then
                ReDim Matches(0 To 0) As String
                Matches(0) = .Item(0).Value
            Else
                ReDim Matches(0 To .Count - 1) As String
                For lIdx = 0 To .Count - 1
                    Matches(lIdx) = .Item(lIdx).Value
                Next
            End If
        End If
        If Not IsMissing(Indexes) Then
            If .Count = 0 Then
                Indexes = Array()
            ElseIf .Count = 1 Then
                Indexes = Array(.Item(0).FirstIndex + 1)
            Else
                ReDim Indexes(0 To .Count - 1) As Variant
                For lIdx = 0 To .Count - 1
                    Indexes(lIdx) = .Item(lIdx).FirstIndex + 1
                Next
            End If
        End If
    End With
End Function

Private Function pvInitRegExp(sPattern As String) As Object
    Dim lIdx            As Long

    Set pvInitRegExp = CreateObject("VBScript.RegExp")
    With pvInitRegExp
        .Global = True
        If Left$(sPattern, 1) = "/" Then
            lIdx = InStrRev(sPattern, "/")
            .Pattern = Mid$(sPattern, 2, lIdx - 2)
            .IgnoreCase = (InStr(lIdx, sPattern, "i") > 0)
            .MultiLine = (InStr(lIdx, sPattern, "m") > 0)
        Else
            .Pattern = sPattern
        End If
    End With
End Function

' based on https://blogs.msdn.microsoft.com/twistylittlepassagesallalike/2011/04/23/everyone-quotes-command-line-arguments-the-wrong-way
Public Function ArgvQuote(ByVal sArg As String, Optional ByVal Force As Boolean) As String
    Const WHITESPACE As String = "*[ " & vbTab & vbVerticalTab & vbCrLf & "]*"
    
    If Not Force And LenB(sArg) <> 0 And Not sArg Like WHITESPACE Then
        ArgvQuote = sArg
    Else
        With pvInitRegExp("(\\+)($|"")|(\\+)")
            ArgvQuote = """" & Replace(.Replace(sArg, "$1$1$2$3"), """", "\""") & """"
        End With
    End If
End Function

Public Function ReadTextFile(sFile As String) As String
    Dim sCharset            As String

    sCharset = "utf-8"
    With CreateObject("ADODB.Stream")
        .Open
        If LenB(sCharset) <> 0 Then
            .Charset = sCharset
        End If
        .LoadFromFile sFile
        ReadTextFile = .ReadText()
    End With
End Function

Public Sub WriteTextFile(sFile As String, sText As String, Optional sCharset As String)
    With CreateObject("ADODB.Stream")
        .Open
        If LenB(sCharset) <> 0 Then
            .Charset = sCharset
        End If
        .WriteText sText
        .SaveToFile sFile, adSaveCreateOverWrite
    End With
End Sub

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
        FileExists = (Err.LastDllError = 32) ' ERROR_SHARING_VIOLATION
    Else
        FileExists = True
    End If
End Function

Public Function ConcatCollection(oCol As Collection, Optional Separator As String) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function
