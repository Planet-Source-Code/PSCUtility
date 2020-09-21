Attribute VB_Name = "mdStartup"
Option Explicit
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

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
