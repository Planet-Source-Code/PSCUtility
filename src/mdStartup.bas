Attribute VB_Name = "mdStartup"
Option Explicit
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

Private Const JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE As Long = &H2000

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION               As String = "1.0"
Private Const STR_CONNSTR               As String = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=MS Access Database;Initial Catalog=%1"

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
    
    Set m_oOpt = GetOpt(vArgs, "list:-list:l:mdb:-mdb:d:password:-password:uploads:-uploads:u:pictures:-pictures:p")
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo list:l mdb:d password uploads:u pictures:p help:h:?")
        vKey = Split(vKey, ":")
        For lIdx = 0 To UBound(vKey)
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, lIdx))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, lIdx))
            End If
        Next
    Next
    If Not m_oOpt.Item("--nologo") Then
        ConsoleError App.ProductName & " v" & STR_VERSION & vbCrLf & Replace(App.LegalCopyright, "©", "(c)") & vbCrLf & vbCrLf
    End If
    If m_oOpt.Item("--error") Then
        ConsoleError m_oOpt.Item("--error") & vbCrLf
    End If
    If m_oOpt.Item("--help") Then
        ConsolePrint "Usage: " & App.EXEName & ".exe [options...]" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -l, --list FOLDER     list all zip archives contents in FOLDER" & vbCrLf & _
                    "  -d, --mdb FILE        PscEnc.mdb database FILE" & vbCrLf & _
                    "  --password SECRET     optional password for PscEnc.mdb" & vbCrLf & _
                    "  -u, --uploads FOLDER  search submissions zip archives in FOLDER (defaults to Uploads)" & vbCrLf & _
                    "  -p, --pictures FOLDER search submissions pictures in FOLDER (defaults to Pictures)" & vbCrLf & _
                    "  -v                    verbose output" & vbCrLf
        GoTo QH
    End If
    If LenB(m_oOpt.Item("--list")) <> 0 Then
        pvListArchives m_oOpt.Item("--list")
        GoTo QH
    End If
    If LenB(m_oOpt.Item("--mdb")) <> 0 Then
        pvUpload m_oOpt.Item("--mdb"), m_oOpt.Item("--password"), m_oOpt.Item("--uploads"), m_oOpt.Item("--pictures")
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

Private Function pvUpload(sDbFile As String, sPassword As String, sUploadsFolder As String, sPictureFolder As String) As Boolean
    Const FUNC_NAME     As String = "pvUpload"
    Dim cn              As Connection
    Dim rs              As Recordset
    Dim sSQL            As String
    Dim sTemplateDir    As String
    Dim sReadmeTempl    As String
    Dim sTempDir        As String
    Dim sRepoName       As String
    Dim oCmdComplete    As ADODB.Command
    Dim vElem           As Variant
    Dim sReadmeText     As String
    Dim oArchive        As cZipArchive
    Dim sPictureFile    As String
    Dim sZipFile        As String
    Dim sText           As String
    
    On Error GoTo EH
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Replace(STR_CONNSTR, "%1", sDbFile), Password:=sPassword
    Set rs = cn.OpenSchema(adSchemaTables)
    rs.Find "TABLE_NAME='Complete'"
    If rs.EOF Then
        cn.Execute "CREATE TABLE Complete(ID INT, WorldID INT, RepoName LONGTEXT, PRIMARY KEY(ID, WorldID))"
    End If
    sSQL = "SELECT      s.ID" & vbCrLf & _
           "            , s.WorldID" & vbCrLf & _
           "            , s.AuthorName" & vbCrLf & _
           "            , s.Title" & vbCrLf & _
           "            , c.Line AS Code" & vbCrLf & _
           "            , s.Description" & vbCrLf & _
           "            , s.Inputs" & vbCrLf & _
           "            , s.Assumes" & vbCrLf & _
           "            , s.CodeReturns" & vbCrLf & _
           "            , s.SideEffects" & vbCrLf & _
           "            , s.ApiDeclarations" & vbCrLf & _
           "            , s.PicturePath" & vbCrLf & _
           "            , s.ZipFilePath" & vbCrLf & _
           "FROM        (Submission AS s" & vbCrLf & _
           "INNER JOIN  Code AS c" & vbCrLf & _
           "ON          s.ID = c.ID AND s.WorldID = c.WorldID)" & vbCrLf & _
           "LEFT JOIN   Complete AS d" & vbCrLf & _
           "ON          s.ID = d.ID AND s.WorldID = d.WorldID" & vbCrLf & _
           "WHERE       c.LineNumber = 1 AND d.ID IS NULL" & vbCrLf & _
           "ORDER BY    s.AuthorName, s.Title, s.ID, s.WorldID"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSQL, cn
    If rs.RecordCount = 0 Then
        GoTo QH
    End If
    DebugLog MODULE_NAME, FUNC_NAME, Replace("Found %1 submissions for upload", "%1", rs.RecordCount)
    sTemplateDir = Left$(sDbFile, InStrRev(sDbFile, "\"))
    sReadmeTempl = ReadTextFile(PathCombine(sTemplateDir, "README.md"))
    If LenB(sUploadsFolder) = 0 Then
        sUploadsFolder = PathCombine(sTemplateDir, "Uploads")
    End If
    If LenB(sPictureFolder) = 0 Then
        sPictureFolder = PathCombine(sTemplateDir, "Pictures")
    End If
    sTempDir = PathCombine(sTemplateDir, "Temp")
    MkPath sTempDir
    ChDir sTempDir
    For Each vElem In EnumFiles(sTempDir)
        pvExec "cmd", Replace("/c rd %1 /s /q", "%1", ArgvQuote(vElem))
    Next
    Set oCmdComplete = New ADODB.Command
    oCmdComplete.CommandText = "INSERT INTO Complete(ID, WorldID, RepoName) SELECT ?, ?, ?"
    oCmdComplete.Parameters.Append oCmdComplete.CreateParameter("ID", adInteger, adParamInput)
    oCmdComplete.Parameters.Append oCmdComplete.CreateParameter("WorldID", adInteger, adParamInput)
    oCmdComplete.Parameters.Append oCmdComplete.CreateParameter("RepoName", adLongVarChar, adParamInput, -1)
    Set oCmdComplete.ActiveConnection = cn
    Do While Not rs.EOF
        sRepoName = LCase$(pvAppend(pvCleanup(C_Str(rs!AuthorName.Value)), "-", pvCleanup(C_Str(rs!Title.Value)))) & "__" & rs!WorldID.Value & "-" & rs!ID.Value
        DebugLog MODULE_NAME, FUNC_NAME, Replace("Uploading %1", "%1", sRepoName)
        If LenB(C_Str(rs!ZipFilePath.Value)) <> 0 Then
            sZipFile = GetFileName(C_Str(rs!ZipFilePath.Value))
            If Not FileExists(PathCombine(sUploadsFolder, sZipFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, Replace("Upload %1 not found", "%1", sZipFile), vbLogEventTypeError
                GoTo Continue
            End If
            Set oArchive = New cZipArchive
            If Not oArchive.OpenArchive(PathCombine(sUploadsFolder, sZipFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, oArchive.LastError, vbLogEventTypeError
                GoTo Continue
            End If
        Else
            Set oArchive = Nothing
            GoTo Continue
        End If
        If LenB(C_Str(rs!PicturePath.Value)) <> 0 Then
            sPictureFile = GetFileName(C_Str(rs!PicturePath.Value))
            If Not FileExists(PathCombine(sPictureFolder, sPictureFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, Replace("Picture %1 not found", "%1", sPictureFile), vbLogEventTypeError
                GoTo Continue
            End If
        Else
            sPictureFile = vbNullString
        End If
        pvExec "gh", Replace("repo create Planet-Source-Code/%1 --public -y", "%1", sRepoName)
        MkDir PathCombine(sTempDir, sRepoName)
        For Each vElem In EnumFiles(sTempDir)
            If LenB(sPictureFile) <> 0 Then
                If FileExists(PathCombine(sPictureFolder, sPictureFile)) Then
                    FileCopy PathCombine(sPictureFolder, sPictureFile), PathCombine(vElem, sPictureFile)
                Else
                    sPictureFile = vbNullString
                End If
            End If
            If Not oArchive Is Nothing Then
                If FileExists(PathCombine(sTemplateDir, ".gitattributes")) Then
                    FileCopy PathCombine(sTemplateDir, ".gitattributes"), PathCombine(vElem, ".gitattributes")
                End If
                If FileExists(PathCombine(sTemplateDir, ".gitignore")) Then
                    FileCopy PathCombine(sTemplateDir, ".gitignore"), PathCombine(vElem, ".gitignore")
                End If
                oArchive.Extract vElem
            End If
            sReadmeText = sReadmeTempl
            sReadmeText = Replace(sReadmeText, "{Title}", rs!Title.Value)
            sReadmeText = Replace(sReadmeText, "{AuthorName}", Zn(C_Str(rs!AuthorName.Value), "Unknown Author"))
            sReadmeText = Replace(sReadmeText, "{PICTURE_IMAGE}", IIf(LenB(sPictureFile) <> 0, "<img src=""" & sPictureFile & """>", vbNullString))
            sReadmeText = Replace(sReadmeText, "{Description}", pvToMarkdown(C_Str(rs!Description.Value)))
            sReadmeText = preg_replace("\s+$", sReadmeText, vbNullString) & vbCrLf
            sText = pvToMarkdown(pvEmptyIf(rs!Inputs.Value, "None"))
            sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvEmptyIf(rs!Assumes.Value, "None")))
            sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvEmptyIf(rs!CodeReturns.Value, "None")))
            sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvEmptyIf(rs!SideEffects.Value, "None")))
            sReadmeText = Replace(sReadmeText, "{EXTRA_TITLE}", IIf(LenB(sText) <> 0, "### More Info", vbNullString))
            sReadmeText = Replace(sReadmeText, "{EXTRA_TEXT}", IIf(LenB(sText) <> 0, sText & vbCrLf, vbNullString))
            sReadmeText = preg_replace("\s+$", sReadmeText, vbNullString) & vbCrLf
            sText = pvEmptyIf(rs!ApiDeclarations.Value, "None")
            sReadmeText = Replace(sReadmeText, "{API_TITLE}", IIf(LenB(sText) <> 0, "### API Declarations", vbNullString))
            sReadmeText = Replace(sReadmeText, "{API_TEXT}", IIf(LenB(sText) <> 0, "```" & vbCrLf & sText & vbCrLf & "```" & vbCrLf, vbNullString))
            sReadmeText = preg_replace("\s+$", sReadmeText, vbNullString) & vbCrLf
            sText = pvEmptyIf(C_Str(rs!code.Value), "Upload")
            sReadmeText = Replace(sReadmeText, "{CODE_TITLE}", IIf(LenB(sText) <> 0, "### Source Code", vbNullString))
            sReadmeText = Replace(sReadmeText, "{CODE_TEXT}", IIf(LenB(sText) <> 0, "```" & vbCrLf & sText & vbCrLf & "```" & vbCrLf, vbNullString))
            sReadmeText = preg_replace("\s+$", sReadmeText, vbNullString) & vbCrLf
            WriteTextFile PathCombine(vElem, "README.md"), sReadmeText
            ChDir vElem
            pvExec "git", "init"
            pvExec "git", "config user.email pscbot@saas.bg"
            pvExec "git", "config user.name pscbot"
            pvExec "git", Replace("remote add origin git@github.com:Planet-Source-Code/%1.git", "%1", sRepoName)
            pvExec "git", "add ."
            pvExec "git", "commit -m ""Initial commit"""
            pvExec "git", "push origin master"
            ChDir sTempDir
            pvExec "cmd", Replace("/c rd %1 /s /q", "%1", ArgvQuote(vElem))
            Exit For
        Next
        '--- mark complete
        oCmdComplete.Parameters("ID").Value = rs!ID.Value
        oCmdComplete.Parameters("WorldID").Value = rs!WorldID.Value
        oCmdComplete.Parameters("RepoName").Value = sRepoName
        oCmdComplete.Execute
Continue:
        rs.MoveNext
    Loop
    '--- success
    pvUpload = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCleanup(ByVal sText As String) As String
    pvCleanup = Replace(Trim$(preg_replace("\s+", preg_replace("[^A-Za-z0-9]", sText, " "), " ")), " ", "-")
End Function

Private Function pvAppend(sText As String, sDelim As String, sAppend As String) As String
    pvAppend = sText & IIf(LenB(sText) <> 0 And LenB(sAppend) <> 0, sDelim, vbNullString) & sAppend
End Function

Private Function pvEmptyIf(Value As Variant, EmptyValue As Variant) As Variant
    If Value <> EmptyValue Then
        pvEmptyIf = Value
    End If
End Function

Private Function pvExec(sFile As String, sParams As String) As String
    With New cExec
        .Run sFile, sParams, StartHidden:=True, LimitFlags:=JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        pvExec = .ReadAllOutput & .ReadAllError
        If C_Bool(m_oOpt.Item("-v")) Then
            ConsolePrint pvExec
        End If
    End With
End Function

Private Function pvToMarkdown(sText As String) As String
    pvToMarkdown = preg_replace("\r?\n", sText, vbCrLf & vbCrLf)
End Function
