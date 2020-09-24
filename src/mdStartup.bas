Attribute VB_Name = "mdStartup"
' Make sure all variables are declared
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
Private Const DBL_EPSILON               As Double = 0.000001

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
                sName = At(oArchive.FileInfo(lIdx), 0)
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
    Const MAX_REPONAME  As Long = 100
    Const STR_EMPTYLIST As String = "none|none.|none,|non|non.|non,|no|no.|no,|nothing|nothing.|nothing,|n/a|n/a.|n/a,|na|na.|na,|nil|nil.|nil,"
    Dim cn              As Connection
    Dim rs              As Recordset
    Dim sSql            As String
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
    Dim sResult         As String
    Dim lIdx            As Long
    
    On Error GoTo EH
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open Replace(STR_CONNSTR, "%1", sDbFile), Password:=sPassword
    pvCreateExtraTables cn
    sSql = "SELECT      s.ID" & vbCrLf & _
           "            , s.WorldId" & vbCrLf & _
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
           "            , s.UserRatingTotal" & vbCrLf & _
           "            , s.NumOfUserRatings" & vbCrLf & _
           "            , sc1.CompatibilityName" & vbCrLf & _
           "            , cat.CategoryName" & vbCrLf & _
           "            , dif.Name AS CodeDifficultyName" & vbCrLf
    sSql = sSql & _
           "FROM        (((((Submission AS s" & vbCrLf & _
           "INNER JOIN  Code AS c" & vbCrLf & _
           "ON          s.ID = c.ID AND s.WorldId = c.WorldId)" & vbCrLf & _
           "LEFT JOIN   Complete AS d" & vbCrLf & _
           "ON          s.ID = d.ID AND s.WorldId = d.WorldId)" & vbCrLf & _
           "LEFT JOIN   SubmissionCompatibility sc1" & vbCrLf & _
           "ON          s.ID = sc1.ID AND s.WorldId = sc1.WorldId)" & vbCrLf & _
           "LEFT JOIN   Category cat" & vbCrLf & _
           "ON          s.CategoryId = cat.CategoryId AND s.WorldId = cat.WorldId)" & vbCrLf & _
           "LEFT JOIN   DifficultyType dif" & vbCrLf & _
           "ON          s.CodeDifficultyTypeId = dif.DifficultyTypeId)" & vbCrLf & _
           "WHERE       c.LineNumber = 1 AND d.ID IS NULL" & vbCrLf & _
           "ORDER BY    s.AuthorName, s.Title, s.ID, s.WorldID"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSql, cn
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
        sRepoName = LCase$(Left$(pvAppend(pvCleanup(C_Str(rs!AuthorName.Value)), "-", pvCleanup(C_Str(rs!Title.Value))), MAX_REPONAME - 10)) & "__" & rs!WorldID.Value & "-" & rs!ID.Value
        DebugLog MODULE_NAME, FUNC_NAME, Replace(Replace(Replace("Uploading %1 [%2/%3]", "%1", sRepoName), "%2", rs.AbsolutePosition), "%3", rs.RecordCount)
        If LenB(Trim$(C_Str(rs!ZipFilePath.Value))) <> 0 Then
            sZipFile = StrConv(StrConv(Trim$(GetFileName(C_Str(rs!ZipFilePath.Value))), vbFromUnicode), vbUnicode)
            If Not FileExists(PathCombine(sUploadsFolder, sZipFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, Replace("Submission %1 not found", "%1", sZipFile), vbLogEventTypeError
                GoTo Continue
            End If
            Set oArchive = New cZipArchive
            If Not oArchive.OpenArchive(PathCombine(sUploadsFolder, sZipFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, oArchive.LastError, vbLogEventTypeError
                GoTo Continue
            End If
        Else
            Set oArchive = Nothing
        End If
        If LenB(C_Str(rs!PicturePath.Value)) <> 0 Then
            sPictureFile = StrConv(StrConv(Trim$(GetFileName(C_Str(rs!PicturePath.Value))), vbFromUnicode), vbUnicode)
            If Not FileExists(PathCombine(sPictureFolder, sPictureFile)) Then
                DebugLog MODULE_NAME, FUNC_NAME, Replace("Picture %1 not found", "%1", sPictureFile), vbLogEventTypeError
                sPictureFile = vbNullString
            End If
        Else
            sPictureFile = vbNullString
        End If
        sResult = pvExec("gh", Replace("repo create Planet-Source-Code/%1 --public -y", "%1", sRepoName))
        MkDir PathCombine(sTempDir, sRepoName)
        For Each vElem In EnumFiles(sTempDir)
            ChDir vElem
            pvExec "git", "init"
            pvExec "git", "config user.email pscbot@saas.bg"
            pvExec "git", "config user.name pscbot"
            pvExec "git", Replace("remote add origin git@github.com:Planet-Source-Code/%1.git", "%1", sRepoName)
            If InStr(1, sResult, "name already exists", vbTextCompare) > 0 Then
                sResult = pvExec("git", "pull origin master")
            End If
            If LenB(sPictureFile) <> 0 Then
                If FileExists(PathCombine(sPictureFolder, sPictureFile)) Then
                    FileCopy PathCombine(sPictureFolder, sPictureFile), PathCombine(vElem, sPictureFile)
                Else
                    sPictureFile = vbNullString
                End If
            ElseIf Not oArchive Is Nothing Then
                For lIdx = 0 To oArchive.FileCount - 1
                    If preg_match("/^[^\\]+\.(gif|jpg|png)$/i", At(oArchive.FileInfo(lIdx), 0)) > 0 Then
                        sPictureFile = At(oArchive.FileInfo(lIdx), 0)
                        DebugLog MODULE_NAME, FUNC_NAME, Replace("Will use %1 picture instead", "%1", sPictureFile)
                        Exit For
                    End If
                Next
            End If
            If Not oArchive Is Nothing Then
                oArchive.Extract vElem
                If FileExists(PathCombine(sTemplateDir, ".gitattributes")) Then
                    FileCopy PathCombine(sTemplateDir, ".gitattributes"), PathCombine(vElem, ".gitattributes")
                End If
                If FileExists(PathCombine(sTemplateDir, ".gitignore")) Then
                    FileCopy PathCombine(sTemplateDir, ".gitignore"), PathCombine(vElem, ".gitignore")
                End If
            End If
            sReadmeText = sReadmeTempl
            sReadmeText = Replace(sReadmeText, "{Title}", rs!Title.Value)
            sReadmeText = Replace(sReadmeText, "{AuthorName}", Zn(C_Str(rs!AuthorName.Value), "Unknown Author"))
            sReadmeText = Replace(sReadmeText, "{PICTURE_IMAGE}", IIf(LenB(sPictureFile) <> 0, "<img src=""" & sPictureFile & """>", vbNullString))
            sReadmeText = Replace(sReadmeText, "{Description}", pvToMarkdown(pvRTrim(C_Str(rs!Description.Value))))
            sText = pvToMarkdown(pvRTrim(pvEmptyIf(C_Str(rs!Inputs.Value), STR_EMPTYLIST)))
            If C_Str(rs!Assumes.Value) <> C_Str(rs!Inputs.Value) Then
                sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvRTrim(pvEmptyIf(C_Str(rs!Assumes.Value), STR_EMPTYLIST))))
            End If
            If C_Str(rs!CodeReturns.Value) <> C_Str(rs!Inputs.Value) And C_Str(rs!CodeReturns.Value) <> C_Str(rs!Assumes.Value) Then
                sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvRTrim(pvEmptyIf(C_Str(rs!CodeReturns.Value), STR_EMPTYLIST))))
            End If
            If C_Str(rs!SideEffects.Value) <> C_Str(rs!Inputs.Value) And C_Str(rs!SideEffects.Value) <> C_Str(rs!Assumes.Value) And C_Str(rs!SideEffects.Value) <> C_Str(rs!CodeReturns.Value) Then
                sText = pvAppend(sText, vbCrLf & vbCrLf, pvToMarkdown(pvRTrim(pvEmptyIf(C_Str(rs!SideEffects.Value), STR_EMPTYLIST))))
            End If
            sReadmeText = Replace(sReadmeText, "{EXTRA_TITLE}", "### More Info")
            sReadmeText = Replace(sReadmeText, "{EXTRA_TEXT}", IIf(LenB(sText) <> 0, sText & vbCrLf, vbNullString))
            sReadmeText = Replace(sReadmeText, "{CategoryName}", C_Str(rs!CategoryName.Value))
            sReadmeText = Replace(sReadmeText, "{CodeDifficultyName}", C_Str(rs!CodeDifficultyName.Value))
            If Abs(rs!NumOfUserRatings.Value) > DBL_EPSILON Then
                sReadmeText = Replace(sReadmeText, "{USER_RATING}", Format$(rs!UserRatingTotal.Value / rs!NumOfUserRatings.Value, "0.0"))
            Else
                sReadmeText = Replace(sReadmeText, "{USER_RATING}", "N/A")
            End If
            sReadmeText = Replace(sReadmeText, "{UserRatingTotal}", C_Str(rs!UserRatingTotal.Value))
            sReadmeText = Replace(sReadmeText, "{NumOfUserRatings}", C_Str(rs!NumOfUserRatings.Value))
            sReadmeText = Replace(sReadmeText, "{CompatibilityName}", C_Str(rs!CompatibilityName.Value))
            sText = pvRTrim(pvEmptyIf(C_Str(rs!ApiDeclarations.Value), STR_EMPTYLIST))
            sReadmeText = Replace(sReadmeText, "{API_TITLE}", IIf(LenB(sText) <> 0, "### API Declarations", vbNullString))
            If (pvLineCount(sText) > 1 Or pvIsCode(sText)) And Not pvIsHtml(sText) Then
                sText = "```" & vbCrLf & sText & vbCrLf & "```"
            End If
            sReadmeText = Replace(sReadmeText, "{API_TEXT}", IIf(LenB(sText) <> 0, sText & vbCrLf, vbNullString))
            sText = pvRTrim(pvEmptyIf(C_Str(rs!Code.Value), "upload"))
            sReadmeText = Replace(sReadmeText, "{CODE_TITLE}", IIf(LenB(sText) <> 0, "### Source Code", vbNullString))
            If (pvLineCount(sText) > 1 Or pvIsCode(sText)) And Not pvIsHtml(sText) Then
                sText = "```" & vbCrLf & sText & vbCrLf & "```"
            End If
            sReadmeText = Replace(sReadmeText, "{CODE_TEXT}", IIf(LenB(sText) <> 0, sText & vbCrLf, vbNullString))
            WriteTextFile PathCombine(vElem, "README.md"), sReadmeText, "utf-8"
            pvExec "git", "add ."
            sResult = pvExec("git", "commit -m ""Initial commit"" --amend")
            If InStr(1, sResult, "fatal: ", vbTextCompare) > 0 Or InStr(1, sResult, "error: ", vbTextCompare) > 0 Then
                sResult = pvExec("git", "commit -m ""Initial commit""")
            End If
            sResult = pvExec("git", "push origin master --force")
            If InStr(1, sResult, "fatal: ", vbTextCompare) > 0 Or InStr(1, sResult, "error: ", vbTextCompare) > 0 Then
                sResult = pvExec("git", "push origin master --force")
            End If
            ChDir sTempDir
            pvExec "cmd", Replace("/c rd %1 /s /q", "%1", ArgvQuote(vElem))
            If InStr(1, sResult, "error: ", vbTextCompare) > 0 Then
                GoTo Continue
            End If
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
    pvCleanup = Replace(Trim$(preg_replace("/[ \t\r\n]+/m", preg_replace("[^A-Za-z0-9]", sText, " "), " ")), " ", "-")
End Function

Private Function pvAppend(sText As String, sDelim As String, sAppend As String) As String
    pvAppend = sText & IIf(LenB(sText) <> 0 And LenB(sAppend) <> 0, sDelim, vbNullString) & sAppend
End Function

Private Function pvEmptyIf(Value As String, EmptyValues As String) As String
    Dim vElem           As Variant
    
    For Each vElem In Split(EmptyValues, "|")
        If LCase$(Value) = LCase$(vElem) Then
            Exit Function
        End If
    Next
    pvEmptyIf = Value
End Function

Private Function pvExec(sFile As String, sParams As String) As String
    Const FUNC_NAME     As String = "pvExec"
    Dim vElem           As Variant
    
    With New cExec
        .Run sFile, sParams, StartHidden:=True, LimitFlags:=JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        pvExec = .ReadAllOutput & .ReadAllError
        If C_Bool(m_oOpt.Item("-v")) Then
            For Each vElem In Split(pvExec, vbLf)
                DebugLog MODULE_NAME, FUNC_NAME, Replace(Replace(vElem, vbCr, vbNullString), vbNullChar, vbNullString)
            Next
        End If
    End With
End Function

Private Function pvToMarkdown(sText As String) As String
    pvToMarkdown = preg_replace("/^[ \t]+/m", preg_replace("/\r?\n|____+ |----+|====+ /m", sText, vbCrLf & vbCrLf), vbNullString)
End Function

Private Function pvRTrim(sText As String) As String
    pvRTrim = preg_replace("/[ \t\r\n]+$/m", sText, vbNullString)
End Function

Private Function pvIsHtml(sText As String) As Boolean
    If preg_match("/<[%?!][^-]/m", sText) = 0 And preg_match("/<script[^<]*>/i", sText) = 0 Then
        If preg_match("<\w+[^<]*>", sText) > 2 And preg_match("</\w+[^<]*>", sText) > 2 Then
            pvIsHtml = True
        ElseIf preg_match("/<br\s*/?/i>", sText) > 2 Then
            pvIsHtml = True
        End If
    End If
End Function

Private Function pvIsCode(sText As String) As Boolean
    If preg_match("/\r?\n'/m", vbCrLf & sText) > 0 Or preg_match("/Declare \w+ \w+ Lib/i", sText) > 0 Or preg_match("\#include", sText) > 0 Then
        pvIsCode = True
    End If
End Function

Private Function pvLineCount(sText As String) As Long
    pvLineCount = preg_match("/\r?\n/m", sText)
End Function

Private Sub pvCreateExtraTables(cn As Connection)
    Dim rs              As Recordset
    Dim sSql            As String
        
    Set rs = cn.OpenSchema(adSchemaTables)
    rs.Find "TABLE_NAME='Complete'"
    If rs.EOF Then
        sSql = "CREATE TABLE Complete(" & vbCrLf & _
               "           ID              INT" & vbCrLf & _
               "           , WorldID       INT" & vbCrLf & _
               "           , RepoName      LONGTEXT" & vbCrLf & _
               "           , PRIMARY KEY(ID, WorldID)" & vbCrLf & _
               "           )"
        cn.Execute sSql
    End If
    Set rs = cn.OpenSchema(adSchemaTables)
    rs.Find "TABLE_NAME='SubmissionCompatibility'"
    If rs.EOF Then
        sSql = "SELECT      scm.*" & vbCrLf & _
               "            , '' AS CompatibilityName" & vbCrLf & _
               "            , 0 AS CompatibilityOrderIndex" & vbCrLf & _
               "            , 0 AS SeqNo" & vbCrLf & _
               "INTO        TempCompatibility" & vbCrLf & _
               "FROM        SubmissionCompatibilityMemb scm"
        cn.Execute sSql
        sSql = "UPDATE      TempCompatibility" & vbCrLf & _
               "INNER JOIN  Compatibility" & vbCrLf & _
               "ON          Compatibility.CompatbilityId = TempCompatibility.CompatbilityId" & vbCrLf & _
               "SET         TempCompatibility.CompatibilityOrderIndex = Compatibility.OrderIndex " & vbCrLf & _
               "            , TempCompatibility.CompatibilityName = Compatibility.Name"
        cn.Execute sSql
        sSql = "SELECT      s1.SubmissionCompatibilityMembId" & vbCrLf & _
               "            , COUNT(*) AS SeqNo" & vbCrLf & _
               "INTO        TempSeqNo" & vbCrLf & _
               "FROM        TempCompatibility s1" & vbCrLf & _
               "INNER JOIN  TempCompatibility s2" & vbCrLf & _
               "ON          s2.SubmissionId = s1.SubmissionId" & vbCrLf & _
               "            AND s2.WorldId = s1.WorldId" & vbCrLf & _
               "            AND s2.CompatibilityOrderIndex <= s1.CompatibilityOrderIndex" & vbCrLf & _
               "GROUP BY    s1.SubmissionCompatibilityMembId"
        cn.Execute sSql
        sSql = "UPDATE      TempCompatibility" & vbCrLf & _
               "INNER JOIN  TempSeqNo s" & vbCrLf & _
               "ON          TempCompatibility.SubmissionCompatibilityMembId = s.SubmissionCompatibilityMembId" & vbCrLf & _
               "SET         TempCompatibility.SeqNo = s.SeqNo"
        cn.Execute sSql
        sSql = "SELECT      s.ID, s.WorldId, sc1.CompatibilityName & IIf(IsNull(sc2.CompatibilityName), '', ', ' & sc2.CompatibilityName) & IIf(IsNull(sc3.CompatibilityName), '', ', ' & sc3.CompatibilityName) " & vbCrLf & _
               "                & IIf(IsNull(sc4.CompatibilityName), '', ', ' & sc4.CompatibilityName) & IIf(IsNull(sc5.CompatibilityName), '', ', ' & sc5.CompatibilityName)" & vbCrLf & _
               "                & IIf(IsNull(sc6.CompatibilityName), '', ', ' & sc6.CompatibilityName) & IIf(IsNull(sc7.CompatibilityName), '', ', ' & sc7.CompatibilityName)" & vbCrLf & _
               "                & IIf(IsNull(sc8.CompatibilityName), '', ', ' & sc8.CompatibilityName) & IIf(IsNull(sc9.CompatibilityName), '', ', ' & sc9.CompatibilityName) AS CompatibilityName" & vbCrLf & _
               "INTO        SubmissionCompatibility" & vbCrLf & _
               "FROM        (((((((((Submission s" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 1) sc1" & vbCrLf & _
               "ON          s.ID = sc1.SubmissionId AND s.WorldId = sc1.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 2) sc2" & vbCrLf & _
               "ON          s.ID = sc2.SubmissionId AND s.WorldId = sc2.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 3) sc3" & vbCrLf & _
               "ON          s.ID = sc3.SubmissionId AND s.WorldId = sc3.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 4) sc4" & vbCrLf & _
               "ON          s.ID = sc4.SubmissionId AND s.WorldId = sc4.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 5) sc5" & vbCrLf & _
               "ON          s.ID = sc5.SubmissionId AND s.WorldId = sc5.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 6) sc6" & vbCrLf & _
               "ON          s.ID = sc6.SubmissionId AND s.WorldId = sc6.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 7) sc7" & vbCrLf & _
               "ON          s.ID = sc7.SubmissionId AND s.WorldId = sc7.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 8) sc8" & vbCrLf & _
               "ON          s.ID = sc8.SubmissionId AND s.WorldId = sc8.WorldId)" & vbCrLf & _
               "LEFT JOIN   (SELECT * FROM TempCompatibility WHERE SeqNo = 9) sc9" & vbCrLf & _
               "ON          s.ID = sc9.SubmissionId AND s.WorldId = sc9.WorldId)"
        cn.Execute sSql
    End If
End Sub
