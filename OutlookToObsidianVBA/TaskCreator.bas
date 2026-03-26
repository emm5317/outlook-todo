Attribute VB_Name = "TaskCreator"
Option Explicit

' Core markdown generation and vault I/O for OutlookToObsidian.
' Ported from OutlookToObsidian/TaskCreator.cs

Public Type TaskOptions
    Subject As String
    DueDate As String           ' "yyyy-mm-dd" or "" for none
    Priority As String          ' emoji string or "" for normal
    Tags As String              ' space-separated #tags
    Notes As String             ' optional user note
    AttachmentCount As Long     ' from MailItem.Attachments.Count
    HasOptions As Boolean       ' True when populated (VBA has no Nothing for Types)
End Type

' ---------------------------------------------------------------------------
' CreateTask — builds Obsidian Tasks-formatted markdown from a MailItem
' ---------------------------------------------------------------------------
Public Function CreateTask(mail As Outlook.MailItem, Optional opts As TaskOptions) As String
    Dim subj As String
    Dim senderName As String
    Dim received As String
    Dim entryId As String
    Dim today As String
    Dim Priority As String
    Dim Tags As String
    Dim dueDate As String
    Dim attachments As Long
    Dim bodyPreview As String
    Dim taskLine As String
    Dim senderLine As String
    Dim idHash As String
    Dim result As String

    If opts.HasOptions And opts.Subject <> "" Then
        subj = opts.Subject
    Else
        subj = SanitizeForMarkdown(SafeString(mail.Subject, "(no subject)"))
    End If

    senderName = SafeString(mail.SenderName, "Unknown")
    received = Format$(mail.ReceivedTime, "yyyy-mm-dd hh:nn")
    entryId = mail.entryId
    today = Format$(Now, "yyyy-mm-dd")

    ' Priority: use options override or auto-detect from Outlook importance
    If opts.HasOptions And opts.Priority <> "" Then
        Priority = opts.Priority
    Else
        Priority = MapPriority(mail.Importance)
    End If

    ' Tags: use options override or auto-detect from categories + default
    If opts.HasOptions Then
        Tags = Trim$(opts.Tags) & " "
    Else
        Tags = "#follow-up " & MapCategories(mail.Categories)
    End If

    ' Due date: from options or empty
    If opts.HasOptions And opts.DueDate <> "" Then
        dueDate = ChrW$(&HD83D) & ChrW$(&HDCC5) & " " & opts.DueDate & " "  ' calendar emoji
    Else
        dueDate = ""
    End If

    ' Attachment count
    If opts.HasOptions Then
        attachments = opts.AttachmentCount
    Else
        attachments = GetAttachmentCount(mail)
    End If

    bodyPreview = GetBodyPreview(mail.body, 140)

    ' Task line: - [ ] subject priority tags due-date created-date
    taskLine = "- [ ] " & subj & " " & Priority & Tags & dueDate & ChrW$(&H2795) & " " & today
    taskLine = RTrimWhitespace(taskLine)
    result = taskLine & vbLf

    ' Sender line: bold name, date, optional attachment count, dedup hash
    idHash = GetShortHash(entryId)
    senderLine = "  > **" & senderName & "** | " & received
    If attachments > 0 Then
        senderLine = senderLine & " | " & ChrW$(&HD83D) & ChrW$(&HDCCE) & " " & attachments  ' paperclip emoji
    End If
    senderLine = senderLine & " | ^" & idHash
    result = result & senderLine & vbLf

    ' Body preview
    If bodyPreview <> "" Then
        result = result & "  > " & bodyPreview & vbLf
    End If

    ' User notes (from detailed dialog)
    If opts.HasOptions And opts.Notes <> "" Then
        result = result & "  > **Note:** " & SanitizeForMarkdown(opts.Notes) & vbLf
    End If

    result = result & vbLf
    CreateTask = result
End Function

' ---------------------------------------------------------------------------
' BuildDefaultOptions — pre-fill TaskOptions from a MailItem
' ---------------------------------------------------------------------------
Public Function BuildDefaultOptions(mail As Outlook.MailItem) As TaskOptions
    Dim opts As TaskOptions
    opts.HasOptions = True
    opts.Subject = SanitizeForMarkdown(SafeString(mail.Subject, "(no subject)"))
    opts.DueDate = ""
    opts.Priority = MapPriority(mail.Importance)
    opts.Tags = "#follow-up " & RTrimWhitespace(MapCategories(mail.Categories))
    opts.Notes = ""
    opts.AttachmentCount = GetAttachmentCount(mail)
    BuildDefaultOptions = opts
End Function

' ---------------------------------------------------------------------------
' AppendToVault — appends markdown to the configured vault file
' Returns the resolved file name on success
' ---------------------------------------------------------------------------
Public Function AppendToVault(markdown As String) As String
    Dim vaultPath As String
    vaultPath = Settings.GetVaultPath()

    If vaultPath = "" Or Dir(vaultPath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1, "TaskCreator", _
            "Obsidian vault path is not configured or does not exist. Please restart Outlook to set it up."
    End If

    Dim fileName As String
    fileName = GetTargetFileName()
    Dim targetPath As String
    targetPath = vaultPath & "\" & fileName

    ' Create file with header if it doesn't exist
    If Dir(targetPath) = "" Then
        Dim header As String
        If Settings.GetUseDailyNotes() Then
            header = "# " & Format$(Now, "yyyy-mm-dd") & vbLf & vbLf
        Else
            header = "# " & RemoveExtension(fileName) & vbLf & vbLf
        End If
        WriteTextFile targetPath, header
    End If

    ' Append markdown
    AppendTextFile targetPath, markdown
    AppendToVault = fileName
End Function

' ---------------------------------------------------------------------------
' IsDuplicate — checks whether a task with this EntryID already exists
' ---------------------------------------------------------------------------
Public Function IsDuplicate(entryId As String) As Boolean
    Dim vaultPath As String
    vaultPath = Settings.GetVaultPath()
    If vaultPath = "" Then
        IsDuplicate = False
        Exit Function
    End If

    Dim fileName As String
    fileName = GetTargetFileName()
    Dim targetPath As String
    targetPath = vaultPath & "\" & fileName

    If Dir(targetPath) = "" Then
        IsDuplicate = False
        Exit Function
    End If

    Dim content As String
    content = ReadTextFile(targetPath)

    Dim hash As String
    hash = GetShortHash(entryId)

    ' Check current format (^hash), HTML comment format, and old Dataview format
    IsDuplicate = (InStr(1, content, "^" & hash) > 0) Or _
                  (InStr(1, content, "<!-- entry-id: " & entryId & " -->") > 0) Or _
                  (InStr(1, content, "[entry-id:: " & entryId & "]") > 0)
End Function

' ---------------------------------------------------------------------------
' MapPriority — maps Outlook importance to Obsidian Tasks priority emoji
' ---------------------------------------------------------------------------
Public Function MapPriority(importance As OlImportance) As String
    Select Case importance
        Case olImportanceHigh
            MapPriority = ChrW$(&H23EB) & " "    ' up-pointing double triangle
        Case olImportanceLow
            MapPriority = ChrW$(&HD83D) & ChrW$(&HDD3D) & " "  ' down-pointing small triangle
        Case Else
            MapPriority = ""
    End Select
End Function

' ---------------------------------------------------------------------------
' MapCategories — converts Outlook categories (comma-separated) to #tags
' ---------------------------------------------------------------------------
Public Function MapCategories(categories As String) As String
    If categories = "" Then
        MapCategories = ""
        Exit Function
    End If

    Dim parts() As String
    parts = Split(categories, ",")

    Dim result As String
    result = ""

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim tag As String
        tag = "#" & LCase$(Replace$(Trim$(parts(i)), " ", "-"))
        If Len(tag) > 1 Then
            If result <> "" Then result = result & " "
            result = result & tag
        End If
    Next i

    If result <> "" Then result = result & " "
    MapCategories = result
End Function

' ---------------------------------------------------------------------------
' GetShortHash — creates an 8-char hash from EntryID for dedup
' Uses SHA256 via .NET COM interop, falls back to simple hash
' ---------------------------------------------------------------------------
Public Function GetShortHash(entryId As String) As String
    On Error GoTo Fallback

    ' Try .NET SHA256 via COM
    Dim sha As Object
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")

    Dim utf8 As Object
    Set utf8 = CreateObject("System.Text.UTF8Encoding")

    Dim bytes() As Byte
    bytes = utf8.GetBytes_4(entryId)

    Dim hash() As Byte
    hash = sha.ComputeHash_2(bytes)

    ' Take first 4 bytes -> 8 hex chars
    Dim result As String
    result = ""
    Dim j As Long
    For j = 0 To 3
        result = result & LCase$(Right$("0" & Hex$(hash(j)), 2))
    Next j

    Set sha = Nothing
    Set utf8 = Nothing
    GetShortHash = result
    Exit Function

Fallback:
    ' Simple string hash fallback (CRC-style)
    Dim h As Long
    h = 5381
    Dim c As Long
    For c = 1 To Len(entryId)
        h = ((h * 33) Xor AscW(Mid$(entryId, c, 1))) And &H7FFFFFFF
    Next c
    GetShortHash = LCase$(Right$("0000000" & Hex$(h), 8))
End Function

' ---------------------------------------------------------------------------
' GetTargetFileName — returns Inbox.md or daily note filename
' ---------------------------------------------------------------------------
Private Function GetTargetFileName() As String
    If Settings.GetUseDailyNotes() Then
        Dim fmt As String
        fmt = Settings.GetDailyNotesFormat()
        If fmt = "" Then fmt = "yyyy-mm-dd"
        GetTargetFileName = Format$(Now, fmt) & ".md"
    Else
        GetTargetFileName = Settings.GetTaskFileName()
    End If
End Function

' ---------------------------------------------------------------------------
' GetBodyPreview — cleans email body and returns first N chars
' ---------------------------------------------------------------------------
Public Function GetBodyPreview(body As String, maxLength As Long) As String
    If Trim$(body) = "" Then
        GetBodyPreview = ""
        Exit Function
    End If

    Dim cleaned As String
    cleaned = body

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    ' Strip URLs
    re.Pattern = "https?://\S+"
    cleaned = re.Replace(cleaned, "")

    ' Strip invisible Unicode characters (zero-width spaces, combining marks, etc.)
    ' VBScript.RegExp doesn't support \uXXXX — use ChrW to build the character class
    re.Pattern = "[" & ChrW$(&H34F) & ChrW$(&H200B) & "-" & ChrW$(&H200F) & _
                 ChrW$(&H2028) & "-" & ChrW$(&H202F) & ChrW$(&HFEFF) & "]"
    cleaned = re.Replace(cleaned, "")

    ' Strip common email junk phrases
    Dim junkPhrases As Variant
    junkPhrases = Array("View Web Version", "View in browser", "View online", _
                        "Unsubscribe", "Click here", "Learn more", _
                        "Having trouble viewing", "Add us to your address book")
    Dim phrase As Variant
    For Each phrase In junkPhrases
        re.Pattern = EscapeRegex(CStr(phrase))
        cleaned = re.Replace(cleaned, "")
    Next phrase

    ' Strip leading/trailing quote characters
    cleaned = StripQuoteChars(cleaned)

    ' Collapse whitespace
    re.Pattern = "\s+"
    cleaned = Trim$(re.Replace(cleaned, " "))

    ' Strip angle-bracket patterns (HTML-like tags)
    re.Pattern = "<[^>]*>"
    cleaned = re.Replace(cleaned, "")
    cleaned = Trim$(re.Replace(cleaned, " "))

    Set re = Nothing

    If cleaned = "" Then
        GetBodyPreview = ""
        Exit Function
    End If

    If Len(cleaned) <= maxLength Then
        GetBodyPreview = cleaned
    Else
        GetBodyPreview = Left$(cleaned, maxLength) & "..."
    End If
End Function

' ---------------------------------------------------------------------------
' SanitizeForMarkdown — strip newlines and brackets
' ---------------------------------------------------------------------------
Public Function SanitizeForMarkdown(text As String) As String
    Dim result As String
    result = Replace$(text, vbCrLf, " ")
    result = Replace$(result, vbLf, " ")
    result = Replace$(result, vbCr, " ")
    result = Replace$(result, "[", "(")
    result = Replace$(result, "]", ")")
    SanitizeForMarkdown = result
End Function

' ===========================================================================
' Private helpers
' ===========================================================================

Private Function GetAttachmentCount(mail As Outlook.MailItem) As Long
    On Error GoTo ErrHandler
    GetAttachmentCount = mail.Attachments.Count
    Exit Function
ErrHandler:
    GetAttachmentCount = 0
End Function

Private Function SafeString(val As String, fallback As String) As String
    If val = "" Then
        SafeString = fallback
    Else
        SafeString = val
    End If
End Function

Private Function RTrimWhitespace(text As String) As String
    Dim i As Long
    For i = Len(text) To 1 Step -1
        If Mid$(text, i, 1) <> " " And Mid$(text, i, 1) <> vbTab Then
            RTrimWhitespace = Left$(text, i)
            Exit Function
        End If
    Next i
    RTrimWhitespace = ""
End Function

Private Function RemoveExtension(fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveExtension = Left$(fileName, pos - 1)
    Else
        RemoveExtension = fileName
    End If
End Function

Private Function EscapeRegex(text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([.^$*+?{}()\[\]\\|])"
    EscapeRegex = re.Replace(text, "\$1")
    Set re = Nothing
End Function

Private Function StripQuoteChars(text As String) As String
    Dim result As String
    result = Trim$(text)
    ' Strip leading/trailing ASCII and smart quotes
    Do While Len(result) > 0
        Dim first As String
        first = Left$(result, 1)
        If first = """" Or first = "'" Or first = ChrW$(&H201C) Or first = ChrW$(&H201D) Then
            result = Mid$(result, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(result) > 0
        Dim last As String
        last = Right$(result, 1)
        If last = """" Or last = "'" Or last = ChrW$(&H201C) Or last = ChrW$(&H201D) Then
            result = Left$(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop
    StripQuoteChars = result
End Function

' ---------------------------------------------------------------------------
' File I/O helpers
' ---------------------------------------------------------------------------
Private Sub WriteTextFile(path As String, content As String)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, content;
    Close #f
End Sub

Private Sub AppendTextFile(path As String, content As String)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, content;
    Close #f
End Sub

Private Function ReadTextFile(path As String) As String
    Dim f As Long
    f = FreeFile
    Dim content As String

    Open path For Input As #f
    If LOF(f) > 0 Then
        content = Input$(LOF(f), #f)
    Else
        content = ""
    End If
    Close #f

    ReadTextFile = content
End Function
