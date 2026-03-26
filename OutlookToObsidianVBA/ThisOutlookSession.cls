Attribute VB_Name = "ThisOutlookSession"
' ==========================================================================
' OutlookToObsidian — ThisOutlookSession
'
' Paste this code into the existing ThisOutlookSession module in Outlook VBA.
' (Alt+F11 → Project1 → Microsoft Outlook Objects → ThisOutlookSession)
'
' Handles:
'   - Application startup (first-run vault path prompt)
'   - Custom toolbar creation with two buttons
'   - Email processing (instant and detailed modes)
'   - Obsidian URI launch for notification
' ==========================================================================
Option Explicit

Private Const TOOLBAR_NAME As String = "OutlookToObsidian"
Private Const BTN_INSTANT_TAG As String = "OTO_Instant"
Private Const BTN_DETAILED_TAG As String = "OTO_Detailed"

' WithEvents for button click handlers
Private WithEvents m_btnInstant As Office.CommandBarButton
Private WithEvents m_btnDetailed As Office.CommandBarButton

' ---------------------------------------------------------------------------
' Application_Startup — runs when Outlook starts
' ---------------------------------------------------------------------------
Private Sub Application_Startup()
    ' First-run: prompt for vault path if not configured
    If Settings.GetVaultPath() = "" Then
        Dim picked As Boolean
        picked = Settings.PromptForVaultPath()
        If Not picked Then
            MsgBox "OutlookToObsidian: No vault selected. Right-click menu will not work until a vault is configured." & vbLf & _
                   "To configure later, run the macro 'ConfigureVault'.", _
                   vbExclamation, "OutlookToObsidian"
        End If
    End If

    ' Create or find the toolbar
    SetupToolbar
End Sub

' ---------------------------------------------------------------------------
' SetupToolbar — creates a CommandBar with two buttons
' ---------------------------------------------------------------------------
Private Sub SetupToolbar()
    On Error Resume Next

    ' Remove existing toolbar to avoid duplicates
    Application.ActiveExplorer.CommandBars(TOOLBAR_NAME).Delete
    On Error GoTo 0

    ' Create new toolbar
    Dim bar As Office.CommandBar
    Set bar = Application.ActiveExplorer.CommandBars.Add(Name:=TOOLBAR_NAME, temporary:=True)
    bar.Visible = True

    ' Button 1: Create Task (instant)
    Set m_btnInstant = bar.Controls.Add(Type:=msoControlButton)
    With m_btnInstant
        .Caption = "Create Task in Obsidian"
        .Tag = BTN_INSTANT_TAG
        .Style = msoButtonIconAndCaption
        .FaceId = 1087    ' Task icon
        .TooltipText = "Create an Obsidian task from selected email(s)"
    End With

    ' Button 2: Create Task (detailed)
    Set m_btnDetailed = bar.Controls.Add(Type:=msoControlButton)
    With m_btnDetailed
        .Caption = "Create Task (detailed...)"
        .Tag = BTN_DETAILED_TAG
        .Style = msoButtonIconAndCaption
        .FaceId = 1088    ' Task accept icon
        .TooltipText = "Create an Obsidian task with custom options"
    End With
End Sub

' ---------------------------------------------------------------------------
' Button click handlers
' ---------------------------------------------------------------------------
Private Sub m_btnInstant_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ProcessEmails False
End Sub

Private Sub m_btnDetailed_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ProcessEmails True
End Sub

' ---------------------------------------------------------------------------
' ProcessEmails — main processing loop (ported from ContextMenuRibbon.cs)
' ---------------------------------------------------------------------------
Private Sub ProcessEmails(detailed As Boolean)
    On Error GoTo ErrHandler

    Dim explorer As Outlook.explorer
    Set explorer = Application.ActiveExplorer

    If explorer Is Nothing Then Exit Sub
    If explorer.Selection Is Nothing Then Exit Sub
    If explorer.Selection.Count = 0 Then
        MsgBox "Please select one or more emails first.", vbInformation, "OutlookToObsidian"
        Exit Sub
    End If

    ' Check vault is configured
    Dim vaultPath As String
    vaultPath = Settings.GetVaultPath()
    If vaultPath = "" Or Dir(vaultPath, vbDirectory) = "" Then
        MsgBox "Obsidian vault path is not configured or does not exist." & vbLf & _
               "Run the macro 'ConfigureVault' to set it up.", _
               vbWarning, "OutlookToObsidian"
        Exit Sub
    End If

    Dim created As Long
    Dim skipped As Long
    Dim lastSubject As String
    Dim lastFileName As String
    created = 0
    skipped = 0

    Dim i As Long
    For i = 1 To explorer.Selection.Count
        Dim item As Object
        Set item = explorer.Selection.item(i)

        If TypeOf item Is Outlook.MailItem Then
            Dim mail As Outlook.MailItem
            Set mail = item

            ' Duplicate check
            If TaskCreator.IsDuplicate(mail.entryId) Then
                skipped = skipped + 1
                GoTo NextItem
            End If

            Dim markdown As String

            If detailed Then
                ' Show detailed entry dialog
                Dim defaults As TaskOptions
                defaults = TaskCreator.BuildDefaultOptions(mail)

                Dim formResult As TaskOptions
                If TaskEntryForm.ShowTaskEntryForm(defaults, formResult) Then
                    ' Preserve attachment count from mail
                    formResult.AttachmentCount = defaults.AttachmentCount
                    markdown = TaskCreator.CreateTask(mail, formResult)
                Else
                    ' User cancelled
                    GoTo NextItem
                End If
            Else
                Dim emptyOpts As TaskOptions
                markdown = TaskCreator.CreateTask(mail, emptyOpts)
            End If

            lastFileName = TaskCreator.AppendToVault(markdown)
            lastSubject = SafeStr(mail.Subject, "(no subject)")
            created = created + 1
        End If

NextItem:
        Set item = Nothing
        Set mail = Nothing
    Next i

    ' Show result notification
    If created > 0 Or skipped > 0 Then
        Dim msg As String
        If created = 1 Then
            msg = "Task created: " & TruncateStr(lastSubject, 40)
        ElseIf created > 1 Then
            msg = created & " tasks created in Obsidian"
        Else
            msg = ""
        End If

        If skipped > 0 Then
            Dim skipMsg As String
            skipMsg = skipped & " duplicate(s) skipped"
            If msg = "" Then
                msg = skipMsg
            Else
                msg = msg & " | " & skipMsg
            End If
        End If

        ' Offer to open in Obsidian
        If created > 0 And lastFileName <> "" Then
            Dim vaultName As String
            vaultName = Settings.GetVaultName()
            Dim answer As VbMsgBoxResult
            answer = MsgBox(msg & vbLf & vbLf & "Open in Obsidian?", _
                           vbInformation + vbYesNo, "OutlookToObsidian")
            If answer = vbYes Then
                OpenInObsidian vaultName, lastFileName
            End If
        Else
            MsgBox msg, vbInformation, "OutlookToObsidian"
        End If
    End If

    Exit Sub

ErrHandler:
    MsgBox "Could not create task:" & vbLf & Err.Description, _
           vbCritical, "OutlookToObsidian"
End Sub

' ---------------------------------------------------------------------------
' OpenInObsidian — launches obsidian:// URI
' ---------------------------------------------------------------------------
Private Sub OpenInObsidian(vaultName As String, fileName As String)
    On Error Resume Next
    Dim fileWithoutExt As String
    fileWithoutExt = RemoveExt(fileName)

    Dim uri As String
    uri = "obsidian://open?vault=" & UrlEncode(vaultName) & "&file=" & UrlEncode(fileWithoutExt)

    Shell "cmd /c start """" """ & uri & """", vbHide
End Sub

' ---------------------------------------------------------------------------
' ConfigureVault — public macro users can run to (re)set vault path
' ---------------------------------------------------------------------------
Public Sub ConfigureVault()
    Dim picked As Boolean
    picked = Settings.PromptForVaultPath()
    If picked Then
        MsgBox "Vault path set to: " & Settings.GetVaultPath(), vbInformation, "OutlookToObsidian"
    Else
        MsgBox "No folder selected. Vault path unchanged.", vbInformation, "OutlookToObsidian"
    End If
End Sub

' ===========================================================================
' Private helpers
' ===========================================================================

Private Function SafeStr(val As String, fallback As String) As String
    If val = "" Then
        SafeStr = fallback
    Else
        SafeStr = val
    End If
End Function

Private Function TruncateStr(text As String, maxLength As Long) As String
    If text = "" Or Len(text) <= maxLength Then
        TruncateStr = text
    Else
        TruncateStr = Left$(text, maxLength) & "..."
    End If
End Function

Private Function RemoveExt(fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveExt = Left$(fileName, pos - 1)
    Else
        RemoveExt = fileName
    End If
End Function

Private Function UrlEncode(text As String) As String
    Dim result As String
    result = ""
    Dim i As Long
    For i = 1 To Len(text)
        Dim c As String
        c = Mid$(text, i, 1)
        Dim code As Long
        code = AscW(c)
        If (code >= 48 And code <= 57) Or _
           (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           c = "-" Or c = "_" Or c = "." Or c = "~" Then
            result = result & c
        Else
            ' Encode as %XX
            If code < 128 Then
                result = result & "%" & Right$("0" & Hex$(code), 2)
            Else
                ' For Unicode chars, encode as UTF-8 bytes
                Dim utf8Bytes() As Byte
                utf8Bytes = StrConv(c, vbFromUnicode)
                Dim b As Long
                For b = LBound(utf8Bytes) To UBound(utf8Bytes)
                    result = result & "%" & Right$("0" & Hex$(utf8Bytes(b)), 2)
                Next b
            End If
        End If
    Next i
    UrlEncode = result
End Function
