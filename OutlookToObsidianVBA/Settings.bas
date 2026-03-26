Attribute VB_Name = "Settings"
Option Explicit

' Registry-based settings for OutlookToObsidian VBA macro.
' Uses VBA's built-in SaveSetting/GetSetting (stores in HKCU\Software\VB and VBA Program Settings).

Private Const APP_NAME As String = "OutlookToObsidian"
Private Const SECTION As String = "Settings"

Public Function GetVaultPath() As String
    GetVaultPath = GetSetting(APP_NAME, SECTION, "VaultPath", "")
End Function

Public Sub SetVaultPath(ByVal path As String)
    SaveSetting APP_NAME, SECTION, "VaultPath", path
End Sub

Public Function GetTaskFileName() As String
    Dim val As String
    val = GetSetting(APP_NAME, SECTION, "TaskFileName", "Inbox.md")
    If val = "" Then val = "Inbox.md"
    GetTaskFileName = val
End Function

Public Sub SetTaskFileName(ByVal fileName As String)
    SaveSetting APP_NAME, SECTION, "TaskFileName", fileName
End Sub

Public Function GetUseDailyNotes() As Boolean
    GetUseDailyNotes = (GetSetting(APP_NAME, SECTION, "UseDailyNotes", "False") = "True")
End Function

Public Sub SetUseDailyNotes(ByVal value As Boolean)
    SaveSetting APP_NAME, SECTION, "UseDailyNotes", IIf(value, "True", "False")
End Sub

Public Function GetDailyNotesFormat() As String
    Dim val As String
    val = GetSetting(APP_NAME, SECTION, "DailyNotesFormat", "yyyy-mm-dd")
    If val = "" Then val = "yyyy-mm-dd"
    GetDailyNotesFormat = val
End Function

Public Sub SetDailyNotesFormat(ByVal fmt As String)
    SaveSetting APP_NAME, SECTION, "DailyNotesFormat", fmt
End Sub

Public Function GetVaultName() As String
    Dim val As String
    val = GetSetting(APP_NAME, SECTION, "VaultName", "")
    If val = "" Then
        Dim vp As String
        vp = GetVaultPath()
        If vp <> "" Then
            ' Extract folder name from path
            If Right$(vp, 1) = "\" Then vp = Left$(vp, Len(vp) - 1)
            Dim pos As Long
            pos = InStrRev(vp, "\")
            If pos > 0 Then
                val = Mid$(vp, pos + 1)
            Else
                val = vp
            End If
        End If
    End If
    GetVaultName = val
End Function

Public Sub SetVaultName(ByVal name As String)
    SaveSetting APP_NAME, SECTION, "VaultName", name
End Sub

Public Sub PromptForVaultPath()
    Dim fd As Office.FileDialog
    Set fd = Application.Session.GetDefaultFolder(olFolderInbox).Application.FileDialog(msoFileDialogFolderPicker)

    fd.Title = "Select your Obsidian vault folder"
    fd.ButtonName = "Select Vault"

    If fd.Show = -1 Then
        SetVaultPath fd.SelectedItems(1)
    End If

    Set fd = Nothing
End Sub
