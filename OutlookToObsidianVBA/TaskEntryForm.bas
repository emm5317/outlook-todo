Attribute VB_Name = "TaskEntryForm"
Option Explicit

' Programmatically creates a UserForm for detailed task entry.
' Ported from OutlookToObsidian/TaskEntryForm.cs
'
' Usage:
'   Dim opts As TaskOptions
'   opts = BuildDefaultOptions(mail)
'   Dim result As TaskOptions
'   If ShowTaskEntryForm(opts, result) Then
'       ' User clicked Create — result contains form values
'   End If

' ---------------------------------------------------------------------------
' ShowTaskEntryForm — collects task details via InputBox dialogs
' Returns True if user completed the form, False if cancelled
' ---------------------------------------------------------------------------
Public Function ShowTaskEntryForm(defaults As TaskOptions, ByRef result As TaskOptions) As Boolean
    ShowTaskEntryForm = ShowDialogSequence(defaults, result)
End Function

' ---------------------------------------------------------------------------
' ShowDialogSequence — uses MsgBox/InputBox to collect task details
' Simpler than a UserForm, works without .frm binary files
' ---------------------------------------------------------------------------
Private Function ShowDialogSequence(defaults As TaskOptions, ByRef result As TaskOptions) As Boolean
    Dim subj As String
    Dim dueDate As String
    Dim Priority As String
    Dim Tags As String
    Dim Notes As String

    ' Subject
    subj = InputBox("Subject:", "Create Task in Obsidian", defaults.Subject)
    If subj = "" And defaults.Subject <> "" Then
        ' User pressed Cancel (InputBox returns "" for cancel)
        ShowDialogSequence = False
        Exit Function
    End If
    If subj = "" Then subj = defaults.Subject

    ' Due Date
    dueDate = InputBox("Due date (yyyy-mm-dd) or leave blank for none:", _
                        "Create Task in Obsidian", defaults.DueDate)
    ' Validate date format if provided
    If dueDate <> "" Then
        If Not IsValidDate(dueDate) Then
            MsgBox "Invalid date format. Using no due date.", vbExclamation, "OutlookToObsidian"
            dueDate = ""
        End If
    End If

    ' Priority
    Dim priorityChoice As String
    priorityChoice = InputBox("Priority (1=High, 2=Medium, 3=Low, Enter=Normal):" & vbLf & _
                              "1 - High " & ChrW$(&H23EB) & vbLf & _
                              "2 - Medium " & ChrW$(&HD83D) & ChrW$(&HDD3C) & vbLf & _
                              "3 - Low " & ChrW$(&HD83D) & ChrW$(&HDD3D) & vbLf & _
                              "(blank = Normal)", _
                              "Create Task in Obsidian", PriorityToIndex(defaults.Priority))
    Priority = IndexToPriority(priorityChoice)

    ' Tags
    Tags = InputBox("Tags (space-separated #tags):", "Create Task in Obsidian", defaults.Tags)

    ' Notes
    Notes = InputBox("Notes (optional):", "Create Task in Obsidian", defaults.Notes)

    ' Build result
    result.HasOptions = True
    result.Subject = SanitizeForMarkdown(Trim$(subj))
    result.DueDate = dueDate
    result.Priority = Priority
    result.Tags = Tags
    result.Notes = Trim$(Notes)
    result.AttachmentCount = defaults.AttachmentCount

    ShowDialogSequence = True
End Function

' ---------------------------------------------------------------------------
' Priority helpers
' ---------------------------------------------------------------------------
Private Function PriorityToIndex(Priority As String) As String
    If InStr(1, Priority, ChrW$(&H23EB)) > 0 Then
        PriorityToIndex = "1"
    ElseIf InStr(1, Priority, ChrW$(&HD83D) & ChrW$(&HDD3C)) > 0 Then
        PriorityToIndex = "2"
    ElseIf InStr(1, Priority, ChrW$(&HD83D) & ChrW$(&HDD3D)) > 0 Then
        PriorityToIndex = "3"
    Else
        PriorityToIndex = ""
    End If
End Function

Private Function IndexToPriority(choice As String) As String
    Select Case Trim$(choice)
        Case "1"
            IndexToPriority = ChrW$(&H23EB) & " "
        Case "2"
            IndexToPriority = ChrW$(&HD83D) & ChrW$(&HDD3C) & " "
        Case "3"
            IndexToPriority = ChrW$(&HD83D) & ChrW$(&HDD3D) & " "
        Case Else
            IndexToPriority = ""
    End Select
End Function

Private Function IsValidDate(dateStr As String) As Boolean
    ' Check yyyy-mm-dd format
    If Len(dateStr) <> 10 Then
        IsValidDate = False
        Exit Function
    End If
    If Mid$(dateStr, 5, 1) <> "-" Or Mid$(dateStr, 8, 1) <> "-" Then
        IsValidDate = False
        Exit Function
    End If
    On Error GoTo Invalid
    Dim d As Date
    d = CDate(dateStr)
    IsValidDate = True
    Exit Function
Invalid:
    IsValidDate = False
End Function
