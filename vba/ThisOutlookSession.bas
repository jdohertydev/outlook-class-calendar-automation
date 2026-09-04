Option Explicit

' Sanitised public version of the Outlook calendar automation.
' Paste this code into the ThisOutlookSession module in classic Outlook for Windows.
'
' Historical behaviour retained:
' - exact subject match
' - Class ID and date/time extracted from the email body
' - +1 hour adjustment
' - skip classes already in the past
' - duplicate check by Class ID in the default calendar
' - 25-minute appointment
' - 15-minute reminder
' - move successfully processed email to Inbox\Processed
'
' Public-version improvements:
' - mailbox identity is configurable rather than hard-coded
' - Inbox sweep iterates backwards before moving messages
' - date/time parsing is explicit rather than locale-dependent
' - lightweight error output is sent to the Immediate window

Private Const SOURCE_MAILBOX_NAME As String = "Teaching Account"
Private Const TRIGGER_SUBJECT As String = "Schedule update: class added"
Private Const PROCESSED_FOLDER_NAME As String = "Processed"
Private Const EVENT_SUBJECT_PREFIX As String = "Private class ("

Private Const CLASS_DURATION_MINUTES As Long = 25
Private Const REMINDER_MINUTES As Long = 15
Private Const TIME_OFFSET_HOURS As Long = 1

Private Sub Application_Startup()
    ProcessInboxSweep
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim entryIds() As String
    Dim i As Long
    Dim itm As Object
    Dim mail As Outlook.MailItem

    entryIds = Split(EntryIDCollection, ",")

    For i = LBound(entryIds) To UBound(entryIds)
        Set itm = Nothing

        On Error Resume Next
        Set itm = Application.Session.GetItemFromID(entryIds(i))
        On Error GoTo 0

        If Not itm Is Nothing Then
            If TypeOf itm Is Outlook.MailItem Then
                Set mail = itm

                If StrComp(mail.Subject, TRIGGER_SUBJECT, vbTextCompare) = 0 Then
                    ProcessClassEmail mail
                End If
            End If
        End If
    Next i
End Sub

Private Sub ProcessInboxSweep()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim itm As Object
    Dim mail As Outlook.MailItem
    Dim i As Long

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.Folders(SOURCE_MAILBOX_NAME).Folders("Inbox")

    ' The original version used For Each while moving processed messages out
    ' of the Inbox. Iterating backwards avoids skipping items as the collection
    ' changes during the sweep.
    For i = inbox.Items.Count To 1 Step -1
        Set itm = inbox.Items(i)

        If TypeOf itm Is Outlook.MailItem Then
            Set mail = itm

            If StrComp(mail.Subject, TRIGGER_SUBJECT, vbTextCompare) = 0 Then
                ProcessClassEmail mail
            End If
        End If
    Next i
End Sub

Private Sub ProcessClassEmail(ByVal mail As Outlook.MailItem)
    On Error GoTo ProcessingError

    Dim classID As String
    Dim dateTimeStr As String
    Dim classDate As Date

    classID = ExtractFirstGroup(mail.Body, "Class ID:\s*(\w+)")
    If Len(classID) = 0 Then Exit Sub

    dateTimeStr = ExtractFirstGroup( _
        mail.Body, _
        "Date/time:\s*(\d{2}\.\d{2}\.\d{4}\s*\d{2}:\d{2})" _
    )
    If Len(dateTimeStr) = 0 Then Exit Sub

    If Not TryParseProviderDateTime(dateTimeStr, classDate) Then Exit Sub

    ' Historical VBA implementation used a fixed +1 hour adjustment.
    classDate = DateAdd("h", TIME_OFFSET_HOURS, classDate)

    If classDate < Now Then Exit Sub
    If CalendarContainsClassID(classID) Then Exit Sub

    CreateClassAppointment classID, classDate
    MoveToProcessedFolder mail
    Exit Sub

ProcessingError:
    Debug.Print "Class calendar automation error " & Err.Number & ": " & Err.Description
End Sub

Private Function ExtractFirstGroup(ByVal sourceText As String, ByVal pattern As String) As String
    Dim regex As Object
    Dim matches As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = pattern

    If regex.Test(sourceText) Then
        Set matches = regex.Execute(sourceText)
        ExtractFirstGroup = matches(0).SubMatches(0)
    Else
        ExtractFirstGroup = vbNullString
    End If
End Function

Private Function TryParseProviderDateTime( _
    ByVal value As String, _
    ByRef parsedDate As Date _
) As Boolean
    On Error GoTo ParseFailed

    Dim dateAndTime() As String
    Dim dateParts() As String
    Dim timeParts() As String

    dateAndTime = Split(Trim$(value), " ")
    If UBound(dateAndTime) <> 1 Then GoTo ParseFailed

    dateParts = Split(dateAndTime(0), ".")
    timeParts = Split(dateAndTime(1), ":")

    If UBound(dateParts) <> 2 Then GoTo ParseFailed
    If UBound(timeParts) <> 1 Then GoTo ParseFailed

    parsedDate = DateSerial( _
        CLng(dateParts(2)), _
        CLng(dateParts(1)), _
        CLng(dateParts(0)) _
    ) + TimeSerial( _
        CLng(timeParts(0)), _
        CLng(timeParts(1)), _
        0 _
    )

    TryParseProviderDateTime = True
    Exit Function

ParseFailed:
    TryParseProviderDateTime = False
End Function

Private Function CalendarContainsClassID(ByVal classID As String) As Boolean
    Dim ns As Outlook.NameSpace
    Dim calendarFolder As Outlook.Folder
    Dim itm As Object
    Dim appt As Outlook.AppointmentItem

    Set ns = Application.GetNamespace("MAPI")
    Set calendarFolder = ns.GetDefaultFolder(olFolderCalendar)

    For Each itm In calendarFolder.Items
        If TypeOf itm Is Outlook.AppointmentItem Then
            Set appt = itm

            If InStr(1, appt.Subject, classID, vbTextCompare) > 0 Then
                CalendarContainsClassID = True
                Exit Function
            End If
        End If
    Next itm

    CalendarContainsClassID = False
End Function

Private Sub CreateClassAppointment(ByVal classID As String, ByVal classDate As Date)
    Dim ns As Outlook.NameSpace
    Dim calendarFolder As Outlook.Folder
    Dim appt As Outlook.AppointmentItem

    Set ns = Application.GetNamespace("MAPI")
    Set calendarFolder = ns.GetDefaultFolder(olFolderCalendar)
    Set appt = calendarFolder.Items.Add(olAppointmentItem)

    With appt
        .Start = classDate
        .End = DateAdd("n", CLASS_DURATION_MINUTES, classDate)
        .Subject = EVENT_SUBJECT_PREFIX & classID & ")"
        .ReminderSet = True
        .ReminderMinutesBeforeStart = REMINDER_MINUTES
        .Save
    End With
End Sub

Private Sub MoveToProcessedFolder(ByVal mail As Outlook.MailItem)
    Dim processedFolder As Outlook.Folder

    Set processedFolder = mail.Parent.Folders(PROCESSED_FOLDER_NAME)
    mail.Move processedFolder
End Sub

Public Sub Run_ProcessClassEmails_Manual()
    ProcessInboxSweep
    MsgBox "Finished processing schedule update emails.", vbInformation
End Sub
