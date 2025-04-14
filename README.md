# Automated Client Assessment Alert System (Excel + Outlook)

This project automates client assessment email alerts using Excel VBA and Outlook. It checks if any assessments are due within 14 days and sends reminder emails to assigned caseworkers.

Internal use only — adapt freely for clinics, case management, and mental health agencies.

---

## Project Files

- `ClientAssessmentSchedule.xlsm`: Main Excel file with client data and embedded macro (consists of 5 sheets: AssessmentTypes (for use with data validation), Sheet1-Sheet4 (denoting each of the 4 case workers))
- Optional: `RunAssessmentAlerts.vbs` (if you want to launch silently via Task Scheduler--Still testing)

---

## Excel Sheet Layout (Sheets1-Sheets4)

| Column               | Type         | Description                                         |
|----------------------|--------------|-----------------------------------------------------|
| `Client Name`        | Manual       | Name of the client                                  |
| `DOB`                | Manual       | Date of birth                                       |
| `Chart No`           | Manual       | Client’s chart number                               |
| `Assessment`         | Dropdown     | Type of assessment (see list below)                 |
| `Last Completed`     | Manual Date  | Last date the assessment was completed             |
| `Next Due Date`      | Formula      | Automatically calculated from type + interval      |
| `Due Soon?`          | Formula      | TRUE if due in next 14 days                         |

### Supported Assessment Types & Intervals (sheet AssessmentType)
If these change, please feel free to edit them in the `ClientAssessmentSchedule.xlsm` workbook. 

| Assessment Type              | Interval |
|-----------------------------|----------|
| CANS Assessment             | 6 months |
| PHQ-9/Depression Scale      | 6 months |
| Treatment Plan              | 6 months |
| Psychosocial                | 12 months|
| MI State Report             | 12 months|

### Example Formula for `Next Due Date`:

```excel
=IF(D2="CANS Assessment", E2+180,
 IF(D2="PHQ-9/Depression Scale", E2+180,
 IF(D2="Treatment Plan", E2+180,
 IF(D2="Psychosocial", E2+365,
 IF(D2="MI State Report", E2+365, "")))))
```
Copy the formula down as necessary.

### Example Formula for `Due Soon?`:

```excel
=AND(F2-TODAY()<=14, F2-TODAY()>=0)
```
Copy the formula down as necessary.

## VBA Macro Code
`SendDueAssessmentAlerts` (Module)

In the VBA editor (`ALT + F11`) or go to Developer Mode in Excel (File -> Options -> Customize Ribbon (right hand column, check box for Developer). Once this is done, in the VBA editor, click Insert → Module. 

```vba
Sub SendDueAssessmentAlerts()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim alertBody As String
    Dim caseworkerEmail As String
    Dim sheetName As String

    ' Start Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If OutlookApp Is Nothing Then
        MsgBox "Outlook could not be started. Please make sure it's installed and you're signed in.", vbCritical
        Exit Sub
    End If

    ' Loop through each caseworker sheet
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name

        ' Map sheet names to email addresses
        Select Case sheetName
            Case "Alice"
                caseworkerEmail = "alice@example.com" ' Replace with the actual email address
            Case "Bob"
                caseworkerEmail = "bob@example.com" ' Replace with the actual email address
            Case "Carlos"
                caseworkerEmail = "carlos@example.com" ' Replace with the actual email address
            Case "Diana"
                caseworkerEmail = "diana@example.com"  ' Replace with the actual email address
            Case Else
                GoTo NextSheet ' Skip unknown sheets
        End Select

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        alertBody = ""

        ' Build the body of the email
        For i = 2 To lastRow
            If ws.Cells(i, "G").Value = True Then ' Column G = “Due Soon?”
                alertBody = alertBody & _
                    "• " & ws.Cells(i, "A").Value & " — Due: " & _
                    Format(ws.Cells(i, "F").Value, "mmm dd, yyyy") & vbCrLf
            End If
        Next i

        ' If there are any alerts, send email
        If alertBody <> "" Then
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = caseworkerEmail
                .Subject = "Upcoming Assessments Due (" & sheetName & ")"
                .Body = "Hello " & sheetName & "," & vbCrLf & vbCrLf & _
                        "The following client assessments are due soon:" & vbCrLf & vbCrLf & _
                        alertBody & vbCrLf & "Please follow up accordingly." & vbCrLf & vbCrLf & _
                        "- Automated Alert System"
                .Send
            End With
        End If

NextSheet:
    Next ws

    MsgBox "Emails sent to all caseworkers.", vbInformation
End Sub
```


### `Workbook_Open()` (Inside `ThisWorkbook`)

While in the VBA editor, on the left hand side find `ThisWorkbook` and double-click it and add the following code:

```vba
Private Sub Workbook_Open()
    Call SendDueAssessmentAlerts
End Sub
```

Save. This would be a good time to ensure that you are saving your work as an `Excel Macro-Enabled Workbook (*.xlsm)`. 

## Automate via Windows Task Scheduler

1. Prerequisites
- Save the workbook as `ClientAssessmentSchedule.xlsm`
- Macro must be enabled
- Outlook must be installed and configured

2. Create Task
Open Task Scheduler
- Create a new basic task:
- Name: Weekly Email Assessment Alert
- Trigger: Weekly on Mondays
- Time: e.g., 8:00 AM

3. Action: Start a program (ensure Administrator privileges are enabled)
- Program:
```text
"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
```
- Arguments:
```text
"C:\Path\To\ClientAssessmentSchedule.xlsm"
```

## Testing
To simulate a test due date for the CANS Assessment:

- Last Completed: 10/16/2024

- That creates a due date of 4/14/2025 (today), which is within the 14-day window and should trigger the email

## Future Optional Enhancements
- Use a .vbs launcher for silent runs

- Add error logging or email confirmations

- Pull caseworker emails from a dedicated column
