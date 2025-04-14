# Automated Client Assessment Alert System (Excel + Outlook)

This project automates client assessment email alerts using Excel VBA and Outlook. It checks if any assessments are due within 14 days and sends reminder emails to assigned caseworkers.

---

## Project Files

- `ClientAssessmentSchedule.xlsm`: Main Excel file with client data and embedded macro
- Optional: `RunAssessmentAlerts.vbs` (if you want to launch silently via Task Scheduler)

---

## Excel Sheet Layout

| Column               | Type         | Description                                         |
|----------------------|--------------|-----------------------------------------------------|
| `Client Name`        | Manual       | Name of the client                                  |
| `DOB`                | Manual       | Date of birth                                       |
| `Chart No`           | Manual       | Client’s chart number                               |
| `Assessment`         | Dropdown     | Type of assessment (see list below)                 |
| `Last Completed`     | Manual Date  | Last date the assessment was completed             |
| `Next Due Date`      | Formula      | Automatically calculated from type + interval      |
| `Due Soon?`          | Formula      | TRUE if due in next 14 days                         |

### Supported Assessment Types & Intervals

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

### Example Formula for `Due Soon?`:

```excel
=AND(F2-TODAY()<=14, F2-TODAY()>=0)
```

## VBA Macro Code
`SendDueAssessmentAlerts` (Module)

```vba
Sub SendDueAssessmentAlerts()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

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

    Set ws = ThisWorkbook.Sheets("Sheet 1") ' Change to your actual sheet name
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, "G").Value = True Then ' Column G = “Due Soon?”
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = "caseworker@example.com" ' Update as needed
                .Subject = "Assessment Due for " & ws.Cells(i, "A").Value
                .Body = "Reminder: An assessment is due for " & ws.Cells(i, "A").Value & _
                        " on " & ws.Cells(i, "F").Value & ". Please follow up."
                .Send
            End With
        End If
    Next i

    MsgBox "Emails sent!", vbInformation
End Sub
```

### `Workbook_Open()` (Inside `ThisWorkbook`)

```vba
Private Sub Workbook_Open()
    Call SendDueAssessmentAlerts
End Sub
```

## Automate via Windows Task Scheduler

1. Prerequisites
- Save the workbook as ClientAssessmentSchedule.xlsm

- Macro must be enabled

- Outlook must be installed and configured

2. Create Task
Open Task Scheduler

- Create a new basic task:

- Name: Weekly Email Assessment Alert

- Trigger: Weekly on Mondays

- Time: e.g., 8:00 AM

3. Action: Start a program

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
