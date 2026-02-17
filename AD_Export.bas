Option Explicit

' Optional reference for early-bound ADO:
'   Tools -> References -> "Microsoft ActiveX Data Objects 2.8 Library"
' This module uses late binding for ADO so the reference is not required.

Private Const ADS_SCOPE_SUBTREE As Long = 2
Private Const CSV_FILENAME As String = "allUsers_allSubdomains.csv"
Private Const DRS_FILENAME As String = "ManagersAndDRs.csv"

Private progressForm As Object

Public Sub ShowUserForm()
    If progressForm Is Nothing Then
        On Error Resume Next
        Set progressForm = VBA.UserForms.Add("ExtractionProgressBox")
        On Error GoTo 0
    End If

    If Not progressForm Is Nothing Then
        progressForm.Show vbModeless
    End If
End Sub

Public Sub WriteLineToUserForm(ByVal text As String)
    If progressForm Is Nothing Then Exit Sub

    On Error Resume Next
    progressForm.TextBox1.Text = progressForm.TextBox1.Text & text & vbCrLf
    progressForm.TextBox1.SelStart = Len(progressForm.TextBox1.Text)
    DoEvents
    On Error GoTo 0
End Sub

Public Sub TestUserForm()
    Dim i As Long

    ShowUserForm
    For i = 1 To 10
        WriteLineToUserForm "Processing item " & i
        Application.Wait (Now + TimeValue("0:00:01"))
        DoEvents
    Next i

    WriteLineToUserForm "Finished processing items"
    WriteLineToUserForm "Feel free to close this window"
End Sub

Public Sub ExportADUsersToCSV()
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object

    Dim fileMain As Integer
    Dim fileDR As Integer
    Dim mainPath As String
    Dim drPath As String

    Dim subdomains As Variant
    Dim subdomain As Variant

    Dim startTime As Date
    Dim endTime As Date
    Dim elapsedSeconds As Double

    Dim recordCount As Long
    Dim directReportCount As Long
    Dim recordsPerSecond As Double
    Dim recordsPerMinute As Double

    Dim fullDirectReports As String

    On Error GoTo CleanFail

    ShowUserForm
    WriteLineToUserForm "Process started at " & Format(Now, "hh:mm:ss AM/PM")

    subdomains = Array( _
        "canada.root.corp/DC=canada,DC=root,DC=corp", _
        "root.corp/DC=root,DC=corp", _
        "DC=accounts,DC=root,DC=corp" _
    )

    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")

    conn.Provider = "ADsDSOObject"
    conn.Open "Active Directory Provider"

    Set cmd.ActiveConnection = conn
    cmd.Properties("Page Size") = 1000
    cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE

    mainPath = ThisWorkbook.Path & Application.PathSeparator & CSV_FILENAME
    drPath = ThisWorkbook.Path & Application.PathSeparator & DRS_FILENAME

    fileMain = FreeFile
    Open mainPath For Output As #fileMain

    fileDR = FreeFile
    Open drPath For Output As #fileDR

    WriteLine fileMain, "distinguishedName|manager|displayName|title|company|department|mail|sAMAccountName|msExchHideFromAddressLists|DRs|directReports|subdomain"
    WriteLine fileDR, "distinguishedName|DR_count|directReports"

    startTime = Now
    WriteLineToUserForm "Connection established. Starting loop at " & Format(startTime, "hh:mm:ss AM/PM")
    DoEvents

    recordCount = 0

    For Each subdomain In subdomains
        WriteLineToUserForm "Querying subdomain: " & CStr(subdomain)

        cmd.CommandText = "SELECT distinguishedName, manager, displayName, title, company, department, mail, sAMAccountName, msExchHideFromAddressLists, directReports " & _
                          "FROM 'LDAP://" & CStr(subdomain) & "' WHERE objectCategory='user'"

        Set rs = cmd.Execute

        Do While Not rs.EOF
            fullDirectReports = BuildDirectReportsString(rs.Fields("directReports").Value, directReportCount)

            If directReportCount > 0 Then
                WriteLine fileDR, SafeFieldValue(rs, "distinguishedName") & "|" & CStr(directReportCount) & "|{" & fullDirectReports & "}"
            End If

            WriteLine fileMain, _
                SafeFieldValue(rs, "distinguishedName") & "|" & _
                SafeFieldValue(rs, "manager") & "|" & _
                SafeFieldValue(rs, "displayName") & "|" & _
                SafeFieldValue(rs, "title") & "|" & _
                SafeFieldValue(rs, "company") & "|" & _
                SafeFieldValue(rs, "department") & "|" & _
                SafeFieldValue(rs, "mail") & "|" & _
                SafeFieldValue(rs, "sAMAccountName") & "|" & _
                SafeFieldValue(rs, "msExchHideFromAddressLists") & "|" & _
                CStr(directReportCount) & "|{" & fullDirectReports & "}|" & _
                SanitizeField(CStr(subdomain))

            recordCount = recordCount + 1

            If (recordCount Mod 5000) = 0 Then
                elapsedSeconds = (Now - startTime) * 24# * 60# * 60#
                If elapsedSeconds > 0 Then
                    recordsPerSecond = recordCount / elapsedSeconds
                Else
                    recordsPerSecond = 0#
                End If
                WriteLineToUserForm "Processed " & recordCount & " records. Records per second: " & Format(recordsPerSecond, "0.00") & "/sec"
                DoEvents
            End If

            rs.MoveNext
            DoEvents
        Loop

        If Not rs Is Nothing Then
            If rs.State <> 0 Then rs.Close
            Set rs = Nothing
        End If
    Next subdomain

    endTime = Now
    elapsedSeconds = (endTime - startTime) * 24# * 60# * 60#

    If elapsedSeconds > 0 Then
        recordsPerSecond = recordCount / elapsedSeconds
    Else
        recordsPerSecond = 0#
    End If
    recordsPerMinute = recordsPerSecond * 60#

    WriteLineToUserForm "Query finished at " & Format(endTime, "hh:mm:ss AM/PM")
    WriteLineToUserForm "Total records: " & Format(recordCount, "0")
    WriteLineToUserForm "Total seconds: " & Format(elapsedSeconds, "0.00")
    WriteLineToUserForm "Total minutes: " & Format((elapsedSeconds / 60#), "0.00")
    WriteLineToUserForm "Records per second: " & Format(recordsPerSecond, "0.00") & "/sec"
    WriteLineToUserForm "Records per minute: " & Format(recordsPerMinute, "0.00") & "/min"
    WriteLineToUserForm "Total AD users exported: " & Format(recordCount, "0")
    WriteLineToUserForm "Feel free to close this window"

CleanExit:
    On Error Resume Next
    If fileMain > 0 Then Close #fileMain
    If fileDR > 0 Then Close #fileDR
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Sub

CleanFail:
    WriteLineToUserForm "Export failed: " & Err.Number & " - " & Err.Description
    MsgBox "ExportADUsersToCSV failed: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub WriteLine(ByVal fileNumber As Integer, ByVal line As String)
    Print #fileNumber, line
End Sub

Private Function BuildDirectReportsString(ByVal fieldValue As Variant, ByRef directReportCount As Long) As String
    Dim reports As Variant
    Dim i As Long
    Dim currentValue As String
    Dim result As String

    directReportCount = 0
    result = vbNullString

    If IsNull(fieldValue) Then
        BuildDirectReportsString = vbNullString
        Exit Function
    End If

    If IsArray(fieldValue) Then
        reports = fieldValue
        directReportCount = UBound(reports) - LBound(reports) + 1
        For i = LBound(reports) To UBound(reports)
            currentValue = SanitizeField(CStr(reports(i)))
            If Len(result) > 0 Then result = result & "}{"
            result = result & currentValue
        Next i
    Else
        directReportCount = 1
        result = SanitizeField(CStr(fieldValue))
    End If

    BuildDirectReportsString = result
End Function

Private Function SafeFieldValue(ByVal rs As Object, ByVal fieldName As String) As String
    Dim rawValue As Variant

    rawValue = rs.Fields(fieldName).Value
    If IsNull(rawValue) Then
        SafeFieldValue = vbNullString
    Else
        SafeFieldValue = SanitizeField(CStr(rawValue))
    End If
End Function

Private Function SanitizeField(ByVal value As String) As String
    Dim cleaned As String
    cleaned = value
    cleaned = Replace(cleaned, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, "|", "/")
    SanitizeField = cleaned
End Function

