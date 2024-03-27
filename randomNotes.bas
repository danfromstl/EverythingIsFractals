' Testing random notes v1


Option Explicit

Public WithEvents myTextBox As MSForms.TextBox

'''' IMPORTANT: ''''
'MAKE SURE TO ENABLE Tools -> References -> Microsoft ActiveX Data Objects 2.8 Library

Sub ShowUserForm()
    Dim frm As New ExtractionProgressBox
    
    Set myTextBox = frm.TextBox1
    frm.Show vbModeless
End Sub

Sub WriteLineToUserForm(text As String)
    If Not myTextBox Is Nothing Then
        myTextBox.text = myTextBox.text & text & vbCrLf
        myTextBox.SelStart = Len(myTextBox.text)
    End If
End Sub



Sub TestUserForm()
    ShowUserForm
    
    Dim i As Long
    For i = 1 To 10
        WriteLineToUserForm "Processing item " & i
        Application.Wait (Now + TimeValue("0:00:01")) 'Pause for 1 second
        DoEvents
    Next i
    
    WriteLineToUserForm "Finished processing items"
    WriteLineToUserForm "Feel free to close this window"
    
End Sub










Sub ExportADUsersToCSV()
    Const ADS_SCOPE_SUBTREE = 2
    Const CSV_FILENAME = "allUsers_allSubdoamins.csv"
    Const DRS_FILENAME = "ManagersAndDRs.csv"
    
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As ADODB.Recordset
    Dim file As Integer
    Dim DRfile As Integer
    Dim username As String
    Dim password As String
    Dim strPath As String
    Dim DRstrPath As String
    Dim strLine As String
    Dim directReport As Variant
    
    Dim recordCount As Long
    Dim startTime As Date
    Dim endTime As Date
    Dim elapsedTime As Double
    
    Dim full_DR_string As String
    Dim DR_num As Long
    
    'Set targeted subdomains
    Dim subdomains(2) As String 'UPDATE ME IF MORE SDs ADDED!!!!!!!!!!!!!!
    subdomains(0) = "canada.root.corp/DC=canada,DC=root,DC=corp"
    subdomains(1) = "root.corp/DC=root,DC=corp"
    subdomains(2) = "DC=accounts,DC=root,DC=corp"
    
    
    'start the timer
    
    'Debug.Print "Process started at " & Format(Now, "hh:mm:ss AM/PM")
    WriteLineToUserForm "Process started at " & Format(Now, "hh:mm:ss AM/PM")
    
    recordCount = 0
    
    ' Prompt for credentials
    'username = InputBox("Please enter your username:", "Username")
    'password = InputBox("Please enter your password:", "Password")
    
    ' Connect to Active Directory
    conn.Provider = "ADsDSOObject"
    'conn.Properties("User ID") = username
    'conn.Properties("Password") = password
    'conn.Properties("Encrypt Password") = True
    conn.Open "Active Directory Provider"
    
    Set cmd.ActiveConnection = conn
    cmd.Properties("Page Size") = 1000
    cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    
    ' Create CSV file
    strPath = ThisWorkbook.Path & Application.PathSeparator & CSV_FILENAME
    DRstrPath = ThisWorkbook.Path & Application.PathSeparator & DRS_FILENAME
    
    'blank main output file
    file = FreeFile
    Open strPath For Output As #file
    Close #file
    
    'blank secondary manager file
    DRfile = FreeFile
    Open DRstrPath For Output As #DRfile
    Close #DRfile
    
    'reopen files for append
    file = FreeFile
    Open strPath For Append As #file
    DRfile = FreeFile
    Open DRstrPath For Append As #DRfile
    
    ' Write header row
    WriteLine file, "distinguishedName|manager|displayName|title|company|department|mail|sAMAccountName|msExchHideFromAddressLists|DRs|directReports|subdomain"
    WriteLine DRfile, "distinguishedName|DR_count|directReports"
    
    ShowUserForm
    
    startTime = Now
    'Debug.Print "Connection established. Starting loop at " & Format(startTime, "hh:mm:ss AM/PM")
    WriteLineToUserForm "Connection established. Starting loop at " & Format(startTime, "hh:mm:ss AM/PM")
    DoEvents
    
    Dim subdomain As Variant
    For Each subdomain In subdomains
        'Debug.Print "Querying subdomain: " & subdomain
        WriteLineToUserForm "Querying subdomain: " & subdomain
        
        cmd.CommandText = "SELECT distinguishedName, manager, displayName, title, company, department, mail, sAMAccountName, msExchHideFromAddressLists, directReports FROM 'LDAP://" & subdomain & "' WHERE objectCategory='user'"
        
        Set rs = cmd.Execute
        
        
        ' Write data rows
        Do While Not rs.EOF
            full_DR_string = ""
            DR_num = 0
            
            
            If Not IsNull(rs.Fields("directReports")) Then
                full_DR_string = Join(rs.Fields("directReports").Value, "}{")
                DR_num = UBound(rs.Fields("directReports").Value) + 1
                WriteLine DRfile, rs.Fields("distinguishedName") & "|" & DR_num & "|{" & full_DR_string & "}"
                'For Each directReport In rs.Fields("directReports")
                '    Write #DRfile, directReport & "_AND_"
                'Next directReport
                'Debug.Print "DR DNs: " & Join(rs.Fields("directReports").Value, ";")
            End If
            
            WriteLine file, rs.Fields("distinguishedName") & "|" & _
                          rs.Fields("manager") & "|" & _
                          rs.Fields("displayName") & "|" & _
                          rs.Fields("title") & "|" & _
                          rs.Fields("company") & "|" & _
                          rs.Fields("department") & "|" & _
                          rs.Fields("mail") & "|" & _
                          rs.Fields("sAMAccountName") & "|" & _
                          rs.Fields("msExchHideFromAddressLists") & "|" & _
                          DR_num & "|{" & _
                          full_DR_string & "}|" & _
                          subdomain
            
            recordCount = recordCount + 1
            If recordCount Mod 5000 = 0 Then
                'Debug.Print "Processed " & recordCount & " records."
                WriteLineToUserForm "Processed " & recordCount & " records. " & _
                "Records per second: " & Format((recordCount / ((Now - startTime) * 24 * 60 * 60)), "0.00") & "/sec"
                DoEvents
            End If
            
            rs.MoveNext
            
        DoEvents
        Loop
    Next subdomain
    
    ' Close CSV file
    Close #file
    Close #DRfile
    
    ' Close AD connection
    rs.Close
    conn.Close
    
    endTime = Now
    elapsedTime = endTime - startTime
    'Debug.Print "Query finished at " & Format(endTime, "hh:mm:ss AM/PM")
    WriteLineToUserForm "Query finished at " & Format(endTime, "hh:mm:ss AM/PM")
    
    Dim totalRecords As Long
    totalRecords = recordCount
    
    Dim totalSeconds As Double
    totalSeconds = elapsedTime * 24 * 60 * 60 ' Convert elapsedTime to seconds
    
    Dim recordsPerSecond As Double
    recordsPerSecond = totalRecords / totalSeconds
    
    'Debug.Print "Total records: " & totalRecords
    'Debug.Print "Total seconds: " & totalSeconds
    'Debug.Print "Total minutes: " & totalSeconds / 60
    'Debug.Print "Records per second: " & recordsPerSecond & "/sec"
    'Debug.Print "Records per minute: " & recordsPerSecond * 60 & "/min"
    
    WriteLineToUserForm "Total records: " & Format(totalRecords, "0.00")
    WriteLineToUserForm "Total seconds: " & Format(totalSeconds, "0.00")
    WriteLineToUserForm "Total minutes: " & Format((totalSeconds / 60), "0.00")
    WriteLineToUserForm "Records per second: " & Format(recordsPerSecond, "0.00") & "/sec"
    WriteLineToUserForm "Records per minute: " & Format((recordsPerSecond * 60), "0.00") & "/min"
    
    
    WriteLineToUserForm "Total AD users exported: " & totalRecords
    WriteLineToUserForm "Feel free to close this window"
    
End Sub

Sub WriteLine(file As Integer, line As String)
    Print #file, line
End Sub










'ANOTHER ONE :-)



'Sub Export_AD_DataROAR()
'    Dim conn As New ADODB.Connection
'    Dim rs As New ADODB.Recordset
'    Dim file As Integer
'
'    Const AD_url As String = "LDAP://DC=accounts,DC=root,DC=corp"
'
'    conn.Provider = "ADsDSOObject"
'    conn.Open "Active Directory Provider"
'
'    rs.Open "SELECT distinguishedName, manager, displayName, title, company, department, mail, sAMAccountName, msExchHideFromAddressLists FROM 'LDAP://DC=accounts,DC=root,DC=corp' WHERE objectCategory='user'", conn, adOpenStatic, adLockOptimistic
'
'    file = FreeFile()
'    Open "C:\Users\c7r2mr\Downloads\rawExports\AD_UsersROAR.csv" For Output As file
'
'    ' Write header row
'    WriteLine file, "distinguishedName;manager;displayName;title;company;department;mail;sAMAccountName;msExchHideFromAddressLists"
'
'    ' Write data rows
'    Do While Not rs.EOF
'        WriteLine file, rs.Fields("distinguishedName") & ";" & _
'                      rs.Fields("manager") & ";" & _
'                      rs.Fields("displayName") & ";" & _
'                      rs.Fields("title") & ";" & _
'                      rs.Fields("company") & ";" & _
'                      rs.Fields("department") & ";" & _
'                      rs.Fields("mail") & ";" & _
'                      rs.Fields("sAMAccountName") & ";" & _
'                      rs.Fields("msExchHideFromAddressLists")
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    conn.Close
'
'    Close file
'End Sub







'Sub OLDGetManagerDNs()
'
'    Dim conn As New ADODB.Connection
'    Dim rs As New ADODB.Recordset
'    Dim cmd As New ADODB.Command
'    Dim pageSize As Long
'    Dim strSQL As String
'    Dim domainName As String
'    Dim ldapPath As String
'    Dim searchFilter As String
'    Dim attributes As String
'
'    pageSize = 1000
'
'    ' Replace with your domain name
'    domainName = "DC=accounts,DC=root,DC=corp"
'    ldapPath = "LDAP://" & domainName
'
'    strSQL = "SELECT distinguishedName, manager FROM 'LDAP://DC=accounts,DC=root,DC=corp' WHERE objectCategory='person' AND objectClass='user'"
'
'    ' Connect to Active Directory
'    conn.Provider = "ADsDSOObject"
'    conn.Open "Active Directory Provider"
'    Set cmd.ActiveConnection = conn
'    cmd.CommandText = strSQL
'    cmd.Properties("Page Size") = pageSize
'
'    ' Define the LDAP search filter and attributes to retrieve
'    'searchFilter = "(&(objectClass=user)(objectCategory=person))"
'    'attributes = "distinguishedName,manager"
'
'    ' Execute the LDAP query
'    'Set rs = conn.Execute("<" & ldapPath & ">;" & searchFilter & ";" & attributes & ";subtree")
'    Set rs = cmd.Execute
'
'    ' Process the results
'    Do Until rs.EOF
'        If Not IsNull(rs.Fields("manager")) Then
'            Debug.Print "User DN: " & rs.Fields("distinguishedName")
'            Debug.Print "Manager DN: " & rs.Fields("manager")
'            Debug.Print "-----"
'        End If
'        rs.MoveNext
'    Loop
'
'    ' Close the connections
'    rs.Close
'    conn.Close
'
'End Sub









'   .d8888b.  d8b                        888          888    888                   888      
'  d88P  Y88b Y8P                        888          888    888                   888      
'  Y88b.                                 888          888    888                   888      
'   "Y888b.   888 88888b.d88b.  88888b.  888  .d88b.  8888888888  8888b.  .d8888b  88888b.  
'      "Y88b. 888 888 "888 "88b 888 "88b 888 d8P  Y8b 888    888     "88b 88K      888 "88b 
'        "888 888 888  888  888 888  888 888 88888888 888    888 .d888888 "Y8888b. 888  888 
'  Y88b  d88P 888 888  888  888 888 d88P 888 Y8b.     888    888 888  888      X88 888  888 
'   "Y8888P"  888 888  888  888 88888P"  888  "Y8888  888    888 "Y888888  88888P' 888  888 
'                               888                                                         
'                               888                                                         
'                               888                                                         


Function SimpleHash(s As String) As Long
    Dim hash As Long
    hash = 5381
    For i = 1 To Len(s)
        hash = ((hash * 33) Mod 2147483647) + Asc(Mid(s, i, 1))
    Next i
    SimpleHash = hash
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Function SimpleHashFormula(cell As Range) As LongLong

    On Error GoTo ErrorHandler
    
    Dim hash As LongLong
    hash = 5381
    For i = 1 To Len(cell.Value)
        hash = ((hash * 33) Mod 2147483647) + Asc(Mid(cell.Value, i, 1))
        
        ' Check if hash has overflowed
        If hash > 32767 Or hash < -32768 Then
            MsgBox "Overflow error after processing character " & i & " ('" & Mid(cell.Value, i, 1) & "')"
            DJB2Hash = CVErr(xlErrValue)
            Exit Function
        End If
        
    Next i
    SimpleHashFormula = hash
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & vbCrLf & vbCrLf & Err.Description
    DJB2Hash = CVErr(xlErrValue)

End Function





'   .d8888b.  d8b                        888           .d8888b.  888                        888                                      
'  d88P  Y88b Y8P                        888          d88P  Y88b 888                        888                                      
'  Y88b.                                 888          888    888 888                        888                                      
'   "Y888b.   888 88888b.d88b.  88888b.  888  .d88b.  888        88888b.   .d88b.   .d8888b 888  888 .d8888b  888  888 88888b.d88b.  
'      "Y88b. 888 888 "888 "88b 888 "88b 888 d8P  Y8b 888        888 "88b d8P  Y8b d88P"    888 .88P 88K      888  888 888 "888 "88b 
'        "888 888 888  888  888 888  888 888 88888888 888    888 888  888 88888888 888      888888K  "Y8888b. 888  888 888  888  888 
'  Y88b  d88P 888 888  888  888 888 d88P 888 Y8b.     Y88b  d88P 888  888 Y8b.     Y88b.    888 "88b      X88 Y88b 888 888  888  888 
'   "Y8888P"  888 888  888  888 88888P"  888  "Y8888   "Y8888P"  888  888  "Y8888   "Y8888P 888  888  88888P'  "Y88888 888  888  888 
'                               888                                                                                                  
'                               888                                                                                                  
'                               888                                                                                                  



Function SimpleChecksum(cell As Range) As Long
    Dim checksum As Long
    Dim i As Integer
    For i = 1 To Len(cell.Value)
        checksum = checksum + Asc(Mid(cell.Value, i, 1))
    Next i
    SimpleChecksum = checksum
End Function








'   .d8888b.   .d8888b.           888    888                   888      
'  d88P  Y88b d88P  Y88b          888    888                   888      
'  888    888 Y88b.               888    888                   888      
'  888         "Y888b.            8888888888  8888b.  .d8888b  88888b.  
'  888            "Y88b.          888    888     "88b 88K      888 "88b 
'  888    888       "888          888    888 .d888888 "Y8888b. 888  888 
'  Y88b  d88P Y88b  d88P          888    888 888  888      X88 888  888 
'   "Y8888P"   "Y8888P"  88888888 888    888 "Y888888  88888P' 888  888 
'                                                                       
'                                                                       
'                                                                                                                                                 



Function SimpleChecksumHash(cell As Range) As String
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim sum As Long
    sum = 0
    For i = 1 To Len(cell.Value)
        sum = (sum * 33 + Asc(Mid(cell.Value, i, 1))) Mod 32767
    Next i

    ' Prepend the last two characters of the cell value to the result
    SimpleChecksumHash = Right(cell.Value, 2) & Right("000000" & CStr(sum), 6)

    Exit Function
    
ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & vbCrLf & vbCrLf & Err.Description
    SimpleChecksumHash = CVErr(xlErrValue)

End Function




Function SimpleChecksumHash_v2(cell As Range) As Long
    Dim i As Integer
    Dim sum As Long
    sum = 0
    For i = 1 To Len(cell.Value)
        sum = (sum * 33 + Asc(Mid(cell.Value, i, 1))) Mod 65535
    Next i

    ' Standardize to 6 characters, with leading '0's if needed
    SimpleChecksumHash_v2 = Right$("000000" & CStr(sum), 6)

    ' Prepend last two characters of original cell value
    SimpleChecksumHash_v2 = Right(cell.Value, 2) & SimpleChecksumHash_v2
End Function




'  888888b.                               .d8888b.      d8888            .d8888b.   .d8888b.           888    888                   888      
'  888  "88b                             d88P  Y88b    d8P888           d88P  Y88b d88P  Y88b          888    888                   888      
'  888  .88P                             888          d8P 888           888    888 Y88b.               888    888                   888      
'  8888888K.   8888b.  .d8888b   .d88b.  888d888b.   d8P  888           888         "Y888b.            8888888888  8888b.  .d8888b  88888b.  
'  888  "Y88b     "88b 88K      d8P  Y8b 888P "Y88b d88   888           888            "Y88b.          888    888     "88b 88K      888 "88b 
'  888    888 .d888888 "Y8888b. 88888888 888    888 8888888888          888    888       "888          888    888 .d888888 "Y8888b. 888  888 
'  888   d88P 888  888      X88 Y8b.     Y88b  d88P       888           Y88b  d88P Y88b  d88P          888    888 888  888      X88 888  888 
'  8888888P"  "Y888888  88888P'  "Y8888   "Y8888P"        888  88888888  "Y8888P"   "Y8888P"  88888888 888    888 "Y888888  88888P' 888  888 
'                                                                                                                                            
'                                                                                                                                            
'                                                                                                                                            


Function Base64ChecksumHash(cell As Range) As String
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim sum As Long
    Dim binarySum As String
    Dim base64 As String
    Dim base64Table As String
    base64Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    sum = 0
    For i = 1 To Len(cell.Value)
        sum = (sum * 33 + Asc(Mid(cell.Value, i, 1))) Mod 16777216 ' We mod by 16777216 (2^24) to ensure our sum fits into 24 bits.
    Next i

    ' Convert sum to a 24 bit binary string
    binarySum = Dec2Bin24(sum)
    
    ' Convert binary string to Base64
    For i = 0 To 3
        base64 = base64 & Mid(base64Table, WorksheetFunction.Bin2Dec(Mid(binarySum, i * 6 + 1, 6)) + 1, 1)
    Next i

    ' Prepend the last two characters of the cell value to the result
    Base64ChecksumHash = Right(cell.Value, 2) & base64

    Exit Function
    
ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & vbCrLf & vbCrLf & Err.Description
    Base64ChecksumHash = CVErr(xlErrValue)
End Function





'  8888888b.                     .d8888b.  888888b.   d8b           .d8888b.      d8888  
'  888  "Y88b                   d88P  Y88b 888  "88b  Y8P          d88P  Y88b    d8P888  
'  888    888                          888 888  .88P                      888   d8P 888  
'  888    888  .d88b.   .d8888b      .d88P 8888888K.  888 88888b.       .d88P  d8P  888  
'  888    888 d8P  Y8b d88P"     .od888P"  888  "Y88b 888 888 "88b  .od888P"  d88   888  
'  888    888 88888888 888      d88P"      888    888 888 888  888 d88P"      8888888888 
'  888  .d88P Y8b.     Y88b.    888"       888   d88P 888 888  888 888"             888  
'  8888888P"   "Y8888   "Y8888P 888888888  8888888P"  888 888  888 888888888        888  
'                                                                                        
'                                                                                        
'                                                                                        



Function Dec2Bin24(decNumber As Long) As String
    Dec2Bin24 = ""
    For i = 23 To 0 Step -1
        If decNumber And (2 ^ i) Then
            Dec2Bin24 = Dec2Bin24 & "1"
        Else
            Dec2Bin24 = Dec2Bin24 & "0"
        End If
    Next i
End Function





'  888888b.                               .d8888b.      d8888            .d8888b.   .d8888b.           888    888                   888                .d8888b.  
'  888  "88b                             d88P  Y88b    d8P888           d88P  Y88b d88P  Y88b          888    888                   888               d88P  Y88b 
'  888  .88P                             888          d8P 888           888    888 Y88b.               888    888                   888               Y88b. d88P 
'  8888888K.   8888b.  .d8888b   .d88b.  888d888b.   d8P  888           888         "Y888b.            8888888888  8888b.  .d8888b  88888b.            "Y88888"  
'  888  "Y88b     "88b 88K      d8P  Y8b 888P "Y88b d88   888           888            "Y88b.          888    888     "88b 88K      888 "88b          .d8P""Y8b. 
'  888    888 .d888888 "Y8888b. 88888888 888    888 8888888888          888    888       "888          888    888 .d888888 "Y8888b. 888  888          888    888 
'  888   d88P 888  888      X88 Y8b.     Y88b  d88P       888           Y88b  d88P Y88b  d88P          888    888 888  888      X88 888  888          Y88b  d88P 
'  8888888P"  "Y888888  88888P'  "Y8888   "Y8888P"        888  88888888  "Y8888P"   "Y8888P"  88888888 888    888 "Y888888  88888P' 888  888 88888888  "Y8888P"  
'                                                                                                                                                                
'                                                                                                                                                                
'                                                                                                                                                                


Function Base64_Hash_8(cell As Range) As String
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim sum As LongLong
    sum = 0
    Dim part1 As LongLong

    Dim doubleString as String: doubleString = cell.Value & cell.Value

    For i = 1 To Len(cell.Value)
        part1 = sum * 33
        part1 = part1 Mod (2 ^ 48 - 1)
        sum = (part1 + Asc(Mid(cell.Value, i, 1))) Mod (2 ^ 48 - 1)
    Next i

    ' Convert to binary
    Dim binSum As String
    binSum = Dec2Bin48(sum)

    ' Split into 6-bit chunks for Base64 encoding
    Dim base64Chars As String
    Dim base64Str As String
    base64Str = ""
    Dim base64Value As Integer
    For i = 1 To 48 Step 6
        base64Chars = Mid(binSum, i, 6)
        base64Value = 0
        For j = 1 To 6
            base64Value = base64Value * 2 + CInt(Mid(base64Chars, j, 1))
        Next j
        base64Str = base64Str & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", base64Value + 1, 1)
    Next i

    ' Prepend the last two characters of the cell value to the result
    Base64_Hash_8 = Right(cell.Value, 2) & base64Str

    Exit Function

ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & vbCrLf & vbCrLf & Err.Description
    Base64_Hash_8 = CVErr(xlErrValue)

End Function




Function Dec2Bin48(decNumber As LongLong) As String
    Dec2Bin48 = ""
    For i = 47 To 0 Step -1
        If decNumber And (2 ^ i) Then
            Dec2Bin48 = Dec2Bin48 & "1"
        Else
            Dec2Bin48 = Dec2Bin48 & "0"
        End If
    Next i

    ' Make sure the result is exactly 48 bits long by adding leading zeros if necessary
    While Len(Dec2Bin48) < 48
        Dec2Bin48 = "0" & Dec2Bin48
    Wend
End Function




Function OffsetEncode(cell As Range) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    
    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"
    
    ' Loop through the rest of the characters
    For i = 1 To Len(cell.Value) - 2
        char = Mid(cell.Value, i, 1)
        char = Chr(((Asc(char) - 32 + 6) Mod 95) + 32)
        result = result & char
    Next i
    
    OffsetEncode = result
End Function






Function OffsetEncode_v2(cell As Range) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    Dim offset As Integer
    Dim chunk As String

    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"

    ' Initialize the offset
    offset = 0

    ' Loop through the characters in chunks of 6
    For i = 1 To Len(cell.Value) - 2 Step 6
        chunk = Mid(cell.Value, i, 6)

        ' Loop through the characters in the current chunk and apply the offset
        For j = 1 To Len(chunk)
            char = Mid(chunk, j, 1)
            char = Chr(((Asc(char) - 32 + 6 + offset) Mod 95) + 32)
            result = result & char
        Next j

        ' Update the offset with the sum of the ASCII values of the current chunk
        offset = offset + SumAscii(chunk)
    Next i

    ' Limit the result to 9 characters
    OffsetEncode_v2 = Left(result, 9)
End Function

' Function to compute the sum of ASCII values of a string
Function SumAscii(s As String) As Integer
    Dim i As Integer
    Dim sum As Integer
    sum = 0
    For i = 1 To Len(s)
        sum = sum + Asc(Mid(s, i, 1))
    Next i
    SumAscii = sum
End Function







Function OffsetEncode_v3(cell As Range) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    Dim remainder As String
    Dim checksum As Long
    
    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"
    
    ' Loop through the first 6 characters and offset them
    For i = 1 To Min(Len(cell.Value) - 2, 6)
        char = Mid(cell.Value, i, 1)
        char = Chr(((Asc(char) - 32 + 6) Mod 95) + 32)
        If char = "|" Then char = "3" ' or any other character you prefer
        result = result & char
    Next i
    
    ' Pad with 'a's if needed
    If Len(cell.Value) - 2 < 6 Then
        result = result & String(6 - (Len(cell.Value) - 2), "a")
    End If

    ' Take the remaining characters and calculate a checksum
    If Len(cell.Value) > 8 Then
        remainder = Mid(cell.Value, 7, Len(cell.Value) - 8)
        For i = 1 To Len(remainder)
            checksum = checksum + Asc(Mid(remainder, i, 1))
        Next i
        result = result & "x" & checksum
    End If
    
    OffsetEncode_v3 = result
End Function

Function Min(a As Integer, b As Integer) As Integer
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function





Function OffsetEncode_v5(cell As Range) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    Dim remainder As String
    Dim checksum As Long
    
    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"
    
    ' Loop through the first 6 characters and offset them
    For i = 1 To Min(Len(cell.Value) - 2, 6)
        char = Mid(cell.Value, i, 1)
        char = Chr(((Asc(char) - 32 + 6) Mod 95) + 32)
        If char = "|" Then char = "P" ' or any other character you prefer
        result = result & char
    Next i
    
    ' Pad with 'a's if needed
    If Len(cell.Value) - 2 < 6 Then
        result = result & String(6 - (Len(cell.Value) - 2), "a")
    End If

    ' Take the remaining characters and calculate a checksum
    If Len(cell.Value) > 8 Then
        remainder = Len(Mid(cell.Value, 7, Len(cell.Value) - 8))
        result = result & "x" & remainder
    End If
    
    OffsetEncode_v5 = result
End Function






Function OffsetEncode_v6(cell As Range, offsetNum as Single) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    
    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"
    
    ' Loop through the rest of the characters
    For i = 1 To Len(cell.Value) - 2
        char = Mid(cell.Value, i, 1)
        char = Chr(((Asc(char) - 32 + offsetNum) Mod 95) + 32)
        If char = "|" Then char = "[bar]" ' or any other character you prefer
        If char = " " Then char = "[sp]" ' or any other character you prefer
        result = result & char
    Next i

    ' Pad with 'a's if needed
    If Len(cell.Value) - 2 < 6 Then
        result = result & String(6 - (Len(cell.Value) - 2), "a")
    End If
    
    OffsetEncode_v6 = result
End Function




Function OffsetDecode_v6(encodedStrCell As Range, offsetNum As Single) As String
    Dim encodedStrString AS String: encodedStrString = encodedStrCell.Value
    Dim result As String
    Dim char As String
    Dim i As Integer
    
    ' Start the result string with the last two characters of the encoded string
    result = Left(encodedStrString, 2)
    
    ' Remove the padding 'a's if they exist
    encodedStrString = Replace(encodedStrString, "a", "")

    ' Loop through the rest of the characters
    For i = 4 To Len(encodedStrString)
        char = Mid(encodedStrString, i, 1)

        ' Handle special cases
        If char = "[" Then
            If Mid(encodedStrString, i, 5) = "[bar]" Then
                char = "|"
                i = i + 4 ' Skip the special characters
            ElseIf Mid(encodedStrString, i, 4) = "[sp]" Then
                char = " "
                i = i + 3 ' Skip the special characters
            End If
        Else
            char = Chr(((Asc(char) - 32 - offsetNum + 95) Mod 95) + 32)
        End If

        result = result & char
    Next i

    OffsetDecode_v6 = result
End Function









'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------



Function OffsetEncode_v7(cell As Range, offsetNum as Single) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    
    ' Start the result string with the last two characters of the cell value
    result = Right(cell.Value, 2) & "_"
    
    ' Loop through the rest of the characters
    For i = 1 To Len(cell.Value) - 2
        char = Mid(cell.Value, i, 1)
        If (i Mod 2) = 0 Then
            char = Chr(((Asc(char) - 32 + offsetNum + 1) Mod 95) + 32)
        Else 
            char = Chr(((Asc(char) - 32 + offsetNum) Mod 95) + 32)
        End If
        If char = "|" Then char = "[bar]" ' or any other character you prefer
        If char = " " Then char = "[sp]" ' or any other character you prefer
        result = result & char
    Next i
    
    OffsetEncode_v7 = result
End Function





Function OffsetDecode_v7(encodedStrCell As Range, offsetNum As Single) As String
    Dim result As String
    Dim char As String
    Dim i As Integer
    
    ' Start the result string with the first part of the encoded string
    result = Left(encodedStrString, 3)
    
    ' Loop through the rest of the characters
    For i = 4 To Len(encodedStrString)
        char = Mid(encodedStrString, i, 1)
        If char = "[" Then
            If Mid(encodedStrString, i, 5) = "[bar]" Then
                char = "|"
                i = i + 4
            ElseIf Mid(encodedStrString, i, 4) = "[sp]" Then
                char = " "
                i = i + 3
            End If
        End If
        
        If (i Mod 2) = 0 Then
            char = Chr(((Asc(char) - 32 - offsetNum - 3) + 95) Mod 95 + 32)
        Else
            char = Chr(((Asc(char) - 32 - offsetNum) + 95) Mod 95 + 32)
        End If
        
        result = result & char
    Next i
    
    ' Reverse the original processing on the last two characters
    result = Left(result, Len(result) - 2) & Right(encodedStrString, 2)
    
    OffsetDecode_v7 = result
End Function



