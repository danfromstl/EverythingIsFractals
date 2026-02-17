Option Explicit

Private Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const PRINTABLE_ASCII_START As Long = 32
Private Const PRINTABLE_ASCII_SPAN As Long = 95
Private Const MOD_24 As Long = 16777216
Private Const MOD_48 As LongLong = 281474976710655

Public Function SimpleHash(ByVal s As String) As Long
    Dim hash As Long
    Dim i As Long

    hash = 5381
    For i = 1 To Len(s)
        hash = ((hash * 33) Mod 2147483647) + Asc(Mid$(s, i, 1))
    Next i

    SimpleHash = hash
End Function

Public Function SimpleHashFormula(ByVal cell As Range) As LongLong
    Dim textValue As String
    Dim hash As LongLong
    Dim i As Long

    textValue = CellValueAsString(cell)
    hash = 5381

    For i = 1 To Len(textValue)
        hash = ((hash * 33) Mod 2147483647) + Asc(Mid$(textValue, i, 1))
    Next i

    SimpleHashFormula = hash
End Function

Public Function SimpleChecksum(ByVal cell As Range) As Long
    Dim textValue As String
    Dim checksum As Long
    Dim i As Long

    textValue = CellValueAsString(cell)
    checksum = 0

    For i = 1 To Len(textValue)
        checksum = checksum + Asc(Mid$(textValue, i, 1))
    Next i

    SimpleChecksum = checksum
End Function

Public Function SimpleChecksumHash(ByVal cell As Range) As String
    Dim textValue As String
    Dim i As Long
    Dim sum As Long

    textValue = CellValueAsString(cell)
    sum = 0

    For i = 1 To Len(textValue)
        sum = (sum * 33 + Asc(Mid$(textValue, i, 1))) Mod 32767
    Next i

    SimpleChecksumHash = Right$(textValue, 2) & Right$("000000" & CStr(sum), 6)
End Function

Public Function SimpleChecksumHash_v2(ByVal cell As Range) As String
    Dim textValue As String
    Dim i As Long
    Dim sum As Long
    Dim standardized As String

    textValue = CellValueAsString(cell)
    sum = 0

    For i = 1 To Len(textValue)
        sum = (sum * 33 + Asc(Mid$(textValue, i, 1))) Mod 65535
    Next i

    standardized = Right$("000000" & CStr(sum), 6)
    SimpleChecksumHash_v2 = Right$(textValue, 2) & standardized
End Function

Public Function Base64ChecksumHash(ByVal cell As Range) As String
    Dim textValue As String
    Dim i As Long
    Dim sum As Long
    Dim binarySum As String
    Dim base64 As String
    Dim sixBits As String
    Dim indexValue As Long

    textValue = CellValueAsString(cell)
    sum = 0

    For i = 1 To Len(textValue)
        sum = (sum * 33 + Asc(Mid$(textValue, i, 1))) Mod MOD_24
    Next i

    binarySum = Dec2Bin24(sum)
    base64 = vbNullString

    For i = 0 To 3
        sixBits = Mid$(binarySum, i * 6 + 1, 6)
        indexValue = BinToDec(sixBits)
        base64 = base64 & Mid$(BASE64_TABLE, indexValue + 1, 1)
    Next i

    Base64ChecksumHash = Right$(textValue, 2) & base64
End Function

Public Function Dec2Bin24(ByVal decNumber As Long) As String
    Dim i As Long
    Dim result As String

    result = vbNullString
    For i = 23 To 0 Step -1
        If (decNumber And (2 ^ i)) <> 0 Then
            result = result & "1"
        Else
            result = result & "0"
        End If
    Next i

    Dec2Bin24 = result
End Function

Public Function Base64_Hash_8(ByVal cell As Range) As String
    Dim textValue As String
    Dim i As Long
    Dim j As Long
    Dim sum As LongLong
    Dim part1 As LongLong
    Dim binSum As String
    Dim base64Chars As String
    Dim base64Str As String
    Dim base64Value As Long

    textValue = CellValueAsString(cell)
    sum = 0

    For i = 1 To Len(textValue)
        part1 = (sum * 33) Mod MOD_48
        sum = (part1 + Asc(Mid$(textValue, i, 1))) Mod MOD_48
    Next i

    binSum = Dec2Bin48(sum)
    base64Str = vbNullString

    For i = 1 To 48 Step 6
        base64Chars = Mid$(binSum, i, 6)
        base64Value = 0
        For j = 1 To 6
            base64Value = base64Value * 2 + CInt(Mid$(base64Chars, j, 1))
        Next j
        base64Str = base64Str & Mid$(BASE64_TABLE, base64Value + 1, 1)
    Next i

    Base64_Hash_8 = Right$(textValue, 2) & base64Str
End Function

Public Function Dec2Bin48(ByVal decNumber As LongLong) As String
    Dim i As Long
    Dim result As String

    result = vbNullString
    For i = 47 To 0 Step -1
        If (decNumber And (2 ^ i)) <> 0 Then
            result = result & "1"
        Else
            result = result & "0"
        End If
    Next i

    While Len(result) < 48
        result = "0" & result
    Wend

    Dec2Bin48 = result
End Function

Public Function OffsetEncode(ByVal cell As Range) As String
    Dim textValue As String
    Dim result As String
    Dim i As Long
    Dim bodyLength As Long

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0

    For i = 1 To bodyLength
        result = result & ShiftPrintable(Mid$(textValue, i, 1), 6)
    Next i

    OffsetEncode = result
End Function

Public Function OffsetEncode_v2(ByVal cell As Range) As String
    Dim textValue As String
    Dim result As String
    Dim i As Long
    Dim j As Long
    Dim offset As Long
    Dim bodyLength As Long
    Dim chunk As String
    Dim ch As String

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    offset = 0
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0

    For i = 1 To bodyLength Step 6
        chunk = Mid$(textValue, i, 6)
        For j = 1 To Len(chunk)
            ch = Mid$(chunk, j, 1)
            result = result & ShiftPrintable(ch, 6 + offset)
        Next j
        offset = offset + SumAscii(chunk)
    Next i

    OffsetEncode_v2 = Left$(result, 9)
End Function

Public Function SumAscii(ByVal s As String) As Long
    Dim i As Long
    Dim sum As Long

    sum = 0
    For i = 1 To Len(s)
        sum = sum + Asc(Mid$(s, i, 1))
    Next i

    SumAscii = sum
End Function

Public Function OffsetEncode_v3(ByVal cell As Range) As String
    Dim textValue As String
    Dim result As String
    Dim ch As String
    Dim i As Long
    Dim remainder As String
    Dim checksum As Long
    Dim bodyLength As Long

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0

    For i = 1 To Min(bodyLength, 6)
        ch = ShiftPrintable(Mid$(textValue, i, 1), 6)
        If ch = "|" Then ch = "3"
        result = result & ch
    Next i

    If bodyLength < 6 Then
        result = result & String$(6 - bodyLength, "a")
    End If

    If Len(textValue) > 8 Then
        remainder = Mid$(textValue, 7, Len(textValue) - 8)
        checksum = SumAscii(remainder)
        result = result & "x" & CStr(checksum)
    End If

    OffsetEncode_v3 = result
End Function

Public Function Min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Public Function OffsetEncode_v5(ByVal cell As Range) As String
    Dim textValue As String
    Dim result As String
    Dim ch As String
    Dim i As Long
    Dim remainderLength As Long
    Dim bodyLength As Long

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0

    For i = 1 To Min(bodyLength, 6)
        ch = ShiftPrintable(Mid$(textValue, i, 1), 6)
        If ch = "|" Then ch = "P"
        result = result & ch
    Next i

    If bodyLength < 6 Then
        result = result & String$(6 - bodyLength, "a")
    End If

    If Len(textValue) > 8 Then
        remainderLength = Len(Mid$(textValue, 7, Len(textValue) - 8))
        result = result & "x" & CStr(remainderLength)
    End If

    OffsetEncode_v5 = result
End Function

Public Function OffsetEncode_v6(ByVal cell As Range, ByVal offsetNum As Single) As String
    Dim textValue As String
    Dim result As String
    Dim i As Long
    Dim bodyLength As Long
    Dim shift As Long
    Dim ch As String

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0
    shift = CLng(offsetNum)

    For i = 1 To bodyLength
        ch = ShiftPrintable(Mid$(textValue, i, 1), shift)
        result = result & EscapeSpecialChar(ch)
    Next i

    If bodyLength < 6 Then
        result = result & String$(6 - bodyLength, "a")
    End If

    OffsetEncode_v6 = result
End Function

Public Function OffsetDecode_v6(ByVal encodedStrCell As Range, ByVal offsetNum As Single) As String
    Dim encodedText As String
    Dim suffix As String
    Dim body As String
    Dim decodedBody As String
    Dim i As Long
    Dim shift As Long
    Dim tokenChar As String

    encodedText = CellValueAsString(encodedStrCell)
    If Len(encodedText) < 3 Then
        OffsetDecode_v6 = encodedText
        Exit Function
    End If

    suffix = Left$(encodedText, 2)
    If Mid$(encodedText, 3, 1) = "_" Then
        body = Mid$(encodedText, 4)
    Else
        body = Mid$(encodedText, 3)
    End If

    decodedBody = vbNullString
    shift = CLng(offsetNum)
    i = 1
    Do While i <= Len(body)
        tokenChar = ReadEncodedToken(body, i)
        decodedBody = decodedBody & ShiftPrintable(tokenChar, -shift)
    Loop

    OffsetDecode_v6 = decodedBody & suffix
End Function

Public Function OffsetEncode_v7(ByVal cell As Range, ByVal offsetNum As Single) As String
    Dim textValue As String
    Dim result As String
    Dim i As Long
    Dim bodyLength As Long
    Dim baseShift As Long
    Dim appliedShift As Long
    Dim ch As String

    textValue = CellValueAsString(cell)
    result = Right$(textValue, 2) & "_"
    bodyLength = Len(textValue) - 2
    If bodyLength < 0 Then bodyLength = 0
    baseShift = CLng(offsetNum)

    For i = 1 To bodyLength
        appliedShift = baseShift
        If (i Mod 2) = 0 Then appliedShift = appliedShift + 1
        ch = ShiftPrintable(Mid$(textValue, i, 1), appliedShift)
        result = result & EscapeSpecialChar(ch)
    Next i

    OffsetEncode_v7 = result
End Function

Public Function OffsetDecode_v7(ByVal encodedStrCell As Range, ByVal offsetNum As Single) As String
    Dim encodedText As String
    Dim suffix As String
    Dim body As String
    Dim decodedBody As String
    Dim i As Long
    Dim bodyPos As Long
    Dim baseShift As Long
    Dim appliedShift As Long
    Dim tokenChar As String

    encodedText = CellValueAsString(encodedStrCell)
    If Len(encodedText) < 3 Then
        OffsetDecode_v7 = encodedText
        Exit Function
    End If

    suffix = Left$(encodedText, 2)
    If Mid$(encodedText, 3, 1) = "_" Then
        body = Mid$(encodedText, 4)
    Else
        body = Mid$(encodedText, 3)
    End If

    decodedBody = vbNullString
    baseShift = CLng(offsetNum)
    bodyPos = 0
    i = 1

    Do While i <= Len(body)
        tokenChar = ReadEncodedToken(body, i)
        bodyPos = bodyPos + 1
        appliedShift = baseShift
        If (bodyPos Mod 2) = 0 Then appliedShift = appliedShift + 1
        decodedBody = decodedBody & ShiftPrintable(tokenChar, -appliedShift)
    Loop

    OffsetDecode_v7 = decodedBody & suffix
End Function

Private Function CellValueAsString(ByVal cell As Range) As String
    On Error GoTo HandleError

    If cell Is Nothing Then
        CellValueAsString = vbNullString
    Else
        CellValueAsString = CStr(cell.Value2)
    End If
    Exit Function

HandleError:
    CellValueAsString = vbNullString
End Function

Private Function ShiftPrintable(ByVal ch As String, ByVal shift As Long) As String
    Dim asciiValue As Long
    Dim shifted As Long

    asciiValue = Asc(ch)
    shifted = ((asciiValue - PRINTABLE_ASCII_START + shift) Mod PRINTABLE_ASCII_SPAN + PRINTABLE_ASCII_SPAN) Mod PRINTABLE_ASCII_SPAN
    ShiftPrintable = Chr$(shifted + PRINTABLE_ASCII_START)
End Function

Private Function EscapeSpecialChar(ByVal ch As String) As String
    If ch = "|" Then
        EscapeSpecialChar = "[bar]"
    ElseIf ch = " " Then
        EscapeSpecialChar = "[sp]"
    Else
        EscapeSpecialChar = ch
    End If
End Function

Private Function ReadEncodedToken(ByVal encodedBody As String, ByRef idx As Long) As String
    If Mid$(encodedBody, idx, 5) = "[bar]" Then
        ReadEncodedToken = "|"
        idx = idx + 5
        Exit Function
    End If

    If Mid$(encodedBody, idx, 4) = "[sp]" Then
        ReadEncodedToken = " "
        idx = idx + 4
        Exit Function
    End If

    ReadEncodedToken = Mid$(encodedBody, idx, 1)
    idx = idx + 1
End Function

Private Function BinToDec(ByVal bits As String) As Long
    Dim i As Long
    Dim value As Long

    value = 0
    For i = 1 To Len(bits)
        value = value * 2 + CInt(Mid$(bits, i, 1))
    Next i

    BinToDec = value
End Function

