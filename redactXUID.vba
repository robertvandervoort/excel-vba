Function redactXUID(inputURI As String) As String
    Dim segments() As String
    Dim result As String
    Dim i As Long

    ' Split the URI into segments
    segments = Split(inputURI, "/")

    ' Handle each segment
    For i = 0 To UBound(segments)
        Dim segment As String
        segment = segments(i)

        ' Check for GUID or UID within the segment
        If IsGUID(segment) Then
            segment = "GUID-REDACTED"
        ElseIf IsCustomUID(segment) Then
            segment = "UID-REDACTED"
        End If

        ' Append the (potentially modified) segment
        result = result & segment & "/"
    Next i

    ' Remove trailing slash (if any)
    If Right(result, 1) = "/" Then
        result = Left(result, Len(result) - 1)
    End If

    ' Handle the case of "/" directly
    If result = "" Then
        result = "/"
    End If

    redactXUID = result ' Assign the result directly to the function name
End Function

' Helper Function to check for basic GUID format
Function IsGUID(strToCheck As String) As Boolean
    Dim i As Long
    
    If Len(strToCheck) <> 36 Then Exit Function ' Wrong length
    
    For i = 1 To Len(strToCheck)
       Select Case Mid(strToCheck, i, 1)
           Case "0" To "9", "a" To "f", "A" To "F", "-"
               ' Valid GUID characters
           Case Else
               IsGUID = False
               Exit Function
        End Select
    Next i
    
    IsGUID = True
End Function

' Helper Function: Custom UID check
Function IsCustomUID(strToCheck As String) As Boolean
    Dim i As Long

    If Len(strToCheck) <> 36 Then Exit Function ' Wrong length

    For i = 1 To Len(strToCheck)
       Select Case Mid(strToCheck, i, 1)
           Case "0" To "9", "a" To "z", "A" To "Z", "-"
               ' Valid custom UID characters
           Case Else
               IsCustomUID = False
               Exit Function
        End Select
    Next i
    
    IsCustomUID = True
End Function


