Function StripGUIDorUIDfromURI(inputURI As String) As String
    ' Custom Excel function to strip GUIDs or UIDs from the end of a URI
    Dim lastSlashPos As Long

    ' Handle the case of "/" directly
    If inputURI = "/" Then
        Exit Function ' Skip unnecessary processing
    End If

    ' Check for specific prefix - this covers non GUID / UID formatted uses
    'If Left(inputURI, 24) = "/api/chat/getattachment/" Then
    '    StripGUIDorUIDorSegment = "/api/chat/getattachment"
    '    Exit Function
    'End If
    
    ' Check for GUID
    lastSlashPos = InStrRev(inputURI, "/")
    If IsGUID(Mid(inputURI, lastSlashPos + 1)) Then
        StripGUIDorUIDorSegment = Left(inputURI, lastSlashPos - 1)
        Exit Function
    End If

    ' Check for custom UID
    If IsCustomUID(Mid(inputURI, lastSlashPos + 1)) Then
        StripGUIDorUIDorSegment = Left(inputURI, lastSlashPos - 1)
        Exit Function
    End If

    StripGUIDorUIDorSegment = inputURI
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

' Helper Function to check for UID
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