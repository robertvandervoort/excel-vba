Attribute VB_Name = "Module1"
Function StripURISegments(ByVal uri As String, ByVal segmentsToKeep As Integer) As String

    ' Early error handling (combined into one check)
    If segmentsToKeep <= 0 Or uri = "" Or uri = "/" Then
        StripURISegments = IIf(segmentsToKeep <= 0, "Invalid segmentsToKeep value", uri)
        Exit Function ' Explicitly exit to prevent "0" return
    End If

    Dim urlParts() As String
    urlParts = Split(uri, "/")

    ' Check for sufficient segments (simplified using UBound)
    If UBound(urlParts) < segmentsToKeep Then
        StripURISegments = uri
        Exit Function ' Explicitly exit to prevent "0" return
    End If

    ' Redimensioning (efficient slicing)
    If segmentsToKeep > 0 Then
        ReDim Preserve urlParts(segmentsToKeep - 1)
    End If
    StripURISegments = Join(urlParts, "/")

    ' Trailing slash (directly added to the result)
    If Right(uri, 1) = "/" Then StripURISegments = StripURISegments & "/"
  
End Function

