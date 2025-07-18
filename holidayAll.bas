Attribute VB_Name = "holidayAll"
Public Sub holidayAll()
    Dim yearNum As Long
    Dim st As Worksheet
    Dim yearAll As New Collection
    Dim tempDay As Variant
    Dim holi As New holiday
    Dim lineNum As Long
    
    Set st = ThisWorkbook.Sheets("holiday")
    yearNum = st.Cells(2, 2).Value
    Set yearAll = yearDayCollectionMake(yearNum)
    lineNum = 3
    For Each tempDay In yearAll
        If holi.holidayHantei1(CDate(tempDay)) Then
            st.Cells(lineNum, 1) = holi.holidayName
            st.Cells(lineNum, 2) = Format(CDate(tempDay), "yyyy”NmŒŽd“ú(aaa)")
            lineNum = lineNum + 1
        End If
    Next
End Sub

Private Function yearDayCollectionMake(yearNum As Long) As Collection
    Dim i As Long
    Dim j As Long
    Dim uniday As String
    Dim monthNum As Long
    Dim dayNum As Long
    Dim st As Worksheet
    Dim hizuk As Date
    Set yearDayCollectionMake = New Collection

    For monthNum = 1 To 12
        For dayNum = 1 To 31
            hizuke = DateSerial(yearNum, monthNum, dayNum)
            If Format(hizuke, "m") > monthNum Then Exit For
            uniday = Format(hizuke, "yyyy/mm/dd")
            yearDayCollectionMake.Add uniday
        Next dayNum
    Next monthNum
End Function
