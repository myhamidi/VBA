Attribute VB_Name = "main"
Sub main()
    Dim colcounter, rowcounter As Integer
    Dim name As String
    
    rowcounter = 2
    colcounter = 0
    For i = 2 To Sheets("List").Cells(Rows.Count, 1).End(xlUp).row
        Select Case (colcounter)
            Case Is = 10:
                rowcounter = rowcounter + 6
                colcounter = 0
        End Select

        name = Sheets("List").Cells(i, 1).Value
        Call Draw_Rect(name, name, rowcounter, 2 + colcounter * 10, 8, 4)
        colcounter = colcounter + 1
    Next
    
    For i = 2 To Sheets("List").Cells(Rows.Count, 3).End(xlUp).row
        
        name = Sheets("List").Cells(i, 3).Value
        conStart = Sheets("List").Cells(i, 4).Value
        conEnd = Sheets("List").Cells(i, 5).Value
        Call Draw_Connector(name, conStart, conEnd)
    Next
    
End Sub
