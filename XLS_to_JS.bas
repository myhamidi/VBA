Attribute VB_Name = "XLS_to_JS"
Option Explicit

Sub export_in_json_format()
    'DIM
    Dim fs As Object
    Dim jsonfile
    Dim rangetoexport As Range
    Dim rowcounter As Long
    Dim columncounter As Long
    Dim linedata, tempString, sheetname As String
    
    'SET
    Set rangetoexport = Selection
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set jsonfile = fs.CreateTextFile(Application.ActiveWorkbook.Path & "\data.js", True)
    sheetname = ActiveSheet.Name
    
    'Set JSON Object Name equal to Sheet Name
    linedata = "{" & Chr(10) & """" & sheetname & """: ["
    
    '->WRITE TO FILE
    jsonfile.WriteLine linedata
    
    For rowcounter = 2 To rangetoexport.Rows.Count
        linedata = ""
        For columncounter = 1 To rangetoexport.Columns.Count
        'Handling of #REF Errors
        If Not rangetoexport.Cells(rowcounter, columncounter).Formula = "=#REF!" Then
            'Handling of #NV Errors
            If Not rangetoexport.Cells(rowcounter, columncounter).Text = "#NV" Then
                'Handling of Special Character: " Quotation Mark: Needs to be done before the JSON QUotationmarls are set
                tempString = Replace(rangetoexport.Cells(rowcounter, columncounter), """", "\""")
            Else
                tempString = "#NV"
            End If
        Else
            tempString = "#REF"
        End If
        
            linedata = linedata & """" & rangetoexport.Cells(1, columncounter) & """" & ":" & """" & tempString & """" & ","
        Next
        linedata = Left(linedata, Len(linedata) - 1)
        If rowcounter = rangetoexport.Rows.Count Then
            linedata = "{" & linedata & "}"
        Else
            linedata = "{" & linedata & "},"
        End If
        
        'Handling of Special Character: New Line
        linedata = Replace(linedata, Chr(10), "</br>")
        
        '->WRITE TO FILE
        jsonfile.WriteLine linedata
    Next
    linedata = "]}"
    
    '->WRITE TO FILE
    jsonfile.WriteLine linedata
    
    jsonfile.Close
    
    Set fs = Nothing
      
End Sub
