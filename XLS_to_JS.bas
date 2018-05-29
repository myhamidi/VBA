Attribute VB_Name = "XLS_to_JS"
Option Explicit

Sub export_in_json_format()

    Dim fs As Object
    Dim jsonfile
    Dim rangetoexport As Range
    Dim rowcounter As Long
    Dim columncounter As Long
    Dim linedata As String
    
    Set rangetoexport = Selection
    
    Set fs = CreateObject("Scripting.FileSystemObject")
   
    Set jsonfile = fs.CreateTextFile(Application.ActiveWorkbook.Path & "\data.js", True)
    
    linedata = "{""Output"": ["
    jsonfile.WriteLine linedata
    For rowcounter = 2 To rangetoexport.Rows.Count
        linedata = ""
        For columncounter = 1 To rangetoexport.Columns.Count
            linedata = linedata & """" & rangetoexport.Cells(1, columncounter) & """" & ":" & """" & rangetoexport.Cells(rowcounter, columncounter) & """" & ","
        Next
        linedata = Left(linedata, Len(linedata) - 1)
        If rowcounter = rangetoexport.Rows.Count Then
            linedata = "{" & linedata & "}"
        Else
            linedata = "{" & linedata & "},"
        End If
        
        jsonfile.WriteLine linedata
    Next
    linedata = "]}"
    jsonfile.WriteLine linedata
    jsonfile.Close
    
    Set fs = Nothing
      
End Sub
