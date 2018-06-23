Attribute VB_Name = "DrawRect"
Public Const SCF = 14.25

Sub Draw_Rect(ByVal text As String, ByVal name As String, ByVal row As Integer, ByVal col As Integer, ByVal width As Integer, ByVal height As Integer)
Attribute Draw_Rect.VB_ProcData.VB_Invoke_Func = " \n14"


    ActiveSheet.Shapes.AddShape(msoShapeRectangle, col * SCF, row * SCF, width * SCF, height * SCF).Select
    
    'Fill and Border
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)
    Selection.ShapeRange.Line.Weight = 2

    'Name und text
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = text
    Selection.name = name
    
    'Font Format
    Selection.ShapeRange(1).TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font.Size = 12
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .NameComplexScript = "Arial"
        .NameFarEast = "Arial"
        .name = "Arial"
    End With
    
    
End Sub

Sub main()
    Dim colcounter, rowcounter As Integer
    Dim name As String
    
    rowcounter = 2
    colcounter = 0
    For i = 1 To Sheets("List").Cells(Rows.Count, 1).End(xlUp).row
        Select Case (colcounter)
            Case Is = 10:
                rowcounter = rowcounter + 6
                colcounter = 0
        End Select
        
        name = Sheets("List").Cells(i, 1).Value
        Call Draw_Rect(name, name, rowcounter, 2 + colcounter * 10, 8, 4)
        colcounter = colcounter + 1
    Next
    
End Sub



