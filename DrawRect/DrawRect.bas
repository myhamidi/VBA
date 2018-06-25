Attribute VB_Name = "DrawRect"
Public Const SCF = 14.25

Sub Draw_Rect(ByVal text As String, ByVal name As String, _
ByVal row As Integer, ByVal col As Integer, ByVal width As Integer, ByVal height As Integer)


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





