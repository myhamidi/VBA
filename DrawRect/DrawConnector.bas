Attribute VB_Name = "DrawConnector"
Sub Draw_Connector(ByVal name, ByVal conStart, ByVal conEnd)

    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 50).Select
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(conStart), 4
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(conEnd), 2
        
End Sub
