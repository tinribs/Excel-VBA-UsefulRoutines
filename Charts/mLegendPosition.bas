Attribute VB_Name = "mLegendPosition"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : mLegendPosition
' Author    : Kane
' Date      : 10/01/2019
' Purpose   : Routines for manipulating the legend on a XY Chart
'			  May work on other chart types, but not tested yet
'---------------------------------------------------------------------------------------

Public Enum LegendPosition
    [_First] = 0
    None = 0
    TopRight = 2 ^ 0
    TopLeft = 2 ^ 1
    BottomRight = 2 ^ 2
    BottomLeft = 2 ^ 3
    MiddleBottom = 2 ^ 4
    MiddleTop = 2 ^ 5
    MiddleRight = 2 ^ 6
    MiddleLeft = 2 ^ 7
    [_Last] = 128
End Enum


'---------------------------------------------------------------------------------------
' Procedure : GetLegendPosition
' Author    : Kane
' Date      : 25/12/2018
' Purpose   : Finds the legend position inside the PlotArea
' Returns:
'   Long = Position of the Legend
'           0 = early exit code (or no Legend)
'          -1 = outside PlotArea
'          -2 = unknown position
'          >0 = position based on enum LegendPosition
'---------------------------------------------------------------------------------------
Public Function GetLegendPosition(cht As Chart) As Long
    Dim LgdCentreX As Double
    Dim LgdCentreY As Double

    If cht Is Nothing Then Exit Function
    If Not cht.HasLegend Then Exit Function
    
    With cht
        LgdCentreX = .Legend.Left + .Legend.Width / 2
        LgdCentreY = .Legend.Top + .Legend.Height / 2
        If (LgdCentreX > .PlotArea.InsideLeft And LgdCentreX < (.PlotArea.InsideWidth + .PlotArea.InsideLeft)) And (LgdCentreY > .PlotArea.InsideTop And LgdCentreY < (.PlotArea.InsideHeight + .PlotArea.InsideTop)) Then
            'Legend is inside the PlotArea
            'split the width/height into 8 sections, and look for the left/right and top/bottom edge of the legend for its place in the section
            If (.Legend.Left < (.PlotArea.InsideLeft + .PlotArea.InsideWidth / 8)) And (.Legend.Top < (.PlotArea.InsideTop + .PlotArea.InsideHeight / 8)) Then
                GetLegendPosition = LegendPosition.TopLeft
            ElseIf (.Legend.Left + .Legend.Width) > (.PlotArea.InsideLeft + 7 * (.PlotArea.InsideWidth / 8)) And (.Legend.Top < (.PlotArea.InsideTop + .PlotArea.InsideHeight / 8)) Then
                GetLegendPosition = LegendPosition.TopRight
            ElseIf (.Legend.Left + .Legend.Width) > (.PlotArea.InsideLeft + 7 * (.PlotArea.InsideWidth / 8)) And (.Legend.Top + .Legend.Height) > (.PlotArea.InsideTop + 7 * (.PlotArea.InsideHeight / 8)) Then
                GetLegendPosition = LegendPosition.BottomRight
            ElseIf (.Legend.Left < (.PlotArea.InsideLeft + .PlotArea.InsideWidth / 8)) And (.Legend.Top + .Legend.Height) > (.PlotArea.InsideTop + 7 * (.PlotArea.InsideHeight / 8)) Then
                GetLegendPosition = LegendPosition.BottomLeft
            'in MiddleTop/MiddleBottom section (3 & 5)
            ElseIf (LgdCentreX < (.PlotArea.InsideLeft + 5 * (.PlotArea.InsideWidth / 8))) And _
                    (LgdCentreX > (.PlotArea.InsideLeft + 3 * (.PlotArea.InsideWidth / 8))) Then
                If (.Legend.Top < (.PlotArea.InsideTop + .PlotArea.InsideHeight / 8)) Then
                    GetLegendPosition = LegendPosition.MiddleTop
                ElseIf (.Legend.Top + .Legend.Height) > (.PlotArea.InsideTop + 7 * (.PlotArea.InsideHeight / 8)) Then
                    GetLegendPosition = LegendPosition.MiddleBottom
                End If
            'in MiddleLeft/MiddleRight section (3 & 5)
            ElseIf (LgdCentreY < (.PlotArea.InsideTop + 5 * (.PlotArea.InsideHeight / 8))) And _
                    (LgdCentreY > (.PlotArea.InsideTop + 3 * (.PlotArea.InsideHeight / 8))) Then
                If (.Legend.Left < (.PlotArea.InsideLeft + .PlotArea.InsideWidth / 8)) Then
                    GetLegendPosition = LegendPosition.MiddleLeft
                ElseIf (.Legend.Left + .Legend.Width) > (.PlotArea.InsideLeft + 7 * (.PlotArea.InsideWidth / 8)) Then
                    GetLegendPosition = LegendPosition.MiddleRight
                End If
            Else
                GetLegendPosition = -2  'unknown position
            End If
        Else
            GetLegendPosition = -1  'outside PlotArea
        End If
    End With
End Function



'---------------------------------------------------------------------------------------
' Procedure : PositionLegend
' Author    : Kane
' Date      : 23/01/2016
' Purpose   : Places the legend in one of the corners of the PlotArea, if possible
'             Considers the line segments between two datapoints and if they intersect
'             the legend (with a defined extended margin). If they intersect (or a
'             datapoint is inside) the legend then the position is marked as being unsuited.
'---------------------------------------------------------------------------------------
Sub PositionLegend(cht As Chart)
    Const PROCNAME As String = "PositionLegend"
    Dim i As Long
    Dim j As Long
    Dim vValuesX As Variant
    Dim vValuesY As Variant
    Dim lFound As Long          'composite value for corners where data points present
    Dim dXScale As Double
    Dim dYScale As Double
    Dim dLgdMargin As Double    'outer margin for the legend
    Dim LgdHeight As Double
    Dim LgdWidth As Double
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
        
    If cht Is Nothing Then Exit Sub
    If Not cht.HasLegend Then Exit Sub
'    Debug.Print cht.ChartType
    
    With cht
        'calculate legend size based on axes values
        dXScale = (.PlotArea.InsideWidth) / (.Axes(xlCategory).MaximumScale - .Axes(xlCategory).MinimumScale)
        dYScale = (.PlotArea.InsideHeight) / (.Axes(xlValue).MaximumScale - .Axes(xlValue).MinimumScale)
        
        dLgdMargin = 3 'additional margin for around the legend box
        LgdHeight = .Legend.Height + 2 * dLgdMargin 'legend height
        LgdWidth = .Legend.Width + 2 * dLgdMargin   'legend width

        'Look for data points in the corners of the chart that would be covered by the legend
        lFound = LegendPosition.None
        For i = 1 To .SeriesCollection.Count
            vValuesY = .SeriesCollection(i).Values
            vValuesX = .SeriesCollection(i).XValues
            'check to see of line between two data points crosses the legend box area
            For j = 2 To UBound(vValuesY)
                x1 = (vValuesX(j - 1) - .Axes(xlCategory).MinimumScale) * dXScale
                y1 = (.Axes(xlValue).MaximumScale - vValuesY(j - 1)) * dYScale
                x2 = (vValuesX(j) - .Axes(xlCategory).MinimumScale) * dXScale
                y2 = (.Axes(xlValue).MaximumScale - vValuesY(j)) * dYScale
                
                'TopLeft
                If LineSegmentIntersectsRectangle(x1, y1, x2, y2, _
                                                0, 0, _
                                                LgdWidth, LgdHeight) Then
'                    Debug.Print "Series "; .SeriesCollection(i).Name; " for Point "; (j - 1) & "-" & j; " crosses legend in TL"
                    'mark lFound as having an intersection in TopLeft corner of the PlotArea
                    If (lFound And LegendPosition.TopLeft) = 0 Then lFound = lFound + LegendPosition.TopLeft 'mark lFound as having an intersection in TopLeft corner of the PlotArea
                End If
                'TopRight
                If LineSegmentIntersectsRectangle(x1, y1, x2, y2, _
                                                (.PlotArea.InsideWidth - LgdWidth), 0, _
                                                .PlotArea.InsideWidth, LgdHeight) Then
'                    Debug.Print "Series "; .SeriesCollection(i).Name; " for Point "; (j - 1) & "-" & j; " crosses legend in TR"
                    If (lFound And LegendPosition.TopRight) = 0 Then lFound = lFound + LegendPosition.TopRight
                End If
                'BottomRight
                If LineSegmentIntersectsRectangle(x1, y1, x2, y2, _
                                                (.PlotArea.InsideWidth - LgdWidth), (.PlotArea.InsideHeight - .PlotArea.InsideTop - LgdHeight), _
                                                .PlotArea.InsideWidth, (.PlotArea.InsideHeight - .PlotArea.InsideTop)) Then
'                    Debug.Print "Series "; .SeriesCollection(i).Name; " for Point "; (j - 1) & "-" & j; " crosses legend in BR"
                    If (lFound And LegendPosition.BottomRight) = 0 Then lFound = lFound + LegendPosition.BottomRight
                End If
                'BottomLeft
                 If LineSegmentIntersectsRectangle(x1, y1, x2, y2, _
                                                0, (.PlotArea.InsideHeight - LgdHeight), _
                                                LgdWidth, .PlotArea.InsideHeight) Then
'                    Debug.Print "Series "; .SeriesCollection(i).Name; " for Point "; (j - 1) & "-" & j; " crosses legend in BL"
                    If (lFound And LegendPosition.BottomLeft) = 0 Then lFound = lFound + LegendPosition.BottomLeft
                End If
                'MiddleBottom
                If LineSegmentIntersectsRectangle(x1, y1, x2, y2, _
                                                (.PlotArea.InsideWidth - LgdWidth) / 2, (.PlotArea.InsideHeight - LgdHeight), _
                                                (.PlotArea.InsideWidth + LgdWidth) / 2, (.PlotArea.InsideHeight)) Then
'                    Debug.Print "Series "; .SeriesCollection(i).Name; " for Point "; (j - 1) & "-" & j; " crosses legend in BL"
                    If (lFound And LegendPosition.MiddleBottom) = 0 Then lFound = lFound + LegendPosition.MiddleBottom
                End If
            Next
        Next
        
'        If lFound And LegendPosition.TopLeft Then Debug.Print "Data point in Top Left Quadrant."
'        If lFound And LegendPosition.TopRight Then Debug.Print "Data point in Top Right Quadrant."
'        If lFound And LegendPosition.BottomLeft Then Debug.Print "Data point in Bottom Left Quadrant."
'        If lFound And LegendPosition.BottomRight Then Debug.Print "Data point in Bottom Right Quadrant."
'        If lFound And LegendPosition.None Then Debug.Print "Data points in all Quadrants."
        
        'Position the legend
        If (lFound And LegendPosition.TopLeft) = 0 Then
            .Legend.Left = .PlotArea.InsideLeft + dLgdMargin
            .Legend.Top = .PlotArea.InsideTop + dLgdMargin
        ElseIf (lFound And LegendPosition.TopRight) = 0 Then
            .Legend.Left = .PlotArea.InsideLeft + .PlotArea.InsideWidth - .Legend.Width - dLgdMargin
            .Legend.Top = .PlotArea.InsideTop + dLgdMargin
        ElseIf (lFound And LegendPosition.BottomRight) = 0 Then
            .Legend.Left = .PlotArea.InsideLeft + .PlotArea.InsideWidth - .Legend.Width - dLgdMargin
            .Legend.Top = .PlotArea.InsideTop + .PlotArea.InsideHeight - .Legend.Height - dLgdMargin
        ElseIf (lFound And LegendPosition.BottomLeft) = 0 Then
            .Legend.Left = .PlotArea.InsideLeft + dLgdMargin
            .Legend.Top = .PlotArea.InsideTop + .PlotArea.InsideHeight - .Legend.Height - dLgdMargin
        ElseIf (lFound And LegendPosition.MiddleBottom) = 0 Then
            .Legend.Left = .PlotArea.InsideLeft + (.PlotArea.InsideWidth / 2) - (.Legend.Width + dLgdMargin) / 2
            .Legend.Top = .PlotArea.InsideTop + .PlotArea.InsideHeight - .Legend.Height - dLgdMargin
        Else
            Debug.Print "Data points in all possible Legend positions."
        End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : ResizeLegend
' Author    : Kane
' Date      : 02/03/2016
' Purpose   : Shrinks the legend size
'---------------------------------------------------------------------------------------
Sub ResizeLegend(cht As Chart)
    Const PROCNAME As String = "ResizeLegend"
    Dim i As Long
    Dim dLgdNewHeight As Double
    Dim dLgdEntryMargin As Double    'outer margin for the legend
    
'    Dim cht As Chart
'    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub
    If Not cht.HasLegend Then Exit Sub
    
    dLgdEntryMargin = 2
    With cht
        With .Legend
'            Debug.Print "H = " & .Height
            For i = 1 To .LegendEntries.Count
                dLgdNewHeight = dLgdNewHeight + .LegendEntries(i).Height + dLgdEntryMargin
                Debug.Print .LegendEntries(i).Height, .LegendEntries(i).Width, .LegendEntries(i).LegendKey.Width
            Next
            .Height = dLgdNewHeight
            .Width = .LegendEntries(1).Width + dLgdEntryMargin
            Debug.Print "H = " & .Height & "  x  W = " & .Width
        End With
        
    End With
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : LineSegmentIntersectsRectangle
' Author    : StackOverflow
' Reference : https://stackoverflow.com/questions/1585525/how-to-find-the-intersection-point-between-a-line-and-a-rectangle
' Date      : 02/03/2016
' Purpose   : Finds if a line (two data points) intersects a rectangle (the legend)
'			  (x1,y1) and (x2,y2) are start and end points of line segment
'			  (minX,minY) and (maxX,maxY) are corner points of the rectangle
'---------------------------------------------------------------------------------------
Public Function LineSegmentIntersectsRectangle(x1 As Variant, y1 As Variant, x2 As Variant, y2 As Variant, minX As Double, minY As Double, maxX As Double, maxY As Double) As Boolean
    Dim m As Double
    Dim y As Double, x As Double
    
    LineSegmentIntersectsRectangle = False
    '' Completely outside.
     If (((x1 <= minX And x2 <= minX) Or (y1 <= minY And y2 <= minY)) Or ((x1 >= maxX And x2 >= maxX) Or (y1 >= maxY And y2 >= maxY))) Then
        Exit Function
    End If
    
    '' Start or end inside.
    If ((x1 > minX And x1 < maxX And y1 > minY And y1 < maxY) Or (x2 > minX And x2 < maxX And y2 > minY And y2 < maxY)) Then
        LineSegmentIntersectsRectangle = True
        Exit Function
    End If

    m = (y2 - y1) / (x2 - x1)

    y = m * (minX - x1) + y1
    If (y > minY And y < maxY) Then
        LineSegmentIntersectsRectangle = True
        Exit Function
    End If

    y = m * (maxX - x1) + y1
    If (y > minY And y < maxY) Then
        LineSegmentIntersectsRectangle = True
        Exit Function
    End If

    x = (minY - y1) / m + x1
    If (x > minX And x < maxX) Then
        LineSegmentIntersectsRectangle = True
        Exit Function
    End If

    x = (maxY - y1) / m + x1
    If (x > minX And x < maxX) Then
        LineSegmentIntersectsRectangle = True
        Exit Function
    End If
    
End Function



