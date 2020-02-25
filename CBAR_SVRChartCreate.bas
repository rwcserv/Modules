Attribute VB_Name = "CBAR_SVRChartCreate"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

Function ChartCreate(ByRef cLeft As Single, ByRef cTop As Single, ByRef cSourceRng As Range, ByRef cWkst As Worksheet, _
                        Optional ByRef cTitle As String, _
                            Optional ByRef cLegendRng As Range, _
                                Optional ByRef cCategoryRng As Range, _
                                        Optional ByRef cLegPos, _
                                                Optional ByRef cType, _
                                                      Optional ByRef cSecType, _
                                                            Optional ByRef cPlot, _
                                                                    Optional ByRef cAxisTit As Boolean, _
                                                                            Optional ByRef cWidth, _
                                                                                    Optional ByRef cHeight)
                        
                        
                        
                        
    'cLeft                          'Left position for chart to be assigned
    'cTop                           'Top position for chart to be assigned
    'cSourceRng                     'range of data source
    'cWkst                          'worksheet to put the chart onto
    'cTitle                         'Title of the Chart
    'cLegendRng                     'RANGE Where the Legends are (typically the headers of the page) - RANGE
    'cCategoryRng                   'RANGE where the Categories are (Typically the Left part/row names) - RANGE
    'cLegPos                        'Position of the Legend (either xlTop, xlBottom, xlLeft, xlRight)
    'cType                          'Type of chart to be created (see Chart Types below)
    'cSecType                       'If a second Series, then what type of chart (see Chart Types below)
    'cPlot                          'How the Chart is organised (either xlColumns or xlRows)
    'cAxisTit                       'Do you want the Value Axis to have labels reflecting the legend? (True of False) - default is False
    'cWidth (in cm)                 'Optional Width - will default to  13 cm
    'cHeight (in cm)                'Optional Height - will default to  8 cm
    
    '*************************************************************************************************
    '******************GENERAL CHART INFO TO HELP BUILD THE CHART YOU WANT'***************************
    '*************************************************************************************************
    '***************************CHART TYPES***********************************************************
    'xl3DArea = 3-D Area                            xl3DAreaStacked = 3-D Stacked Area              xl3DAreaStacked100 = 3-D Stacked Area
    'xl3DBarClustered = 3-D Clustered Bar           xl3DBarStacked = 3-D Stacked Bar                xl3DBarStacked100 = 3-D 100% Stacked Bar
    'xl3DColumn = 3-D Column                        xl3DColumnClustered = 3-D Clustered Column      xl3DColumnStacked = 3-D Stacked Column
    'xl3DColumnStacked100 = 3-D 100% Stacked Column  xl3DLine = 3-D Line                            xl3DPie = 3-D Pie
    'xl3DPieExploded = Exploded 3-D Pie             xlArea = Area                                   xlAreaStacked  = Stacked Area
    'xlAreaStacked100 = 100% Stacked Area           xlBarClustered = Clustered Bar                  xlBarOfPie = Bar of Pie
    'xlBarStacked = Stacked Bar                     xlBarStacked100 = 100% Stacked Bar              xlBubble = Bubble
    'xlBubble3DEffect = Bubble with 3-D Effects     xlColumnClustered = Clustered Column            xlColumnStacked = Stacked Column
    'xlColumnStacked100 = 100% Stacked Column       xlConeBarClustered = Clustered Cone Bar         xlConeBarStacked = Stacked Cone Bar
    'xlConeBarStacked100 = 100% Stacked Cone Bar    xlConeCol = 3-D Cone Column                     xlConeColClustered = Clustered Cone Column
    'xlConeColStacked = Stacked Cone Column         xlConeColStacked100 = 100% Stacked Cone Column  xlCylinderBarStacked = Stacked Cylinder Bar
    'xlDoughnut = Doughnut                          xlDoughnutExploded = Exploded Doughnut          xlLineMarkers  = Line with Data Markers
    'xlLine = Line                                  xlLineStacked = Stacked Line                    xlPie = Pie
    'xlPieExploded = Exploded Pie                   xlPieOfPie = Pie of Pie                         xlPyramidBarClustered = Clustered Pyramid Bar
    'xlRadar = Radar                                xlRadarFilled = Filled Radar                    xlRadarMarkers = Radar with Data Markers
    'xlStockHLC = High-Low-Close                    xlStockOHLC = Open-High-Low-Close               xlStockVHLC = Volume-High-Low-Close
    'xlStockVOHLC = Volume-Open-High-Low-Close      xlSurface = 3-D Surface                         xlSurfaceTopView = Surface (Top View)
    'xlSurfaceWireframe = 3-D Surface(wire-frame)   xlXYScatter = Scatter                           xlXYScatterLines = Scatter with Lines
    'xlPyramidCol = 3-D Pyramid Column              xlPyramidColStacked = Stacked Pyramid Column    xlPyramidColClustered = Clustered Pyramid Column
    'xlCylinderCol = 3-D Cylinder Column                        xlCylinderColStacked = Stacked Cylinder Column
    'xlCylinderBarClustered = Clustered Cylinder Bar            xlCylinderBarStacked100 = 100% Stacked Cylinder Bar
    'xlCylinderColClustered = Clustered Cylinder Column         xlCylinderColStacked100 = 100% Stacked Cylinder Column
    'xlLineMarkersStacked100 = 100% Stacked Line w Markers      xlLineStacked100 = 100% Stacked Line
    'xlSurfaceTopViewWireframe = Surface (Top View wire-frame)  xlLineMarkersStacked = Stacked Line w Data Markers
    'xlPyramidBarStacked = Stacked Pyramid Bar                  xlPyramidBarStacked100 = 100% Stacked Pyramid Bar
    'xlPyramidColStacked100 = 100% Stacked Pyramid Column       xlXYScatterSmoothNoMarkers = Scatter with Smoothed Lines and No Data Markers
    'xlXYScatterSmooth = Scatter with SmoothedLines             xlXYScatterLinesNoMarkers = Scatter with Lines and No Data Markers
    '************************************************************************************************************************
    
    Dim cht As ChartObject, numGen As Long, NumSeries As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    
    If IsMissing(cWidth) Then
        cWidth = 13
    ElseIf cWidth = 0 Then
        cWidth = 13
    End If
    
    If IsMissing(cHeight) Then
        cHeight = 8
    ElseIf cHeight = 0 Then
        cHeight = 8
    End If
    ' INFO 2.54 cm/Inch and 72 Points/Inch or 28.34645669 points per cm
    cWidth = Round(cWidth * 28.34645669, 0)
    cHeight = Round(cHeight * 28.34645669, 0)
    
    'On Error GoTo Err_Routine
    
    With cWkst
    
        'Create a chart
        Set cht = cWkst.ChartObjects.Add(Left:=cLeft, Width:=cWidth, Top:=cTop, Height:=cHeight)
        
        'Give chart some data
        If IsMissing(cPlot) Then cPlot = xlColumns
        cht.Chart.SetSourceData Source:=cSourceRng, PlotBy:=cPlot
        
        'Determine the amount of Series from the datasource
        NumSeries = cht.Chart.SeriesCollection.Count
            
        'Apply Legend and Axis names and formatting
        If Not IsMissing(cLegendRng) Then
            With cht.Chart
                For numGen = 1 To NumSeries
                    If numGen = 1 And Not IsMissing(cType) Then
                        .SeriesCollection(numGen).Name = cLegendRng.Resize(1, 1).Offset(0, numGen - 1).Value
                        .HasAxis(xlCategory, xlPrimary) = True
                        .SeriesCollection(numGen).ChartType = cType
                        .HasAxis(xlValue, xlPrimary) = True
                        .Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
                        If cAxisTit = True Then
                        .Axes(xlValue, xlPrimary).HasTitle = True
                        .Axes(xlValue, xlPrimary).AxisTitle.text = cLegendRng.Resize(1, 1).Offset(0, numGen - 1).Value
                        End If
                    End If
                    If numGen = 2 And Not IsMissing(cSecType) Then
                        .SeriesCollection(numGen).ChartType = cSecType
                        .SeriesCollection(numGen).AxisGroup = xlSecondary
                        '.HasAxis(xlCategory, xlSecondary) = False
                        .HasAxis(xlValue, xlSecondary) = True
                        .Axes(xlCategory, xlSecondary).CategoryType = xlAutomatic
                        If cAxisTit = True Then
                        .Axes(xlValue, xlSecondary).HasTitle = True
                        .Axes(xlValue, xlSecondary).AxisTitle.text = cLegendRng.Resize(1, 1).Offset(0, numGen - 1).Value
                        End If
                    End If
                    If numGen > 1 And IsMissing(cSecType) Then
                        .SeriesCollection(numGen).Name = cLegendRng.Resize(1, 1).Offset(0, numGen - 1).Value
                        .HasAxis(xlCategory, xlPrimary) = True
                        .SeriesCollection(numGen).ChartType = cType
                        .HasAxis(xlValue, xlPrimary) = True
                        .Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
                        If cAxisTit = True Then
                        .Axes(xlValue, xlPrimary).HasTitle = True
                        .Axes(xlValue, xlPrimary).AxisTitle.text = cLegendRng.Resize(1, 1).Offset(0, numGen - 1).Value
                        End If
                    
                    End If
                Next
            End With
        End If
        
        'Apply Category names
        If Not cCategoryRng Is Nothing Then
    '        numCategory = cCategoryRng.Cells.Count
            cht.Chart.Axes(xlCategory, xlPrimary).CategoryNames = cCategoryRng
        End If
        
        'adding title
        If Not IsMissing(cTitle) Then
            cht.Chart.HasTitle = True
            cht.Chart.ChartTitle.text = cTitle
        End If
        
        
        'now position the legend
        If IsEmpty(cLegPos) Then cLegPos = xlRight
        cht.Chart.Legend.Position = cLegPos
            
        
    End With
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-ChartCreate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
