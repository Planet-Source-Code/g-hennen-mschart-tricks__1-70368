VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.1#0"; "mschart.ocx"
Begin VB.Form frmMainChart 
   Caption         =   "Main Chart..."
   ClientHeight    =   4680
   ClientLeft      =   6915
   ClientTop       =   2010
   ClientWidth     =   7695
   Icon            =   "fmrMainChart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7695
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4575
      Left            =   0
      OleObjectBlob   =   "fmrMainChart.frx":000C
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmMainChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '2 dimension array to hold chart data
    Dim vChartData(1 To 12, 2) As Variant
    Dim iColumn As Integer

With MSChart1
    'setup chart
    .chartType = VtChChartType2dLine
    .AllowDithering = False
    .AllowDynamicRotation = False
    .AllowSelections = True
    .AllowSeriesSelection = False
    .AutoIncrement = False
    .ShowLegend = True

    'set chart type
    .chartType = VtChChartType2dLine
    
    ' Use manual scale to display y axis otherwise the Y axis always starts at 0
    .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    
    'NOTE: Major and Minor division values are set in the Form resize event at startup and when user resizes form.
    
    'I manually set chart limits here since I know the minimum and maximum values in this example.
    'In the real world you may want to add a routine that determines Min and Max values of your actual data.
    'the frmZoomChart sub GetChartData has one approach to this.
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 1000
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 4200
    
    'fill array with data for charting
    'first load column labels. In this case month abbreviations.
    'If the first series of a multi-dimensional array contains strings, those strings will become the X axis labels of the chart
    For iColumn = 1 To 12
        vChartData(iColumn, 0) = MonthName(iColumn, True)
    Next
    
    'Series 1.  12 sample data points for each month
    vChartData(1, 1) = 1100
    vChartData(2, 1) = 2300
    vChartData(3, 1) = 3100
    vChartData(4, 1) = 1800
    vChartData(5, 1) = 2400
    vChartData(6, 1) = 3500
    vChartData(7, 1) = 3900
    vChartData(8, 1) = 4100
    vChartData(9, 1) = 3700
    vChartData(10, 1) = 2600
    vChartData(11, 1) = 1400
    vChartData(12, 1) = 1800
    
    'Series 2. 12 sample data points for each month
    vChartData(1, 2) = 1800
    vChartData(2, 2) = 1400
    vChartData(3, 2) = 2600
    vChartData(4, 2) = 3700
    vChartData(5, 2) = 4100
    vChartData(6, 2) = 3900
    vChartData(7, 2) = 3500
    vChartData(8, 2) = 2400
    vChartData(9, 2) = 1800
    vChartData(10, 2) = 3100
    vChartData(11, 2) = 2300
    vChartData(12, 2) = 1100
    
    'load chart with array
    MSChart1.ChartData = vChartData

    'add series1 label (current year)
    .Column = 1
    .ColumnLabel = Year(Date)
    'add series2 label (previous year)
    .Column = 2
    .ColumnLabel = Year(Date) - 1

    'add chart title
    .TitleText = "Order Units By Month"

    'this adds a 3d ball to each datapoint in the series.
    'This routine will handle multiple data series as needed
    Dim serX As series
    For Each serX In MSChart1.Plot.SeriesCollection
        serX.SeriesMarker.Show = False
        serX.SeriesMarker.Auto = False
        serX.DataPoints.Item(-1).Marker.Size = 180
        serX.DataPoints.Item(-1).Marker.Style = VtMarkerStyle3dBall
        serX.SeriesMarker.Show = True
    Next
End With

End Sub

Private Sub Form_Resize()
    'Use manual scale to display y axis (value axis) based on size of form.
    'The larger the form the more major and minor divisions display for more detail.

With MSChart1
    'simple chart resize.
    .Top = 0
    .Left = 0
    'to resize chart height subtract the difference between the form height (Me.Height) and chart height (MSChart1.Height)
    'at default or design time.  In this example, at design time the form height is 5220 and the chart height is 4575
    '5220 - 4575 = 645.  During resize, 645 is subtracted from the form height to keep the chart in porportion to the form size.
    .Height = Me.Height - 645
    'chart width is equal to the form width
    .Width = Me.Width
    
    'For the Chart MajorDivision value at startup or default form size, I divide the form default startup height
    '(or height at design time) by the number of major divisions I want, in this case 20.
    'In this example, the form height at design time is 5220 so I divide it by 20 to get the proper divisor (5220/20 = 261).
    'Now as the user resizes the form, more or less major divisions will be displayed on the chart.
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Round(Me.Height / 261, 0)
    
    'For minor divisions, I divide the MajorDivision value by 19 (which equals 20 major divisions at default form size)
    'to get 0 MinorDivisions (i.e. no minor divisions) at startup.
    'If the form is resized larger, then minor divisions will be displayed for more detail.
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision / 19
End With
End Sub

Private Sub MSChart1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'these public variables keep track of mouse pointer location over chart.
    'If user clicks on data point to activate zoom chart then they are used to position
    'frmZoomChart close to mousepointer
    gsngXpos = X
    gsngYpos = Y + Me.Left
    
    'display a tool tip when mouse pointer hits a datapoint
    Dim iPart As Integer, iSeries As Integer, iDatapoint As Integer, iIndex3 As Integer, iIndex4 As Integer
    Dim dblValue As Double, iNullflag As Integer
    
    On Error GoTo Errorhandler
    
    With MSChart1
        .TwipsToChartPart X, Y, iPart, iDatapoint, iSeries, iIndex3, iIndex4
    
        If iPart = VtChPartTypePoint Then
            .DataGrid.GetData iSeries, iDatapoint, dblValue, iNullflag
            
            .Column = iDatapoint
            
            .ToolTipText = Format(dblValue, "###,###") & "  click to zoom chart"
        Else
            .ToolTipText = ""
        End If
        
    End With
    
Errorhandler:
    
End Sub

Private Sub MSChart1_PointSelected(series As Integer, datapoint As Integer, MouseFlags As Integer, Cancel As Integer)
    'When user selects a data point on the chart this routine passes chart data to frmZoomChart
    'to display the data with more detail.  In this example, the zoom chart displays 5 data points
    'from each series displayed on the main chart.
    
    Dim iX As Integer
    Dim iX2 As Integer
    Dim iZ As Integer
    Dim iSeries As Integer
    
    'get 5 data points based on selected row
    With MSChart1
        'redim passing array to hold data for zoom chart
        ReDim gsngZoomData(4, MSChart1.Plot.SeriesCollection.Count) As Variant
        
        'redim passing array to hold data series labels
        ReDim gsSeriesLabels(MSChart1.Plot.SeriesCollection.Count) As String
        
        'the DataPoint integer passes the user selected data point to this routine.  In this example
        'I want to display 5 of the 12 data points in the zoom chart. If user selects datapoint #6
        'then this routine uses 2 data points before and 2 data points after the selected
        'data point. So data points 4,5,6,7 & 8 are passed to the zoom chart. If the user
        'selects a data point close the the begining or end of the series then we need
        'some logic to select the proper starting data point for the zoom chart.
        
        'based on user selected data point, adjust the integer DataPoint to point to
        'the first data point we want to display
        Select Case datapoint
            Case 1, 2
                'if user selects data point 1 or 2 then start at 1
                datapoint = 1
            Case 3 To 10
                'if user selects points 3 to 10
                datapoint = datapoint - 2
            Case 11, 12
                'if user selects data point 11 or 12 then start at point 8
                datapoint = 8
        End Select
        
        'load 5 data points to msngZoomData() to pass data to frmZoomChart
        For iX = datapoint To datapoint + 4
            'point to row
            .Row = iX
            
            'save row label
            gsngZoomData(iZ, 0) = .RowLabel
            
            'this loads the global array gsngZoomData() with series data
            'from the main chart. it will pass all data series
            For iX2 = 1 To MSChart1.Plot.SeriesCollection.Count
                'save row data value
                gsngZoomData(iZ, iX2) = .ChartData(iX, iX2)
                
                'save data series labels to passing string array
                .Column = iX2
                gsSeriesLabels(iX2) = .ColumnLabel
            Next
            
            'increment iz
            iZ = iZ + 1
        Next
    End With
                        
    frmZoomChart.Show

End Sub
