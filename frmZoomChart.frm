VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.1#0"; "mschart.ocx"
Begin VB.Form frmZoomChart 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Zoom Chart"
   ClientHeight    =   4170
   ClientLeft      =   8745
   ClientTop       =   3210
   ClientWidth     =   5055
   Icon            =   "frmZoomChart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5055
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4140
      Left            =   15
      OleObjectBlob   =   "frmZoomChart.frx":000C
      TabIndex        =   0
      Top             =   15
      Width           =   4980
   End
End
Attribute VB_Name = "frmZoomChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim iX As Integer

    On Error GoTo Errorhandler
    
    'setup form position based on mouse pointer location on frmMainChart
    Me.Top = gsngXpos
    Me.Left = gsngYpos
    
    'setup chart
    With MSChart1
        'chart setup
        .chartType = VtChChartType2dBar
        .AllowDithering = False
        .AllowDynamicRotation = False
        .AllowSelections = True
        .AllowSeriesSelection = False
        .AutoIncrement = False
        .ShowLegend = True
    End With
    
    'display zoom data
    Call GetChartData
   
Errorhandler:

End Sub

Private Sub Form_Resize()
    'Use manual scale to display y axis (value axis) based on size of form.
    'The larger the form the more major and minor divisions display for more detail.
    'see frmMainChart.Form_Resize notes for details
With MSChart1
    'simple chart resize.
    .Top = 0
    .Left = 0
    .Height = Me.Height - 570
    .Width = Me.Width
    
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Round(Me.Height / 235, 0)
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision / 19
End With

End Sub

Private Sub Form_Terminate()
    Set frmZoomChart = Nothing
End Sub

Sub GetChartData()
    Dim iX As Integer
    Dim iZ As Integer
    Dim sngMin As Single
    Dim sngMax As Single
    Dim sValue As String
    
    On Error GoTo Errorhandler:
    
    'self adjusting Y axis based on passed data points minimum and maximum values
    'set minimum value high to start
    sngMin = 1000000
    
    'get each data point value and determine if minimum or maximum value needs adjusting
    'this routine will handle all data series from Main Chart
    For iX = 0 To 4
        'find number or series by using ubound on 2nd dimension
        For iZ = 1 To UBound(gsngZoomData, 2)
            If gsngZoomData(iX, iZ) > sngMax Then sngMax = gsngZoomData(iX, iZ)
            If gsngZoomData(iX, iZ) < sngMin Then sngMin = gsngZoomData(iX, iZ)
        Next
    Next
    
    '+ 2% to sngMax and - 2% to sngMin so chart max
    'and min values extend above or below actual max/min values.
    'This is extra fluff but makes the chart a bit more readable
    sngMax = sngMax + (sngMax * 0.02)
    sngMin = sngMin - (sngMin * 0.02)
    
    With MSChart1
        ' Display chart
        ' Use manual scale to display y axis (value axis)
        .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
        .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
        .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 19
        .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = sngMin
        .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = sngMax
        
        'load chart with array
        .ChartData = gsngZoomData
        
        'add series label(s)
        For iX = 1 To MSChart1.Plot.SeriesCollection.Count
            .Column = iX
            .ColumnLabel = gsSeriesLabels(iX)
        Next
    End With
    
    'additional chart series settings.
    Dim serX As series
    For Each serX In MSChart1.Plot.SeriesCollection
        serX.SeriesMarker.Show = False
        serX.SeriesMarker.Auto = False
        serX.DataPoints.Item(-1).Marker.Size = 140
        serX.DataPoints.Item(-1).Marker.Style = VtMarkerStyle3dBall
        serX.SeriesMarker.Show = True
    Next
    
Errorhandler:
End Sub

