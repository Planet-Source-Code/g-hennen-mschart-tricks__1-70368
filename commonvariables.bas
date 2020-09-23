Attribute VB_Name = "commonvariables"
Option Explicit

'public variables that keep track of mouse pointer location
'on the main chart.  If user clicks as data point
'they pass the mouse pointer location to the zoom chart
'so it displays close to the selected data point
Public gsngXpos As Single
Public gsngYpos As Single

'public array to pass data values and labels to frmZoomChart
Public gsngZoomData() As Variant

'public array to pass series labels to frmZoomChart
Public gsSeriesLabels() As String

