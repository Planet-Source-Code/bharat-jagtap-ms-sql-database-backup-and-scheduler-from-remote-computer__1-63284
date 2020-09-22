Attribute VB_Name = "FX_MSChartBas"

Public Sub addDataArray(MyChart As Object, ByRef MyDataArray() As Variant, resetGraph As Boolean)
On Error Resume Next

Dim varChartType As Variant

'  Pass an array of new graph values to the chart object
  
  MyChart.ChartData = MyDataArray
 
  If resetGraph Then
    Call FX_ResetMSChart(MyChart)
  End If
 
End Sub


Public Sub FX_ResetMSChart(MyChart)

Dim numSeries As Integer
  
  With MyChart
  

    .chartType = VtChChartType2dBar
  
'   Establish the number of items in the group
    numSeries = .Plot.SeriesCollection.Count
            
'   Now add a black line to the border of each of the shapes
    For iCount = 1 To numSeries
'      .GraphObj.Plot.SeriesCollection(iCount).DataPoints(-1).Brush.FillColor.Set fillColours(iScheme, iCount, 1), fillColours(iScheme, iCount, 2), fillColours(iScheme, iCount, 3)
      .Plot.SeriesCollection(iCount).DataPoints(-1).EdgePen.VtColor.Set 0, 0, 0
    Next iCount
        
'   Turn off the background grids
    .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Axis(VtChAxisIdY2).AxisGrid.MajorPen.Style = VtPenStyleNull
    .Plot.Wall.Pen.Style = VtPenStyleNull
          
        
'   Setup the colours of the pens
          
'    For iCount = 1 To numSeries
'      .Plot.SeriesCollection(iCount).Pen.VtColor.Set penColours(iScheme, iCount, 1), penColours(iScheme, iCount, 2), penColours(iScheme, iCount, 3)
'    Next iCount
    
' Define the background colour to white
  
 
  .Backdrop.Fill.Brush.FillColor.Set 255, 255, 255
  .Backdrop.Fill.Style = VtFillStyleBrush
  

    
  End With

End Sub

Public Sub AddLegend(MyChart As Object, legendTgl As Boolean)
On Error Resume Next

' Turn on the Auto settings

MyChart.Plot.AutoLayout = True
  
  With MyChart.Legend

   If legendTgl Then
  
  '   Add the legend in the required position
      
      .Location.Visible = True
      .VtFont.Name = "Arial"
      .VtFont.Size = 8
      
  
      .Location.LocationType = VtChLocationTypeTop
      .VtFont.Effect = VtFontStyleBold
  Else
  
'   Turn the legend off

     .Location.Visible = False
  End If

End With
  

End Sub
Public Sub AddTitle(MyChart As Object, TitleVar As String, titleTgl As Boolean)
On Error Resume Next

' TitleID As TextBox, titleOn As Boolean, Optional strTitle As Variant, Optional titlePosition As Variant

' Toggle the title on and off and display the title in different quadrants on the screen
' If you wish to have the title displayed at the top but inside the chart,
' use the TitlePosition = topManual.  This will cause problems if you then start
' using the legend

  With MyChart
    
    If titleTgl Then
    
      .Title.Text = TitleVar
      .Title.VtFont.Name = "Arial"
      .Title.VtFont.Size = 12
    
      .Plot.AutoLayout = True
      .Title.Location.LocationType = VtChLocationTypeTop
     
    Else
      .Title.Text = ""

    End If
  End With
End Sub

Public Function getFreeSpaceOnDriveInMB(Optional ByVal strDriveLetter As String = "C") As String
On Error GoTo errHand
    Dim varMB As Variant
    Dim varCapacity As Variant
    Dim varUsed As Variant
    Dim varFree As Variant
    Dim fso As New FileSystemObject
    Dim oDrive As Drive
    
    varMB = CDec(1024) * CDec(1024) * CDec(1024)
    
 
    
    Set oDrive = fso.GetDrive(strDriveLetter)
    varFree = CDec(oDrive.FreeSpace) / (varMB)
    varCapacity = CDec(oDrive.TotalSize) / (varMB)
    varUsed = CDec(varCapacity) - CDec(varFree)
    
    getFreeSpaceOnDriveInMB = Format(CDec(oDrive.FreeSpace) / (varMB), "#############.00")

Exit Function
errHand:
    getFreeSpaceOnDriveInMB = "<ERROR>"
End Function
 
Public Function getUsedSpaceOnDriveInMB(Optional ByVal strDriveLetter As String = "C") As String
On Error GoTo errHand
    Dim varMB As Variant
    Dim varCapacity As Variant
    Dim varUsed As Variant
    Dim varFree As Variant
    
    varMB = CDec(1024) * CDec(1024) * CDec(1024)
    Set fso = New FileSystemObject
 
    
    Set oDrive = fso.GetDrive(strDriveLetter)
    varFree = CDec(oDrive.FreeSpace) / (varMB)
    varCapacity = CDec(oDrive.TotalSize) / (varMB)
    varUsed = CDec(varCapacity) - CDec(varFree)
    
    getUsedSpaceOnDriveInMB = Format(varUsed, "#############.00")

Exit Function
errHand:
    getUsedSpaceOnDriveInMB = "<ERROR>"
End Function
 

