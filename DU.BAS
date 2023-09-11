Sub Main
Debug.Clear
Dim Arr(14) As String, Field(14) As String, zCol(14) As Integer, Srf As Variant
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
If Plot = "map.txt" Then
	Plot.Close 2
	Set Plot = SurferApp.ActiveDocument
End If
Set Shapes = Plot.Shapes
SurferApp.ScreenUpdating = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''GridOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set xSize = 12.5  				 ' X Step grid, m
		Set ySize = 25		   			 ' Y Step grid, m
		Set SearchRad = 150		 ' Radius Search
		Set ScalingRatio = 200		  ' Map scaling ratio 1/N
		Set NumPasses = 1 			' Number Passes Filter
		Set axisVisible = 0				' On or Off Axes
		Set ShowColorScale = 0	   ' On or Off Scale
		Set OnFilter = 1				  ' On or Off Grid Filter
		Set OnBlank = 0				   ' On or Off Blanking (Need map.bln file in SourcePath)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''LayersOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set StepTop = 5   		   ' Step page units between maps
		Set StepLeft = 15   		   ' Step page units between maps
'''''''''''''''''''''''''''''''''''''''''''''''''''''''SourceOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set SourcePath = "d:\Grids\"
		Set SourceFile = "d:\Grids\map0.txt"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''GridsNames'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Field(1) = "dU.1"              : Arr(1) = "1_dU.1"       ' Names columns Z in SourceFile (В Arr(i) вписывать название столбцов из таблицы)
		Field(2) = "dU.5"        		: Arr(2) = "1_dU.5"
		Field(3) = "dU.10"        		: Arr(3) = "1_dU.10"
		Field(4) = "dU.15"              : Arr(4) = "1_dU.15"
		Field(5) = "dU.20"              : Arr(5) = "1_dU.20"
		Field(6) = "dU.25"        		: Arr(6) = "1_dU.25"
		Field(7) = "dU.30"       		 : Arr(7) = "1_dU.30"
		Field(8) = "dU.35"              : Arr(8) = "1_dU.35"
		Field(9) = "dU.40"               : Arr(9) = "1_dU.40"
		Field(10) = "dU.45"        	: Arr(10) = "1_dU.45"
		Field(11) = "dU.50"       		 : Arr(11) = "1_dU.50"
		Field(12) = "dU.55"              : Arr(12) = "1_dU.55"
		Field(13) = "dU.60"               : Arr(13) = "1_dU.60"
		Field(14) = "dU.63"        		: Arr(14) = "1_dU.63"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set Wks = SurferApp.Documents.Open(SourceFile)
CRows = Wks.Columns(1).Count
CCols = Wks.Rows(1).Count
Set WksHandle = Wks.Cells(Row:=1,Col:=1,LastRow:=1,LastCol:=CCols)
i=1
On Error GoTo FileError
For Each vCell In WksHandle
	If vCell.Value = "X" Or vCell.Value = "x" Then
		xCol = vCell.Column
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = "Y" Or vCell.Value = "y" Then
	   	yCol = vCell.Column
	   	Exit For
	End If
Next
For i = 1 To 14
	For Each vCell In WksHandle
		If vCell.Value = Arr(i) Then
			zCol(i) = vCell.Column
			Exit For
		End If
	Next
Next
'Plot.Activate
i=1
'''''''''''''''''''''''''''''''' Create Grid ''''''''''''''''''''''''''''''''''''''''
		While i < 15
			If zCol(i)=0 Then
				GoTo n
			End If
			Zmin = Wks.Columns(zCol(i)).Statistics.Minimum
			Zmax = Wks.Columns(zCol(i)).Statistics.Maximum
			SurferApp.GridData6(SourceFile, ZLimitMin:=Zmin, ZLimitMax:=Zmax, ZLimitMinMethod:=2, ZLimitMaxMethod:=2, SearchRad1:=SearchRad, SearchRad2:=SearchRad, BlankOutsideHull:= 0, xCol:=xCol, yCol:= yCol, zCol:=zCol(i), xSize:= xSize, ySize:= ySize, Algorithm:=srfKriging, OutGrid:= SourcePath & Field(i) & H & ".grd")
			If OnFilter = 1 Then 
				SurferApp.GridFilter2(InGrid:= SourcePath & Field(i) & H & ".grd", Filter:=srfFilterGaussian, OutGrid:= SourcePath & Field(i) & H & ".grd", EdgeOp:=3, BlankOp:=2, NumPasses:=NumPasses)
			End If
			If OnBlank = 1 Then
				BLN_file = SourcePath & "map.bln"
				SurferApp.GridBlank(SourcePath & Field(i) & H & ".grd", BLN_file, SourcePath & Field(i) & H & ".grd")
			End If
'''''''''''''''''''''''''''' Create ContourMaps '''''''''''''''''''''''''''''''''''''
			Set MapFrame = Shapes.AddContourMap(GridFileName:=SourcePath & Field(i) & H & ".grd")
			Set Axes = MapFrame.Axes
			Set ContourLayer = MapFrame.Overlays(1)
			ContourLayer.FillContours = True
			ContourLayer.FillForegroundColorMap.LoadPreset ("Rainbow3")
			If i > 7 And i < 15 Then
				MapFrame.Top=-50+(i-7)*StepTop
				MapFrame.Left=-20+Int((i-0.01)/7)*StepLeft
			Else
				MapFrame.Top=-50+i*StepTop
				MapFrame.Left=-20
			End If
			MapFrame.ScalingRatio = ScalingRatio
				For Each axis In Axes
					axis.Visible = axisVisible
				Next
			ContourLayer.ShowColorScale = ShowColorScale
			ContourLayer.ShowMajorLabels = 0
			Debug.Print Field(i) & H & ".grd"
			'Kill(SourcePath & Field(i) & H & ".grd")
n:
			i=i+1
		Wend
Wks.Close 2
SurferApp.ScreenUpdating = True
End
'Error instructions
	FileError:
	Debug.Print  "Not Found:	" + Arr(i) + "						" + Err.Description
	Resume Next
End Sub
