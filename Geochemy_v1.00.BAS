Sub Main
Debug.Clear
Dim Arr(21) As String, Field(21) As String, zCol(21) As Integer, Srf As Variant
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
If Plot = "map.txt" Then
	Plot.Close 2
	Set Plot = SurferApp.ActiveDocument
End If
Set Shapes = Plot.Shapes
SurferApp.ScreenUpdating = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''GridOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set xSize = 25  				  ' X Step grid, m
		Set ySize = 25		   			 ' Y Step grid, m
		Set SearchRad = 5000		 ' Radius Search
		Set ScalingRatio = 100		  ' Map scaling ratio 1/N
		Set NumPasses = 1 			' Number Passes Filter
		Set axisVisible = 0				' On or Off Axes
		Set ShowColorScale = 0	   ' On or Off Scale
		Set OnFilter = 0				  ' On or Off Grid Filter
		Set OnBlank = 0				   ' On or Off Blanking (Need map.bln file in SourcePath)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''LayersOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set StepTop = 15    		   ' Step page units between maps
		Set StepLeft = 25   		   ' Step page units between maps
'''''''''''''''''''''''''''''''''''''''''''''''''''''''SourceOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set SourcePath = "d:\Grids\"
		Set SourceFile = "d:\Grids\map.txt"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''GridsNames'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Field(1) = "Ag"              : Arr(1) = "Ag"       ' Names columns Z in SourceFile (В Arr(i) вписывать название столбцов из таблицы)
		Field(2) = "As"        		: Arr(2) = "As"
		Field(3) = "Au"        		: Arr(3) = "Au"
		Field(4) = "Bi"              : Arr(4) = "Bi"
		Field(5) = "Cd"              : Arr(5) = "Cd"
		Field(6) = "Cr"        		: Arr(6) = "Cr"
		Field(7) = "Cu"       		 : Arr(7) = "Cu"
		Field(8) = "Fe"              : Arr(8) = "Fe"
		Field(9) = "Hg"               : Arr(9) = "Hg"
		Field(10) = "Mn"        	: Arr(10) = "Mn"
		Field(11) = "Mo"       		 : Arr(11) = "Mo"
		Field(12) = "Ni"              : Arr(12) = "Ni"
		Field(13) = "Pb"               : Arr(13) = "Pb"
		Field(14) = "Pt"        		: Arr(14) = "Pt"
		Field(15) = "Rb"       		 : Arr(15) = "Rb"
		Field(16) = "Sn"              : Arr(16) = "Sn"
		Field(17) = "Sr"               : Arr(17) = "Sr"
		Field(18) = "Ti"        		: Arr(18) = "Ti"
		Field(19) = "V"       			 : Arr(19) = "V"
		Field(20) = "Zn"              : Arr(20) = "Zn"
		Field(21) = "Zr"               : Arr(21) = "Zr"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set Wks = SurferApp.Documents.Open(SourceFile)
CRows = Wks.Columns(1).Count
CCols = Wks.Rows(1).Count
Set WksHandle = Wks.Cells(Row:=1,Col:=1,LastRow:=1,LastCol:=CCols)
i=1
On Error GoTo FileError
For Each vCell In WksHandle
	If vCell.Value = "X" Then
		xCol = vCell.Column
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = "Y" Then
	   	yCol = vCell.Column
	   	Exit For
	End If
Next
For i = 1 To 21
	For Each vCell In WksHandle
		If vCell.Value = Arr(i) Then
			zCol(i) = vCell.Column
			Exit For
		End If
	Next
Next
Plot.Activate
i=1
'''''''''''''''''''''''''''''''' Create Grid ''''''''''''''''''''''''''''''''''''''''
		While i < 22
			If zCol(i)=0 Then
				GoTo n
			End If
			Zmin = Wks.Columns(zCol(i)).Statistics.Minimum
			Zmax = Wks.Columns(zCol(i)).Statistics.Maximum
			SurferApp.GridData6(SourceFile, ZLimitMin:=Zmin, ZLimitMax:=Zmax, ZLimitMinMethod:=2, ZLimitMaxMethod:=2, SearchRad1:=SearchRad, SearchRad2:=SearchRad, BlankOutsideHull:= 1, xCol:=xCol, yCol:= yCol, zCol:=zCol(i), xSize:= xSize, ySize:= ySize, Algorithm:=srfKriging, OutGrid:= SourcePath & Field(i) & H & ".grd")
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
			ElseIf i >14 And i <22 Then
				MapFrame.Top=-50+(i-14)*StepTop
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
			Kill(SourcePath & Field(i) & H & ".grd")
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
