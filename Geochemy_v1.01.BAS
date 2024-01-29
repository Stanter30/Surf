Sub Main
Debug.Clear
Dim Arr(64) As String, Field(64) As String, zCol(64) As Integer, Srf As Variant
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
If Plot = "map.txt" Then
	Plot.Close 2
	Set Plot = SurferApp.ActiveDocument
End If
Set Shapes = Plot.Shapes
SurferApp.ScreenUpdating = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''GridOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set xSize = 100				  ' X Step grid, m
		Set ySize = 100	   			      ' Y Step grid, m
		Set SearchRad = 1000	 ' Radius Search
		Set ScalingRatio = 100		  ' Map scaling ratio 1/N
		Set NumPasses = 0 			' Number Passes Filter
		Set axisVisible = 0				' On or Off Axes
		Set ShowColorScale = 0	   ' On or Off Scale
		Set OnFilter = 0				  ' On or Off Grid Filter

		Set BlankOutsideHull = 0   ' On or Off Blank Outside Hull
		Set InflateHull = 0            ' When set to zero, the boundary connects the outside data points exactly. When set to a positive value, the area assigned the NoData value is moved outside the convex hull boundary by the number of map units specified. When set to a negative value, the area assigned the NoData value is moved inside the convex hull boundary by the number of map units specified.
		Set BlankBln = 0				     ' On or Off Blank (Need map.bln file in SourcePath)

		Set Alg = srfKriging	' srfInverseDistance, srfKriging, srfMinCurvature, srfShepards, srfNaturalNeighbor, srfNearestNeighbor, srfRadialBasis, srfTriangulation, srfMovingAverage, srfLocalPolynomial
		Set OnSplineSmooth = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''LayersOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set StepTop = 0    		   ' Step page units between maps
		Set StepLeft = 0   		   ' Step page units between maps

'''''''''''''''''''''''''''''''''''''''''''''''''''''''SourceOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set SourcePath = "D:\Grids\"
		Set SourceFile = "D:\Grids\map.txt"
		Set X_col = "X" ' Name X coord column
		Set Y_col = "Y" ' Name Y coord column

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''GridsNames'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Field(1) = "Ag"              : Arr(1) = "Ag" ' Names columns Z in SourceFile (В Arr(i) вписывать название столбцов из таблицы)
		Field(2) = "Al"              : Arr(2) = "Al"
		Field(3) = "As"        		: Arr(3) = "As"
		Field(4) = "Au"        		: Arr(4) = "Au"
		Field(5) = "Ba"              : Arr(5) = "Ba"
		Field(6) = "Bi"              : Arr(6) = "Bi"
		Field(7) = "Ca"        		: Arr(7) = "Ca"
		Field(8) = "Cd"       		 : Arr(8) = "Cd"
		Field(9) = "Ce"              : Arr(9) = "Ce"
		Field(10) = "Co"               : Arr(10) = "Co"
		Field(11) = "Cr"        	: Arr(11) = "Cr"
		Field(12) = "Cs"       		 : Arr(12) = "Cs"
		Field(13) = "Cu"              : Arr(13) = "Cu"
		Field(14) = "Dy"               : Arr(14) = "Dy"
		Field(15) = "Er"        		: Arr(15) = "Er"
		Field(16) = "Eu"       		 : Arr(16) = "Eu"
		Field(17) = "Fe"              : Arr(17) = "Fe"
		Field(18) = "Ga"               : Arr(18) = "Ga"
		Field(19) = "Gd"        		: Arr(19) = "Gd"
		Field(20) = "Hg"       			 : Arr(20) = "Hg"
		Field(21) = "In"              : Arr(21) = "In"
		Field(22) = "K"               : Arr(22) = "K"
		Field(23) = "La"              : Arr(23) = "La"
		Field(24) = "Li"        		: Arr(24) = "Li"
		Field(25) = "Mg"        		: Arr(25) = "Mg"
		Field(26) = "Mn"              : Arr(26) = "Mn"
		Field(27) = "Nb"              : Arr(27) = "Nb"
		Field(28) = "Nd"        		: Arr(28) = "Nd"
		Field(29) = "Ni"       		 : Arr(29) = "Ni"
		Field(30) = "P"              : Arr(30) = "P"
		Field(31) = "Pb"               : Arr(31) = "Pb"
		Field(32) = "Pd"        	: Arr(32) = "Pd"
		Field(33) = "Pr"       		 : Arr(33) = "Pr"
		Field(34) = "Pt"              : Arr(34) = "Pt"
		Field(35) = "Rb"               : Arr(35) = "Rb"
		Field(36) = "Sb"        		: Arr(36) = "Sb"
		Field(37) = "Sc"       		 : Arr(37) = "Sc"
		Field(38) = "Sm"              : Arr(38) = "Sm"
		Field(39) = "Sn"               : Arr(39) = "Sn"
		Field(40) = "Sr"        		: Arr(40) = "Sr"
		Field(41) = "Ta"       			 : Arr(41) = "Ta"
		Field(42) = "Tb"              : Arr(42) = "Tb"
		Field(43) = "Te"               : Arr(43) = "Te"
		Field(44) = "Th"              : Arr(44) = "Th"
		Field(45) = "Ti"        		: Arr(45) = "Ti"
		Field(46) = "Tl"        		: Arr(46) = "Tl"
		Field(47) = "U"              : Arr(47) = "U"
		Field(48) = "V"              : Arr(48) = "V"
		Field(49) = "W"        		: Arr(49) = "W"
		Field(50) = "Y"       		 : Arr(50) = "Y"
		Field(51) = "Yb"              : Arr(51) = "Yb"
		Field(52) = "Zn"               : Arr(52) = "Zn"
		Field(53) = "Zr"        	: Arr(53) = "Zr"
		Field(54) = "Mo"        	: Arr(54) = "Mo"
		Field(55) = "Be"        	: Arr(55) = "Be"
		Field(56) = "Ge"        	: Arr(56) = "Ge"
		Field(57) = "Hf"        	: Arr(57) = "Hf"
		Field(58) = "Na"        	: Arr(58) = "Na"
		Field(59) = "Re"        	: Arr(59) = "Re"
		Field(60) = "S"        	: Arr(60) = "S"
		Field(61) = "Se"        	: Arr(61) = "Se"
		Field(62) = "Ho"        	: Arr(62) = "Ho"
		Field(63) = "Lu"        	: Arr(63) = "Lu"
		Field(64) = "Tm"        	: Arr(64) = "Tm"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set Wks = SurferApp.Documents.Open(SourceFile)
CRows = Wks.Columns(1).Count
CCols = Wks.Rows(1).Count
Set WksHandle = Wks.Cells(Row:=1,Col:=1,LastRow:=1,LastCol:=CCols)
i=1
On Error GoTo FileError
For Each vCell In WksHandle
	If vCell.Value = X_col Then
		xCol = vCell.Column
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = Y_col Then
	   	yCol = vCell.Column
	   	Exit For
	End If
Next
For i = 1 To 64
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
		While i < 65
			If zCol(i)=0 Then
				GoTo n
			End If
			Zmin = Wks.Columns(zCol(i)).Statistics.Minimum
			Zmax = Wks.Columns(zCol(i)).Statistics.Maximum
			SurferApp.GridData6(SourceFile, ZLimitMin:=Zmin, ZLimitMax:=Zmax, ZLimitMinMethod:=2, ZLimitMaxMethod:=2, SearchRad1:=SearchRad, SearchRad2:=SearchRad, BlankOutsideHull:= BlankOutsideHull, InflateHull:= InflateHull, xCol:=xCol, yCol:= yCol, zCol:=zCol(i), xSize:= xSize, ySize:= ySize, Algorithm:=Alg, OutGrid:= SourcePath & Field(i) & H & ".grd")
			If OnFilter = 1 Then 
				SurferApp.GridFilter2(InGrid:= SourcePath & Field(i) & ".grd", Filter:=srfFilterGaussian, OutGrid:= SourcePath & Field(i) & ".grd", EdgeOp:=3, BlankOp:=2, NumPasses:=NumPasses)
			End If
			If BlankBln = 1 Then
				BLN_file = SourcePath & "map.bln"
				SurferApp.GridBlank(SourcePath & Field(i) & ".grd", BLN_file, SourcePath & Field(i) & ".grd")
			End If
			If OnSplineSmooth = 1 Then
                SurferApp.GridSplineSmooth2(InGrid:=SourcePath & Field(i) & ".grd", nRow:=15, nCol:=15, method:=srfSplineInsert, OutGrid:= SourcePath & Field(i) & H & ".grd")
			End If
'''''''''''''''''''''''''''' Create ContourMaps '''''''''''''''''''''''''''''''''''''
			'Set MapFrame = Shapes.AddContourMap(GridFileName:=SourcePath & Field(i) & H & ".grd")
			'Set Axes = MapFrame.Axes
			'Set ContourLayer = MapFrame.Overlays(1)
			'ContourLayer.FillContours = True
			'ContourLayer.FillForegroundColorMap.LoadPreset ("Rainbow3")
			'If i > 7 And i < 15 Then
				'MapFrame.Top=-50+(i-7)*StepTop
				'MapFrame.Left=-20+Int((i-0.01)/7)*StepLeft
			'ElseIf i >14 And i <22 Then
				'MapFrame.Top=-50+(i-14)*StepTop
				'MapFrame.Left=-20+Int((i-0.01)/7)*StepLeft
			'Else
				'MapFrame.Top=-50+i*StepTop
				'MapFrame.Left=-20
			'End If
			'MapFrame.ScalingRatio = ScalingRatio
				'For Each axis In Axes
					'axis.Visible = axisVisible
				'Next
			'ContourLayer.ShowColorScale = ShowColorScale
			'ContourLayer.ShowMajorLabels = 0
			'Debug.Print Field(i) & H & ".grd"
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
