Sub Main
Debug.Clear
Dim Arr(4) As String, Field(4) As String, zCol(4) As Integer, Srf As Variant
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
If Plot = "map.txt" Then
	Plot.Close 2
	Set Plot = SurferApp.ActiveDocument
End If
Set Shapes = Plot.Shapes
SurferApp.ScreenUpdating = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''GridOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set xSize = 25					'X Step grid, m шаг сетки грида
		Set ySize = 25			 'Y Step grid, m
		Set SearchRad = 250		 ' Radius Search радиус поиска
		Set ScalingRatio = 150		  'Map scaling ratio 1/N масштаб карты
		Set NumPasses = 5 			' Number Passes Filter количество проходов фильтра сглаживания
		Set axisVisible = 0				' On or Off Axes
		Set ShowColorScale = 0	   ' On or Off Color Scale
		Set OnFilter = 1				  'On or Off Grid Filter Вкл или Выкл фильтра сглаживания
		Set OnBlank = 1				   'On or Off Blanking (Need map.bln file in SourcePath) Вкл или Выкл обрезки
		Set BlankOutsideHull = 0	'On or Off BlankOutsideHull режим обрезки
		Set InflateHull = 0             'Data inflated инфляция обрезки
		Set Alg = srfKriging	' srfInverseDistance, srfKriging, srfMinCurvature, srfShepards, srfNaturalNeighbor, srfNearestNeighbor, srfRadialBasis, srfTriangulation, srfMovingAverage, srfLocalPolynomial

'''''''''''''''''''''''''''''''''''''''''''''''''''''''LayersOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set Built = 1					  ' 0 - Built in depth (Flow), 1 - Built in Height (Flat), 2 - Built one layer file (Режимы карт: 0 - в форме рельефа, 1 - плоские срезы, 2 - одна карта для файла без глубин и высот)
		Set MinHLayer = 600      ' Min H, m минимальная отметка высоты
		Set StepHLayer = 100	 ' Step between layers шаг между срезами по высотам
		Set ThicknessLayer = 20  ' Layer thickness толщина слоя для осреднения
		Set NumHLayer = 8		    ' Number layers количество слоев
		Set StepTop = 20    		   ' Step page units between maps расстояние по Y между графиками на листе
		Set StepLeft = 20				 ' Step page units between maps расстояние по X между графиками на листе

'''''''''''''''''''''''''''''''''''''''''''''''''''''''SourceOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set SourcePath = "d:\Grids\"
		Set SourceFile = "d:\Grids\map.txt"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''GridsNames'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Field(1) = "Ro_"         : Arr(1) = "log10_ro"       ' В Arr(i) вписывать название столбцов из таблицы
		Field(2) = "Ita_"        		: Arr(2) = "ita"
		Field(3) = "MF_"        : Arr(3) = "MF"
		Field(4) = "1MF_"              : Arr(4) = "1MF"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set Wks = SurferApp.Documents.Open(SourceFile)
CRows = Wks.Columns(1).Count
Debug.Print CRows
CCols = Wks.Rows(1).Count
Debug.Print CCols
Set WksHandle = Wks.Cells(Row:=1,Col:=1,LastRow:=1,LastCol:=CCols)
i=1
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
For i = 1 To 4
	For Each vCell In WksHandle
		If vCell.Value = Arr(i) Then
			zCol(i) = vCell.Column
			Exit For
		End If
	Next
Next
If zCol(1)=0 Then
	Beep
	MsgBox("Нет параметра "& Arr(1) &" в таблице")
	SurferApp.ScreenUpdating = 1
	End
End If
If Built = 2 Then
	GoTo b
End If
For Each vCell In WksHandle
	If vCell.Value = "Z" Or vCell.Value = "H" Then
		Height = vCell.EntireColumn.Name
   		Height = Replace(Height, ":", "_"): Height = Replace(Height, "0", "", 3): Height = Replace(Height, "1", "")
   		hCol = vCell.Column
   		MinH = Wks.Columns(hCol).Statistics.Minimum
   		MaxH = Wks.Columns(hCol).Statistics.Maximum
   		Exit For
	End If
Next
If Built = 1 Then
	If MinH>MinHLayer Then
		Beep
		MsgBox("Заданного MinHLayer = "& MinHLayer &" нет в таблице")
		SurferApp.ScreenUpdating = 1
	End
	End If
End If
For Each vCell In WksHandle
	If vCell.Value = "H0" Or vCell.Value = "Depth" Then
		H0 = vCell.EntireColumn.Name
   		H0 = Replace(H0, ":", "_"): H0 = Replace(H0, "0", "", 3): H0 = Replace(H0, "1", "")
   		h0Col = vCell.Column
   		MinH0 = Wks.Columns(h0Col).Statistics.Minimum
   		MaxH0 = Wks.Columns(h0Col).Statistics.Maximum
   		Exit For
	End If
Next
'''''''''''''''''''''''''''''''' Find H0 ''''''''''''''''''''''''''''''''''''''''
If Built = 0 Then
	MaxH = 0
	If H0 = Empty Then
		CRows = Wks.Columns(1).Count
		Debug.Print CRows
		CCols = Wks.Rows(1).Count
		Debug.Print CCols
		Set WksRange = Wks.Cells(Row:=1,Col:=1,LastRow:=CRows,LastCol:=CCols+1)
		WksRange.Sort(Col1:=xCol, Order1:=wksSortAscending, Col2:=yCol, Order2:=wksSortAscending, Col3:=hCol, Order3:=wksSortDescending, Header:=True, MatchCase:=False)
		Dim Arr0 ()
		ReDim Arr0 (CRows, CCols)
		Dim H00
		ReDim H00(CRows, 1)
        Arr0 = WksRange.Value
        H0=Arr0(2, hCol)

        For i=1 To UBound(Arr0, 1)-2
			If Arr0(i+1,hCol)>Arr0(i+2,hCol) Then
				H00(i,0) = Arr0(i+1,hCol)-H0
				Else
				H00(i,0) = Arr0(i+1,hCol)-H0
				H0=Arr0(i+2,hCol)
			End If
        Next

		Set NewWksRange = Wks.Cells(Row:=1,Col:=CCols+1,LastRow:=CRows,LastCol:=CCols+1)
		H00(UBound(H00)-1,0)=Arr0(UBound(Arr0),hCol)-H0
		NewWksRange.Value = H00
		Wks.Cells(1, CCols+1).Value = "H0"
	End If
End If

If Built = 0 And Left(H0, 1) <> "_" Then
	H0 = Wks.Cells(Row:=1,Col:=CCols+1,LastRow:=1,LastCol:=CCols+1).EntireColumn.Name
	H0 = Replace(H0, ":", "_"): H0 = Replace(H0, "0", "", 3): H0 = Replace(H0, "1", "")
End If

H = MinHLayer

b:
On Error GoTo FileError
j=1
i=1
'''''''''''''''''''''''''''''''' Create Grid ''''''''''''''''''''''''''''''''''''''''
	While H <= MaxH
		While zCol(i) <> 0
			H1 = H - ThicknessLayer/2: H2 = H + ThicknessLayer/2
			If Built = 0 Then
				Srf = H0 & "<= "& H1-0.01 &" Or "& H0 &" >= "& H2+0.01 &"
			ElseIf Built = 1 Then
				Srf = Height & "<= "& H1-0.01 &" Or "& Height &" >= "& H2+0.01 &"
			End If
			Zmin = Wks.Columns(zCol(i)).Statistics.Minimum
			Zmax = Wks.Columns(zCol(i)).Statistics.Maximum
			SurferApp.GridData6(SourceFile, ExclusionFilter:= Srf, ZLimitMin:=Zmin, ZLimitMax:=Zmax, ZLimitMinMethod:=2, ZLimitMaxMethod:=2, SearchRad1:=SearchRad, SearchRad2:=SearchRad, BlankOutsideHull:= BlankOutsideHull, InflateHull:= InflateHull, xCol:=xCol, yCol:= yCol, zCol:=zCol(i), xSize:= xSize, ySize:= ySize, Algorithm:=Alg, OutGrid:= SourcePath & Field(i) & H & ".grd")
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
			If i <= 1.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("Blue-OrangeRed")' "Blue-OrangeRed"
		ElseIf i <= 2.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("Scale_ita")
		ElseIf i <= 3.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("DarkPurple-Yellow")
		ElseIf i <= 4.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("DarkPurple-Yellow")
		End If
			MapFrame.Top=-70+j*StepTop
			MapFrame.Left=-30+Int(i-0.00001)*StepLeft
			MapFrame.ScalingRatio = ScalingRatio
				For Each axis In Axes
					axis.Visible = axisVisible
				Next
			ContourLayer.ShowColorScale = ShowColorScale
			ContourLayer.ShowMajorLabels = 0
			Debug.Print Field(i) & H & ".grd"
			Kill(SourcePath & Field(i) & H & ".grd")
			i=i+1
			If i > 4 Then
				Exit While
			End If
		Wend
		i=1
		j=j+1
		H = H + StepHLayer
		If H >= MinHLayer + (NumHLayer*StepHLayer) Or (Built = 0 And H > 0) Then
			Exit While
		End If
	Wend
Wks.Close 2
'Error instructions
	FileError:
	Debug.Print  "Error:	" + grid_file + "						" + Err.Description
	Resume Next
SurferApp.ScreenUpdating = True
End Sub
