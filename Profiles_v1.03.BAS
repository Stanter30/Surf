Sub Main
Debug.Clear
Dim Arr(4) As String, Field(4) As String, zCol(4) As Integer
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
If Plot = "map.txt" Then
	Plot.Close 2
	Set Plot = SurferApp.ActiveDocument
End If
Set Shapes = Plot.Shapes
SurferApp.ScreenUpdating = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''GridOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set xSize = 10 					' X Step grid, m
		Set ySize = 10	   			' Y Step grid, m
		Set NumPasses = 1	       ' Number Passes Filter
		Set SearchRad = 200	 ' Radius Search
		Set axisVisible =1 		     	' On or Off Axes
		Set Pickets = 1				' On or Off Pickets
		Set ShowColorScale =0	   ' On or Off Scale
		Set OnFilter = 1				  ' On or Off Filter
		Set OnBlank = 1				   ' On or Off Blanking
'''''''''''''''''''''''''''''''''''''''''''''''''''''''ProfilesOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set MinH = -400               ' Min H, m
		Set MinProf = 3: MaxProf = 3	 ' Min & Max profile 
		Set StepProfiles = 1		                ' Step profile numbers
		Set MajorInterval = 50	              ' Left Axis Major interval
		Set StepLeftShift = 20    		            ' Step page units between maps
		Set ScalingRatio = 100		           	' Map scaling ratio 1/N
		Set SortXY = "X"							'Sort col (working only without Dist in table)
		Set FontFace = "Times New Roman"					'Font
		Set FontSize = 11							'Font Size
'''''''''''''''''''''''''''''''''''''''''''''''''''''''SourceOptions '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Set SourcePath = "D:\Grids\"
		Set SourceFile = "D:\Grids\map.txt"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''GridsNames''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Field(1) = "Ro_"         : Arr(1) = "log10_ro"       ' Names columns Z in SourceFile
		Field(2) = "Ita_"        		: Arr(2) = "ita"
		Field(3) = "Tau_"        : Arr(3) = "tau"
		Field(4) = "MF_"              : Arr(4) = "MF"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set Wks = SurferApp.Documents.Open(SourceFile)
CRows = Wks.Columns(1).Count
Debug.Print CRows
CCols = Wks.Rows(1).Count
Debug.Print CCols
Set WksRange = Wks.Cells(Row:=1,Col:=1,LastRow:=CRows,LastCol:=CCols)
Set WksHandle = Wks.Cells(Row:=1,Col:=1,LastRow:=1,LastCol:=CCols)
i=1
For Each vCell In WksHandle
	If vCell.Value = "Dist" Or vCell.Value = "dist" Then
		xCol = vCell.Column
	End If
Next
If xCol="" Then
	Beep
	MsgBox("Парамера Dist нет в таблице с данными!")
	SurferApp.ScreenUpdating = 1
	End
End If
For Each vCell In WksHandle
	If vCell.Value = "H" Or vCell.Value = "h" Or vCell.Value = "Z" Or vCell.Value = "z" Then
	   	yCol = vCell.Column
	   	MinHTest = Wks.Columns(yCol).Statistics.Minimum
	   	MaxH = Wks.Columns(yCol).Statistics.Maximum
	   	If MinHTest>MinH Then
			Beep
			MsgBox("Заданного MinH нет в таблице с данными!")
			SurferApp.ScreenUpdating = 1
			End
		End If
		For Each vCell In WksHandle
			If vCell.Value = Arr(i) Then
				zCol(i) = vCell.Column
				i=i+1
				If i > 4 Then
					Exit For
				End If
			End If
		Next
	End If
Next
If zCol(1)=0 Then
	Beep
	MsgBox("Нет параметра "& Arr(1) &"либо Z в таблице")
	SurferApp.ScreenUpdating = 1
	End
End If
For Each vCell In WksHandle
	If vCell.Value = "PR" Or vCell.Value = "Pr" Or vCell.Value = "Profile" Or vCell.Value = "Prof" Or vCell.Value = "Line" Then
		ProfCol = vCell.Column
		Prof = vCell.EntireColumn.Name
   		Prof = Replace(Prof, ":", "_"): Prof = Replace(Prof, "0", "", 3): Prof = Replace(Prof, "1", "")
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = "PK" Or vCell.Value = "Pk" Or vCell.Value = "Point" Then
		PKCol = vCell.Column
		PK= vCell.EntireColumn.Name
   		PK = Replace(PK, ":", "_"): PK = Replace(PK, "0", "", 3): PK = Replace(PK, "1", "")
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = "X" Or vCell.Value = "x" Then
		XSortCol = vCell.Column
		Exit For
	End If
Next
For Each vCell In WksHandle
	If vCell.Value = "Y" Or vCell.Value = "y" Then
		YSortCol = vCell.Column
		Exit For
	End If
Next
If SortXY="X" Or SortXY="x" Then
	SortCol=XSortCol
Else
	SortCol=YSortCol
End If
'''''''''''''''''''''''''''Calc distance''''''''''''''''''''''''''''''''''''''''''
If xCol = Empty Then
	Set WksRange = Wks.Cells(Row:=1,Col:=1,LastRow:=CRows,LastCol:=CCols)
	WksRange.Sort(Col1:=ProfCol, Order1:=wksSortAscending, Col2:=SortCol, Order2:=wksSortAscending, Header:=True, MatchCase:=False)
	Dim ArrXY As Variant
	ArrXY = Wks.Cells(Row:=2,Col:=1,LastRow:=CRows+1,LastCol:=CCols+1)
	i=1
	X=ArrXY(1,XSortCol): Y=ArrXY(1,YSortCol)
	While ArrXY(i,1) <> ""
		If ArrXY(i,ProfCol)=ArrXY(i+1,ProfCol) Then
			dist=Sqr((ArrXY(i,XSortCol)-X)^2+(ArrXY(i,YSortCol)-Y)^2)
			ArrXY(i, CCols+1)=dist
			i=i+1
		Else
			ArrXY(i, CCols+1)=dist
			X=ArrXY(i+1,XSortCol): Y=ArrXY(i+1,YSortCol)
			i=i+1
		End If
	Wend
	Wks.Cells(Row:=2,Col:=1,LastRow:=CRows,LastCol:=CCols+1).Value=ArrXY
	Wks.Cells(Row:=1,Col:=CCols+1,LastRow:=1,LastCol:=CCols+1).Value="Dist"
	xCol=CCols+1
	Wks.SaveAs(SourceFile, "Delimiter=tab")
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(Dir$(SourcePath & MinProf & ".bln*")) <> 0 Then
	If Pickets = 1 And Len(Dir$(SourcePath & MinProf & ".txt*")) <> 0  Then
		x=111
		'Хрюк
		ElseIf Pickets = 1 And Len(Dir$(SourcePath & MinProf & ".txt*")) = 0 Then
			GoTo n
			'Хрюк
	End If
	'Хрюк
Else
    '''''''''''''''''''''''''''''''Nightmare''''''''''''''''''''''''''''''''''''''''
n:
CRows = Wks.Columns(1).Count
Debug.Print CRows
CCols = Wks.Rows(1).Count
Debug.Print CCols
Set WksRange = Wks.Cells(Row:=1,Col:=1,LastRow:=CRows,LastCol:=CCols)
If ProfCol=Empty Then
	Beep
	MsgBox("Не задан профиль!")
	SurferApp.ScreenUpdating = 1
	End
End If
WksRange.Sort(Col1:=ProfCol, Order1:=wksSortAscending, Col2:=xCol, Order2:=wksSortAscending, Col3:= yCol, Order3:=wksSortDescending, Header:=True, MatchCase:=False)
Wks.SpatialFilter(xSourceCol:=ProfCol , ySourceCol:=xCol, zSourceCol:=yCol, xTargetCol:=CCols+1, yTargetCol:=CCols+2, zTargetCol:=CCols+3, DupMethod:=srfDupFirst, xDupTol:=0.1, yDupTol:=0.1)
Wks.SpatialFilter(xSourceCol:=PKCol , ySourceCol:=0, zSourceCol:=0, xTargetCol:=CCols+4, yTargetCol:=0, zTargetCol:=0, DupMethod:=srfDupFirst, xDupTol:=0.1, yDupTol:=0.1)
Set WksRangeNew = Wks.Cells(Row:=1,Col:=1,LastRow:=CRows,LastCol:=CCols+4)
WksRangeNew.Sort(Col1:=CCols+1, Order1:=wksSortAscending, Col2:=CCols+2, Order2:=wksSortAscending, Header:=True, MatchCase:=False)
WksRange.Delete(Direction:=wksDeleteColumns)
Wks.Cells(1, 1).Value = "Pr": Wks.Cells(1, 2).Value = "Dist": Wks.Cells(1, 3).Value = "H": Wks.Cells(1, 4).Value = "Pk"
Set WksRangeForDel = Wks.Cells(Row:=1,Col:=5,LastRow:=CRows,LastCol:=6)
WksRangeForDel.Delete(Direction:=wksDeleteColumns)
CRows = Wks.Columns(1).Count
Debug.Print CRows
CCols = Wks.Rows(1).Count
Debug.Print CCols
Dim Arr0 As Variant
Dim Arr1 ()
ReDim Arr1(CRows+1, 2)
Dim ArrPK
Arr0 = Wks.Cells(Row:=2,Col:=1,LastRow:=CRows+1,LastCol:=CCols)
Set WksBln = SurferApp.Documents.Add(srfDocWks)
i = 0
j = 0
c = 0
While i < CRows - 1
i=i+1
j=j+1
	PrBln = Arr0(i, 1)
	If Arr0(i, 1) = Arr0(i + 1, 1) Then
	    Arr1(j-1, 0) = Arr0(i, 2)
        Arr1(j-1, 1) = Arr0(i, 3)
        If PKCol = Empty Then
			c=c+1
        	Arr1(j-1, 2) = c
        Else
			Arr1(j-1, 2) = Arr0(i, 4)
        End If
        ArrPK = Arr1
        Else
			Arr1(j-1, 0) = Arr0(i, 2)
            Arr1(j-1, 1) = Arr0(i, 3)
            Arr1(j, 0) = Arr0(i, 2) + 0.1
            Arr1(j, 1) = Arr0(i, 3)
            Arr1(j + 1, 0) = Arr0(i, 2) + 0.1
            Arr1(j + 1, 1) = MinH - 0.1
            Arr1(j + 2, 0) = -0.1
            Arr1(j + 2, 1) = MinH - 0.1
			WksBln.Cells(Row:=3,Col:=1,LastRow:=CRows,LastCol:=2).Value = Arr1
			WksBln.Cells(Row:=1,Col:=3,LastRow:=CRows,LastCol:=3).Clear
			WksBln.Cells(1, 1).Value = j + 4
			WksBln.Cells(1, 2).Value = 0
			WksBln.Cells(2, 1).Value = -0.1
			WksBln.Cells(2, 2).Value = WksBln.Cells(3, 2).Value
            WksBln.SaveAs(Filename:=SourcePath & PrBln & ".bln", FileFormat:= 14)
            If Pickets = 1 Then
            	Set DistPK = SurferApp.Documents.Add(srfDocWks)
				DistPK.Cells(Row:=2,Col:=1,LastRow:=CRows,LastCol:=2).Value = ArrPK
				DistPK.Cells(1, 1).Value = "Dist"
				DistPK.Cells(1, 2).Value = "H"
				DistPK.Cells(1, 3).Value = "PK"
	            DistPK.SaveAs(Filename:=SourcePath & PrBln & ".txt")
	            DistPK.Close 2
            End If
            BlnRows = WksBln.Columns(1).Count
            Set WRB = WksBln.Cells(Row:=1,Col:=1,LastRow:=BlnRows,LastCol:=2)
			WRB.Clear
            ReDim Arr1(CRows, 2)
            j = 0
            c = 0
	End If
Wend
WksBln.Close 2
Wks.Close 2
End If
Set Wks = SurferApp.Documents.Open(SourceFile)
i=1
j=1
''''''''''''''''''''''''''''''''''''''''''''''' CreateGrid'''''''''''''''''''''''''''''''''''''''''''''''''
	For Pr = MinProf To MaxProf Step StepProfiles
		'On Error GoTo q
		Set Wks1 = SurferApp.Documents.Open(SourcePath & Pr & ".bln")
		xMax = Wks1.Columns(1).Statistics.Maximum
		Wks1.Close 2: BLN_file = SourcePath & Pr & ".bln"
		While zCol(i) <> 0
			Zmin = Wks.Columns(zCol(i)).Statistics.Minimum
			Zmax = Wks.Columns(zCol(i)).Statistics.Maximum
			'Plot.Activate
			SurferApp.GridData6(SourceFile, ExclusionFilter:= Prof & "<>" + Pr + "", BlankOutsideHull:= 0, SearchRad1:=SearchRad, SearchRad2:=SearchRad, ZLimitMin:=Zmin, ZLimitMax:=Zmax, ZLimitMinMethod:=2, ZLimitMaxMethod:=2, yMin:=MinH, xMax:=xMax, xCol:=xCol, yCol:= yCol, zCol:=zCol(i), xSize:= xSize, ySize:= ySize, Algorithm:=srfKriging, OutGrid:= SourcePath & Field(i) & Pr & ".grd")
			If OnFilter = 1 Then
				SurferApp.GridFilter2(InGrid:= SourcePath & Field(i) & Pr & ".grd", Filter:=srfFilterGaussian, OutGrid:= SourcePath & Field(i) & Pr & ".grd", EdgeOp:=3, BlankOp:=2, NumPasses:=NumPasses)
			End If
			If OnBlank = 1 Then
				SurferApp.GridBlank(SourcePath & Field(i) & Pr & ".grd", BLN_file, SourcePath & Field(i) & Pr & ".grd")
			End If
'''''''''''''''''''''''''''''''''' 'Create ContourMaps '''''''''''''''''''''''''''''''''''''''''''''''''
			Set MapFrame = Shapes.AddContourMap(GridFileName:=SourcePath & Field(i) & Pr & ".grd")
			If Pickets = 1 Then
				Set PostLayer = Shapes.AddPostLayer(Map:=MapFrame, DataFileName:=SourcePath & Pr & ".txt", xCol:=1, yCol:=2, LabCol:=3)
			End If
			Set Axes = MapFrame.Axes
			Set ContourLayer = MapFrame.Overlays(1)
			ContourLayer.FillContours = True
			If i <= 1.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("BlueGreen-OrangeRed")
		ElseIf i <= 2.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("Rainbow3")
		ElseIf i <= 3.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("Blue-Red")
		ElseIf i <= 4.001 Then
			ContourLayer.FillForegroundColorMap.LoadPreset ("Purple-Yellow")
		End If
			MapFrame.Left=-150+j*StepLeftShift
			MapFrame.Top=30-Int(i-0.00001)*12
			MapFrame.ScalingRatio = ScalingRatio
				For Each axis In Axes
					Axis.Visible = axisVisible
				Next
			Axes(3).SetScale(FirstMajorTick:=MinH, LastMajorTick:=MaxH, MajorInterval:=MajorInterval)
			Set title=SourcePath & Field(i) & Pr & ".grd"
			Axes(1).Title = "Расстояние, м"
			Axes(2).AxisLine.Style = "invisible"
			Axes(2).Title = "Профиль "& Pr &""
			Axes(3).Title = "Абсолютные отметки, м"
			For ax = 1 To 3
				Axes(ax).TitleFont.Face = FontFace
				Axes(ax).TitleFont.Size  = FontSize
			Next
			ContourLayer.ShowColorScale = ShowColorScale
			ContourLayer.ShowMajorLabels = 0
			Debug.Print Field(i) & Pr & ".grd"
			Kill(SourcePath & Field(i) & Pr & ".grd")
			i = i+1
			If i > 4 Then
				Exit While
			End If
		Wend
		j=j+1
		q:
		On Error GoTo 0
		i=1
	Next
Wks.Close 2
SurferApp.ScreenUpdating = True
End Sub
