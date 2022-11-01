Sub main
Debug.Clear
Set SurferApp = GetObject(, "Surfer.Application")
Set Plot = SurferApp.ActiveDocument
Set Shapes = Plot.Shapes
Set Selection = Plot.Selection
SurferApp.Visible = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''ExportOptions ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Set dpi = 200 'DPI
	Set Extention = ".tif" 'Format
	Set Direct = "D:\Grids\" 'Path
	Set Prefix = "" ' Prefix
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
i=0
		While i < Plot.Shapes.Count
			i = i+1
			Set MapFrame = Plot.Shapes.Item(i)
			If MapFrame.Type = srfShapeMapFrame Then
			NameContours = MapFrame.Overlays(1).Name
			NameContours = Replace(NameContours, "Contours-", "")
			NameContours = Replace(NameContours, ".grd", "")
			NameContours = Prefix & NameContours
			Selection.DeselectAll
			MapFrame.Select
				ExportFile = Direct & NameContours & Extention
			 	Plot.Export(FileName:=ExportFile, SelectionOnly:=True, Options:="Defaults=1,HDPI="& dpi &",VDPI="& dpi &",ColorDepth=8,")
			MapFrame.Selected = Clear
			End If
		Wend
SurferApp.Visible = True
End Sub
