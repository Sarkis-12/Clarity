# Clarity
HOH Database
HOH Database Sub Macro1() ' ' Macro1 Macro ' ' Keyboard Shortcut: Ctrl+Shift+F ' Dim Rng As Range

ActiveSheet.Name = "ClaretData"

ActiveSheet.UsedRange.Select

Selection.RemoveSubtotal
ActiveCell.Select




ActiveCell.Rows("1:1").EntireRow.Select
Selection.Find(What:="Order number", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
ActiveCell.Columns("A:A").EntireColumn.Select
Selection.Copy
Sheets.Add.Name = "PickSheets"

Sheets("PickSheets").Select

ActiveSheet.Paste
ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
ActiveCell.Offset(2, 1).Range("A1").Select

Sheets("ClaretData").Select

ActiveCell.Rows("1:1").EntireRow.Select
Selection.Find(What:="Customer Name", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
ActiveCell.Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PickSheets").Select
ActiveCell.Offset(-2, 0).Range("A1").Select
ActiveSheet.Paste
ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
ActiveCell.Offset(30, 2).Range("A1").Select
Sheets("ClaretData").Select
ActiveCell.Rows("1:1").EntireRow.Select






 Set Rng = Selection.Find(What:="Customer Country", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)

     If Rng Is Nothing Then

      Set Rng = Selection.Find(What:="Delivery Country", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)




Rng.Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PickSheets").Select
ActiveCell.Offset(-30, -1).Range("A1").Select
ActiveSheet.Paste
ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit

Else

Rng.Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PickSheets").Select
ActiveCell.Offset(-30, -1).Range("A1").Select
ActiveSheet.Paste
ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit

End If

Sheets("ClaretData").Select

ActiveCell.Rows("1:1").EntireRow.Select Selection.Find(What:="Product ID", After:=ActiveCell, LookIn:=xlFormulas _ , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _ MatchCase:=False, SearchFormat:=False).Activate ActiveCell.Columns("A:A").EntireColumn.Select Application.CutCopyMode = False Selection.Copy Sheets("PickSheets").Select ActiveCell.Offset(0, 1).Range("A1").Select ActiveSheet.Paste ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit Sheets("ClaretData").Select ActiveCell.Rows("1:1").EntireRow.Select

Selection.Find(What:="Size", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
ActiveCell.Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PickSheets").Select
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveSheet.Paste
Sheets("ClaretData").Select
ActiveCell.Rows("1:1").EntireRow.Select

Selection.Find(What:="Sales Order Quantity", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
ActiveCell.Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("PickSheets").Select
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveSheet.Paste
Selection.ColumnWidth = 8.57
ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
ActiveCell.Offset(0, 1).Range("A1").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "SKU"
ActiveCell.Offset(1, 0).Range("A1").Select
ActiveCell.FormulaR1C1 = "=RC[-3]&"" ""&RC[-2]"
ActiveCell.Select
Selection.AutoFill Destination:=ActiveCell.Range("A1:A490")
ActiveCell.Range("A1:A490").Select
ActiveWindow.SmallScroll Down:=-27
ActiveCell.Columns("A:A").EntireColumn.Select
Selection.Copy
ActiveCell.Offset(0, -3).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft
ActiveCell.Offset(0, -2).Columns("A:A").EntireColumn.Select
Selection.Delete Shift:=xlToLeft
ActiveCell.Select
ActiveCell.FormulaR1C1 = "Qty"
ActiveCell.Offset(1, 0).Range("A1").Select
ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 6.71

ActiveSheet.UsedRange.Select

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

 ActiveWorkbook.Worksheets("PickSheets").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("PickSheets").Sort.SortFields.Add Key:=ActiveCell. _
    Range("A1:A490"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
With ActiveWorkbook.Worksheets("PickSheets").Sort
    .SetRange ActiveSheet.UsedRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With



Sheets.Add.Name = "Pivot"
Sheets("Pivot").Select

Sheets("PickSheets").Select









ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "PickSheets!R1C1:R491C5", Version:=6).CreatePivotTable TableDestination:= _
    "Pivot!R1C1", TableName:="PivotTable", DefaultVersion:=6
Sheets("Pivot").Select
Cells(1, 1).Select
ActiveCell.Offset(6, 1).Range("A1").Select
With ActiveSheet.PivotTables("PivotTable")
    .InGridDropZones = True
    .RowAxisLayout xlTabularRow
End With
With ActiveSheet.PivotTables("PivotTable").PivotFields("SKU")
    .Orientation = xlRowField
    .Position = 1
End With
With ActiveSheet.PivotTables("PivotTable").PivotFields("Qty")
    .Orientation = xlRowField
    .Position = 2
End With
ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
    "PivotTable").PivotFields("Qty"), "Count of Qty", xlCount
With ActiveSheet.PivotTables("PivotTable").PivotFields("Count of Qty")
    .Caption = "Sum of Qty"
    .Function = xlSum
End With
ActiveCell.Offset(-5, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "Location"
ActiveCell.Offset(2, 0).Range("A1").Select








Sheets("PickSheets").Select









ActiveCell.Range("A1:G1").Select
Selection.Font.Bold = True
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
With Selection.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
End With
ActiveCell.Offset(0, 5).Range("A1").Select
ActiveCell.FormulaR1C1 = "Location"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "Dims & Weight"
ActiveCell.Offset(1, 0).Range("A1").Select
ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 16.86
ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.ColumnWidth = 11.57
ActiveCell.Offset(0, -1).Columns("A:A").EntireColumn.Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.Select
ActiveCell.FormulaR1C1 = "Carton"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "Carton"
ActiveCell.Offset(0, -1).Range("A1").Select
ActiveCell.FormulaR1C1 = "Location"
ActiveCell.Offset(2, 0).Range("A1").Select
ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 11.71
ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 8.86
ActiveSheet.UsedRange.Select

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Application.PrintCommunication = True
ActiveSheet.PageSetup.PrintArea = ""
Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = ""
    .RightFooter = ""
    .LeftMargin = Application.InchesToPoints(0.708661417322835)
    .RightMargin = Application.InchesToPoints(0.708661417322835)
    .TopMargin = Application.InchesToPoints(0.748031496062992)
    .BottomMargin = Application.InchesToPoints(0.748031496062992)
    .HeaderMargin = Application.InchesToPoints(0.31496062992126)
    .FooterMargin = Application.InchesToPoints(0.31496062992126)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintComments = xlPrintNoComments
    .PrintQuality = 600
    .CenterHorizontally = False
    .CenterVertically = False
    .Orientation = xlPortrait
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 100
    .PrintErrors = xlPrintErrorsDisplayed
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .ScaleWithDocHeaderFooter = True
    .AlignMarginsHeaderFooter = True
    .EvenPage.LeftHeader.Text = ""
    .EvenPage.CenterHeader.Text = ""
    .EvenPage.RightHeader.Text = ""
    .EvenPage.LeftFooter.Text = ""
    .EvenPage.CenterFooter.Text = ""
    .EvenPage.RightFooter.Text = ""
    .FirstPage.LeftHeader.Text = ""
    .FirstPage.CenterHeader.Text = ""
    .FirstPage.RightHeader.Text = ""
    .FirstPage.LeftFooter.Text = ""
    .FirstPage.CenterFooter.Text = ""
    .FirstPage.RightFooter.Text = ""
End With
Application.PrintCommunication = True
Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(5), _
    Replace:=True, PageBreaks:=True, SummaryBelowData:=True
Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .PrintTitleRows = "$1048573:$1048573"
    .PrintTitleColumns = ""
End With
Application.PrintCommunication = True
ActiveSheet.PageSetup.PrintArea = ""
Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = "Page &P of &N"
    .RightFooter = ""
    .LeftMargin = Application.InchesToPoints(0)
    .RightMargin = Application.InchesToPoints(0)
    .TopMargin = Application.InchesToPoints(0.748031496062992)
    .BottomMargin = Application.InchesToPoints(0.748031496062992)
    .HeaderMargin = Application.InchesToPoints(0.31496062992126)
    .FooterMargin = Application.InchesToPoints(0.31496062992126)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintComments = xlPrintNoComments
    .PrintQuality = 600
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlLandscape
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 85
    .PrintErrors = xlPrintErrorsDisplayed
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .ScaleWithDocHeaderFooter = True
    .AlignMarginsHeaderFooter = True
    .EvenPage.LeftHeader.Text = ""
    .EvenPage.CenterHeader.Text = ""
    .EvenPage.RightHeader.Text = ""
    .EvenPage.LeftFooter.Text = ""
    .EvenPage.CenterFooter.Text = ""
    .EvenPage.RightFooter.Text = ""
    .FirstPage.LeftHeader.Text = ""
    .FirstPage.CenterHeader.Text = ""
    .FirstPage.RightHeader.Text = ""
    .FirstPage.LeftFooter.Text = ""
    .FirstPage.CenterFooter.Text = ""
    .FirstPage.RightFooter.Text = ""
End With
Application.PrintCommunication = True

 Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .PrintTitleRows = "$1:$1"
    .PrintTitleColumns = ""
End With
Application.PrintCommunication = True
ActiveSheet.PageSetup.PrintArea = ""
Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = "Page &P of &N"
    .RightFooter = ""
    .LeftMargin = Application.InchesToPoints(0)
    .RightMargin = Application.InchesToPoints(0)
    .TopMargin = Application.InchesToPoints(0.748031496062992)
    .BottomMargin = Application.InchesToPoints(0.748031496062992)
    .HeaderMargin = Application.InchesToPoints(0.31496062992126)
    .FooterMargin = Application.InchesToPoints(0.31496062992126)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintComments = xlPrintSheetEnd
    .PrintQuality = 600
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlLandscape
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 85
    .PrintErrors = xlPrintErrorsDisplayed
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .ScaleWithDocHeaderFooter = True
    .AlignMarginsHeaderFooter = True
    .EvenPage.LeftHeader.Text = ""
    .EvenPage.CenterHeader.Text = ""
    .EvenPage.RightHeader.Text = ""
    .EvenPage.LeftFooter.Text = ""
    .EvenPage.CenterFooter.Text = ""
    .EvenPage.RightFooter.Text = ""
    .FirstPage.LeftHeader.Text = ""
    .FirstPage.CenterHeader.Text = ""
    .FirstPage.RightHeader.Text = ""
    .FirstPage.LeftFooter.Text = ""
    .FirstPage.CenterFooter.Text = ""
    .FirstPage.RightFooter.Text = ""
End With
Application.PrintCommunication = True
ActiveCell.Offset(1, 5).Range("A1").Select
ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-2],Pivot!C[-5]:C[-3],3,FALSE),"" "")"
ActiveCell.Select
Selection.AutoFill Destination:=ActiveCell.Range("A1:A507"), Type:= _
    xlFillDefault
ActiveCell.Range("A1:A507").Select
ActiveWindow.SmallScroll Down:=-9
ActiveWindow.ScrollRow = 464
ActiveWindow.ScrollRow = 460
ActiveWindow.ScrollRow = 453
ActiveWindow.ScrollRow = 445
ActiveWindow.ScrollRow = 437
ActiveWindow.ScrollRow = 430
ActiveWindow.ScrollRow = 418
ActiveWindow.ScrollRow = 406
ActiveWindow.ScrollRow = 366
ActiveWindow.ScrollRow = 328
ActiveWindow.ScrollRow = 280
ActiveWindow.ScrollRow = 251
ActiveWindow.ScrollRow = 233
ActiveWindow.ScrollRow = 217
ActiveWindow.ScrollRow = 199
ActiveWindow.ScrollRow = 170
ActiveWindow.ScrollRow = 156
ActiveWindow.ScrollRow = 135
ActiveWindow.ScrollRow = 113
ActiveWindow.ScrollRow = 108
ActiveWindow.ScrollRow = 87
ActiveWindow.ScrollRow = 80
ActiveWindow.ScrollRow = 73
ActiveWindow.ScrollRow = 51
ActiveWindow.ScrollRow = 44
ActiveWindow.ScrollRow = 19
ActiveWindow.ScrollRow = 18
ActiveWindow.ScrollRow = 17
ActiveWindow.ScrollRow = 15
ActiveWindow.ScrollRow = 13
ActiveWindow.ScrollRow = 9
ActiveWindow.ScrollRow = 7
ActiveWindow.ScrollRow = 4
ActiveWindow.ScrollRow = 1
ActiveCell.Offset(-1, -5).Range("A1").Select

End Sub
