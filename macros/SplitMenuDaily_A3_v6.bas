Attribute VB_Name = "DailyMenuPrintV6"
Option Explicit

'================================================================
' SplitMenuDaily_A3_V6
' Split a stacked daily-menu sheet into one sheet per day.
' Each daily sheet prints on ONE A3 portrait page.
'
' Target: the ACTIVE workbook and its active sheet, so the
' macro can live on the PC itself - e.g. in the Personal Macro
' Workbook (PERSONAL.XLSB) - and process any menu workbook /
' sheet you open and select. Guards stop it on PERSONAL.XLSB
' itself and on chart sheets.
'
' V6 change (from V5):
'  - The three remaining menu blocks (joshoku / nansai /
'    blender) are set to EQUAL total width, widened so the
'    PRINTED table just fills the width of A3 portrait.
'    A day is always taller than one page, so fit-to-page
'    shrinks everything for height; the widths are widened
'    by that same shrink factor to compensate (on screen the
'    table looks wider than the page - that is expected).
'    The column proportions inside each block are kept.
'
' V5 changes (from V4):
'  - The Iwato-joshoku block (columns A-I) is removed from each
'    new sheet and the other blocks shift left, so the print
'    range is A:AA (27 columns). The field line, the date line
'    and the meal headers (in column A) are restored into the
'    new column A after the shift.
'  - Stamp pictures are NOT deleted. The source sheet is left
'    completely unchanged, and each day's stamp picture is
'    copied onto its daily sheet (and _B sheet), aligned to
'    the top-right corner of the table.
'
' Each daily sheet is built by copying ONLY that day's rows
' (about 100 rows) onto a fresh sheet - light and reliable.
' Row heights, formats and merged cells come with the row
' copy; column widths and header/footer are copied explicitly.
' Select menus: a <B> band is cut out of the main sheet into a
' "_B" sheet (day header + meal header + B band only).
'
' If it stops, the error message shows WHICH STEP failed (and
' the day being processed). Tell the author that step name and
' the error text - that pinpoints the cause.
'
' This file is ASCII only, so it never garbles (mojibake).
' Japanese strings for matching cell text are built from hex
' code points by the CJ() helper:
'   date header prefix : 5E74 6708 65E5 FF1A
'   daily sheet prefix : 65E5 5225 (+ "_")
'   breakfast/lunch/dinner : 671D 98DF / 663C 98DF / 5915 98DF
'   field line prefix  : 73FE 5834
'================================================================

Private gStep As String
Private gStampCount As Long

Public Sub SplitMenuDaily_A3_V6()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, i As Long, n As Long
    Dim dateRows() As Long
    Dim dateLabels() As String
    Dim startRow As Long, endRow As Long, blockLen As Long
    Dim sheetName As String
    Dim prefix As String
    Dim bSrc As Long, bEndSrc As Long, mealSrc As Long
    Dim r2 As Long, c As Long
    Dim selectCount As Long

    prefix = CJ("65E5 5225") & "_"
    gStep = "start"

    ' V4: process the workbook in front of you, not the
    ' workbook that stores this macro.
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No workbook is open.", vbExclamation
        Exit Sub
    End If
    If UCase$(wb.Name) = "PERSONAL.XLSB" Then
        MsgBox "Open the menu workbook first and select its sheet, then run again.", vbExclamation
        Exit Sub
    End If
    If Not TypeOf wb.ActiveSheet Is Worksheet Then
        MsgBox "Select a worksheet (not a chart sheet), then run again.", vbExclamation
        Exit Sub
    End If
    Set wsSrc = wb.ActiveSheet

    If wsSrc.Name Like prefix & "*" Then
        MsgBox "Please select the original menu sheet first, then run again.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrHandler

    '--- Delete existing daily sheets (safe re-run) ---
    gStep = "delete old daily sheets"
    For i = wb.Sheets.Count To 1 Step -1
        If wb.Sheets(i).Name Like prefix & "*" Then wb.Sheets(i).Delete
    Next i

    '--- Find date header rows in column A ---
    gStep = "find date rows"
    lastRow = wsSrc.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    n = 0
    For r = 1 To lastRow
        If CellText(wsSrc.Cells(r, 1)) Like CJ("5E74 6708 65E5 FF1A") & "*" Then
            n = n + 1
            ReDim Preserve dateRows(1 To n)
            ReDim Preserve dateLabels(1 To n)
            dateRows(n) = r
            dateLabels(n) = DateLabel(CellText(wsSrc.Cells(r, 1)))
        End If
    Next r
    If n = 0 Then
        MsgBox "No date header rows found in column A.", vbExclamation
        GoTo CleanUp
    End If

    ' V5: the source sheet is NOT modified. Stamps stay as
    ' they are and are copied onto each daily sheet instead.

    '--- Build one sheet per day ---
    selectCount = 0
    gStampCount = 0
    For i = 1 To n
        startRow = dateRows(i) - 2
        If startRow < 1 Then startRow = 1
        If i < n Then
            endRow = dateRows(i + 1) - 3
        Else
            endRow = lastRow
        End If
        blockLen = endRow - startRow + 1
        sheetName = prefix & dateLabels(i)
        gStep = "build sheet " & sheetName

        ' Find select-menu markers in SOURCE coordinates
        bSrc = 0: bEndSrc = 0: mealSrc = 0
        For r2 = startRow To endRow
            For c = 1 To 36
                If CellText(wsSrc.Cells(r2, c)) = "<B>" Then
                    bSrc = r2
                    Exit For
                End If
            Next c
            If bSrc > 0 Then Exit For
        Next r2
        If bSrc > 0 Then
            bEndSrc = endRow
            For r2 = bSrc + 1 To endRow
                If IsMeal(CellText(wsSrc.Cells(r2, 1))) Then
                    bEndSrc = r2 - 1
                    Exit For
                End If
            Next r2
            For r2 = bSrc - 1 To startRow Step -1
                If IsMeal(CellText(wsSrc.Cells(r2, 1))) Then
                    mealSrc = r2
                    Exit For
                End If
            Next r2
        End If

        ' New sheet with only this day's rows copied over
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
        CopyColumnWidths wsSrc, ws
        wsSrc.Rows(startRow & ":" & endRow).Copy Destination:=ws.Rows(1)
        Application.CutCopyMode = False

        ' Select menu: build the _B sheet first, then cut the
        ' B band out of the main daily sheet
        If bSrc > 0 Then
            gStep = "build _B sheet for " & sheetName
            BuildBSheet wb, wsSrc, sheetName & "_B", _
                        startRow, mealSrc, bSrc, bEndSrc
            ws.Rows((bSrc - startRow + 1) & ":" & (bEndSrc - startRow + 1)).Delete
            blockLen = blockLen - (bEndSrc - bSrc + 1)
            selectCount = selectCount + 1
            gStep = "build sheet " & sheetName
        End If

        ' V5: drop the Iwato-joshoku columns (A-I) and shift
        ' left, then copy this day's stamp onto the sheet.
        gStep = "remove Iwato columns " & sheetName
        RemoveIwatoColumns ws, blockLen
        gStep = "equalize block widths " & sheetName
        EqualizeBlockWidths ws, wsSrc, blockLen
        gStep = "copy stamp " & sheetName
        CopyStampToSheet wsSrc, ws, dateRows(i)
        ApplyPageSetup ws, wsSrc, blockLen
    Next i

    gStep = "finish"
    wb.Sheets(prefix & dateLabels(1)).Activate
    MsgBox "Done. Daily sheets created in '" & wb.Name & "': " & n & vbCrLf & _
           "Select-menu _B sheets: " & selectCount & vbCrLf & _
           "Stamps copied: " & gStampCount & vbCrLf & _
           "Iwato-joshoku columns were removed from the new sheets." & vbCrLf & _
           "The 3 menu blocks were set to equal width, filling the A3 portrait width." & vbCrLf & _
           "The source sheet was not changed. This file can be saved as usual (.xlsx is fine).", vbInformation

CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Exit Sub

ErrHandler:
    MsgBox "Stopped with an error." & vbCrLf & _
           "Step: " & gStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

' Delete all daily sheets (cleanup for a half-finished run).
' The original sheet is NOT touched by this sub.
Public Sub CleanupDailySheets()
    Dim wb As Workbook
    Dim i As Long
    Dim prefix As String
    prefix = CJ("65E5 5225") & "_"
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If UCase$(wb.Name) = "PERSONAL.XLSB" Then
        MsgBox "Open the workbook to clean first, then run again.", vbExclamation
        Exit Sub
    End If
    Application.DisplayAlerts = False
    For i = wb.Sheets.Count To 1 Step -1
        If wb.Sheets(i).Name Like prefix & "*" Then wb.Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True
    MsgBox "Daily sheets deleted.", vbInformation
End Sub

' Build a _B sheet from 3 row chunks of the source sheet:
' day header (4 rows) + meal header row + the <B> band.
Private Sub BuildBSheet(ByVal wb As Workbook, ByVal wsSrc As Worksheet, _
                        ByVal sheetName As String, ByVal startRow As Long, _
                        ByVal mealSrc As Long, ByVal bSrc As Long, _
                        ByVal bEndSrc As Long)
    Dim wsB As Worksheet
    Dim target As Long
    Dim totalRows As Long
    Set wsB = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsB.Name = sheetName
    CopyColumnWidths wsSrc, wsB

    wsSrc.Rows(startRow & ":" & startRow + 3).Copy Destination:=wsB.Rows(1)
    target = 5
    If mealSrc > 0 Then
        wsSrc.Rows(mealSrc).Copy Destination:=wsB.Rows(target)
        target = target + 1
    End If
    wsSrc.Rows(bSrc & ":" & bEndSrc).Copy Destination:=wsB.Rows(target)
    Application.CutCopyMode = False
    totalRows = (target - 1) + (bEndSrc - bSrc + 1)
    RemoveIwatoColumns wsB, totalRows
    EqualizeBlockWidths wsB, wsSrc, totalRows
    CopyStampToSheet wsSrc, wsB, startRow + 2
    ApplyPageSetup wsB, wsSrc, totalRows
End Sub

Private Sub CopyColumnWidths(ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet)
    Dim c As Long
    For c = 1 To 36
        wsDst.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
    Next c
End Sub

Private Sub ApplyPageSetup(ByVal ws As Worksheet, ByVal wsSrc As Worksheet, _
                           ByVal lastRowNum As Long)
    With ws.PageSetup
        .LeftHeader = wsSrc.PageSetup.LeftHeader
        .CenterHeader = wsSrc.PageSetup.CenterHeader
        .RightHeader = wsSrc.PageSetup.RightHeader
        .LeftFooter = wsSrc.PageSetup.LeftFooter
        .CenterFooter = wsSrc.PageSetup.CenterFooter
        .RightFooter = wsSrc.PageSetup.RightFooter
        .LeftMargin = wsSrc.PageSetup.LeftMargin
        .RightMargin = wsSrc.PageSetup.RightMargin
        .TopMargin = wsSrc.PageSetup.TopMargin
        .BottomMargin = wsSrc.PageSetup.BottomMargin
        .HeaderMargin = wsSrc.PageSetup.HeaderMargin
        .FooterMargin = wsSrc.PageSetup.FooterMargin
        .PrintArea = "$A$1:$AA$" & lastRowNum
        .PaperSize = xlPaperA3
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ws.ResetAllPageBreaks
End Sub

' Cell text that never raises an error on error-valued cells.
Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value) Then
        CellText = ""
    Else
        CellText = CStr(cell.Value)
    End If
End Function

' Build a Japanese string from space-separated hex code points.
Private Function CJ(ByVal hexList As String) As String
    Dim parts() As String
    Dim i As Long
    Dim s As String
    parts = Split(hexList, " ")
    For i = LBound(parts) To UBound(parts)
        s = s & ChrW$(CLng("&H" & parts(i)))
    Next i
    CJ = s
End Function

' True for the breakfast / lunch / dinner header texts.
Private Function IsMeal(ByVal txt As String) As Boolean
    IsMeal = (txt = CJ("671D 98DF")) Or _
             (txt = CJ("663C 98DF")) Or _
             (txt = CJ("5915 98DF"))
End Function

' Date header text -> month/day label.
' The header starts with the same year/month/day chars, so the
' search for the year char starts at position 5 to skip them.
Private Function DateLabel(ByVal txt As String) As String
    Dim p As Long, q As Long
    p = InStr(5, txt, CJ("5E74"))
    If p > 0 Then q = InStr(p, txt, CJ("65E5"))
    If p > 0 And q > p Then
        DateLabel = Mid$(txt, p + 1, q - p)
    Else
        DateLabel = txt
    End If
End Function

' V5: remove the Iwato-joshoku block (columns A-I) from a built
' daily sheet and shift the rest left. The field line, the date
' line and the meal headers live in column A, so their texts
' (and fonts) are saved first and written back into the new
' column A after the delete.
Private Sub RemoveIwatoColumns(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim saved As Collection
    Set saved = CaptureAColTexts(ws, lastRow)
    ws.Columns("A:I").Delete
    RestoreAColTexts ws, saved
End Sub

Private Function CaptureAColTexts(ByVal ws As Worksheet, ByVal lastRow As Long) As Collection
    Dim col As New Collection
    Dim r As Long
    Dim t As String
    Dim cell As Range
    For r = 1 To lastRow
        Set cell = ws.Cells(r, 1)
        t = CellText(cell)
        If Len(t) > 0 Then
            If Left$(t, 2) = CJ("73FE 5834") Or _
               Left$(t, 3) = CJ("5E74 6708 65E5") Or _
               IsMeal(t) Then
                col.Add Array(r, t, cell.Font.Name, cell.Font.Size, _
                              cell.Font.Bold, cell.Font.Color, _
                              (cell.Interior.Pattern <> xlNone), _
                              cell.Interior.Color, cell.HorizontalAlignment)
            End If
        End If
    Next r
    Set CaptureAColTexts = col
End Function

Private Sub RestoreAColTexts(ByVal ws As Worksheet, ByVal saved As Collection)
    Dim item As Variant
    Dim cell As Range
    For Each item In saved
        Set cell = ws.Cells(CLng(item(0)), 1)
        cell.Value = CStr(item(1))
        cell.Font.Name = CStr(item(2))
        cell.Font.Size = CDbl(item(3))
        cell.Font.Bold = CBool(item(4))
        cell.Font.Color = CLng(item(5))
        If CBool(item(6)) Then cell.Interior.Color = CLng(item(7))
        cell.HorizontalAlignment = CLng(item(8))
    Next item
End Sub

' V6: set the three menu blocks (columns 1-9, 10-18, 19-27
' after the Iwato columns are removed) to the SAME total width,
' widened so the printed table just fills the width of A3
' portrait. A day's rows are always taller than the printable
' height, so fit-to-page shrinks the sheet by
' (printable height / content height); the target width is
' widened by the inverse of that factor. For short sheets
' (like a _B sheet) that fit the height as they are, the
' target is simply the printable width. Column proportions
' inside each block are preserved. Widths are measured in
' points and corrected over a few passes, because setting
' ColumnWidth does not map 1:1 to points.
Private Sub EqualizeBlockWidths(ByVal ws As Worksheet, ByVal wsSrc As Worksheet, _
                                ByVal lastRow As Long)
    Dim printableW As Double, printableH As Double
    Dim contentH As Double, ratio As Double, target As Double
    Dim b As Long, c As Long, r As Long, iter As Long
    Dim blockW As Double, factor As Double
    printableW = Application.InchesToPoints(297# / 25.4) _
                 - wsSrc.PageSetup.LeftMargin - wsSrc.PageSetup.RightMargin
    printableH = Application.InchesToPoints(420# / 25.4) _
                 - wsSrc.PageSetup.TopMargin - wsSrc.PageSetup.BottomMargin
    If printableW <= 0 Or printableH <= 0 Then Exit Sub
    contentH = 0
    For r = 1 To lastRow
        contentH = contentH + ws.Rows(r).Height
    Next r
    ratio = contentH / printableH
    If ratio < 1 Then ratio = 1
    target = printableW * ratio / 3
    For b = 0 To 2
        For iter = 1 To 3
            blockW = 0
            For c = b * 9 + 1 To b * 9 + 9
                blockW = blockW + ws.Columns(c).Width
            Next c
            If blockW <= 0 Then Exit For
            factor = target / blockW
            If Abs(factor - 1) < 0.001 Then Exit For
            For c = b * 9 + 1 To b * 9 + 9
                ws.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth * factor
            Next c
        Next iter
    Next b
End Sub

' V5: copy the day's stamp picture from the source sheet onto a
' built daily sheet, aligned to the top-right corner of the
' table (columns A:AA after the Iwato columns are removed).
' A picture counts as the day's stamp when its top-left cell is
' in column AH (33) or later, and within 3 rows above the date
' line. Returns True when a stamp was copied.
Private Function CopyStampToSheet(ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet, _
                                  ByVal dateRow As Long) As Boolean
    Dim shp As Shape
    Dim newShp As Shape
    Dim rightEdge As Double
    CopyStampToSheet = False
    For Each shp In wsSrc.Shapes
        If shp.Type = msoPicture Then
            If shp.TopLeftCell.Column >= 33 And _
               shp.TopLeftCell.Row >= dateRow - 3 And _
               shp.TopLeftCell.Row <= dateRow Then
                wsDst.Activate
                shp.Copy
                wsDst.Paste
                Application.CutCopyMode = False
                Set newShp = wsDst.Shapes(wsDst.Shapes.Count)
                rightEdge = wsDst.Columns("AA").Left + wsDst.Columns("AA").Width
                newShp.Top = 0
                newShp.Left = rightEdge - newShp.Width
                If newShp.Left < 0 Then newShp.Left = 0
                gStampCount = gStampCount + 1
                CopyStampToSheet = True
                Exit Function
            End If
        End If
    Next shp
End Function
