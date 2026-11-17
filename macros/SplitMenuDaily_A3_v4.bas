Attribute VB_Name = "DailyMenuPrintV4"
Option Explicit

'================================================================
' SplitMenuDaily_A3_V4
' Split a stacked daily-menu sheet into one sheet per day.
' Each daily sheet prints on ONE A3 portrait page (columns A:AJ).
'
' V4 change: the target is now the ACTIVE workbook and its
' active sheet (previously it was fixed to the workbook that
' contains the macro). This lets the macro live on the PC
' itself - e.g. in the Personal Macro Workbook (PERSONAL.XLSB) -
' and process any menu workbook / sheet you open and select.
'
' V3 change: instead of copying the WHOLE source sheet once per
' day and deleting thousands of rows (heavy, could fail halfway
' in Excel), each daily sheet is built by copying ONLY that
' day's rows (about 100 rows) onto a fresh sheet. Much lighter.
' Row heights, formats and merged cells come with the row copy;
' column widths and header/footer are copied explicitly.
' Pictures (approval stamps) are never copied to daily sheets.
'
' If it still stops, the error message now shows WHICH STEP
' failed (and the day being processed). Tell the author that
' step name and the error text - that pinpoints the cause.
'
' This file is ASCII only, so it never garbles (mojibake).
' Japanese strings for matching cell text are built from hex
' code points by the CJ() helper:
'   date header prefix : 5E74 6708 65E5 FF1A
'   daily sheet prefix : 65E5 5225 (+ "_")
'   breakfast/lunch/dinner : 671D 98DF / 663C 98DF / 5915 98DF
'================================================================

Private gStep As String

Public Sub SplitMenuDaily_A3_V4()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, i As Long, n As Long
    Dim dateRows() As Long
    Dim dateLabels() As String
    Dim shp As Shape
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

    '--- Delete stamp pictures on the source sheet ---
    ' A picture counts as a stamp when its top-left cell is in
    ' column AH (33) or later, and within 3 rows above the date line.
    gStep = "delete stamp pictures"
    For i = wsSrc.Shapes.Count To 1 Step -1
        Set shp = wsSrc.Shapes(i)
        If shp.Type = msoPicture Then
            If IsStampPosition(shp, dateRows, n) Then shp.Delete
        End If
    Next i

    '--- Build one sheet per day ---
    selectCount = 0
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

        ApplyPageSetup ws, wsSrc, blockLen
    Next i

    gStep = "finish"
    wb.Sheets(prefix & dateLabels(1)).Activate
    MsgBox "Done. Daily sheets created in '" & wb.Name & "': " & n & vbCrLf & _
           "Select-menu _B sheets: " & selectCount & vbCrLf & _
           "The macro stays on this PC, so this file can be saved as usual (.xlsx is fine).", vbInformation

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
        .PrintArea = "$A$1:$AJ$" & lastRowNum
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

Private Function IsStampPosition(ByVal shp As Shape, dateRows() As Long, ByVal n As Long) As Boolean
    Dim i As Long
    For i = 1 To n
        If shp.TopLeftCell.Column >= 33 And _
           shp.TopLeftCell.Row >= dateRows(i) - 3 And _
           shp.TopLeftCell.Row <= dateRows(i) Then
            IsStampPosition = True
            Exit Function
        End If
    Next i
End Function
