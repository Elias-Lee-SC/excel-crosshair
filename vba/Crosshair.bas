Attribute VB_Name = "modCrosshair"
Option Explicit

Private Const CROSSHAIR_MARKER As String = "__CROSSHAIR_XLAM_RULE__"
Private Const CROSSHAIR_STATUS_PREFIX As String = "Crosshair: "
Private Const SETTINGS_APP As String = "CrosshairAddin"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const MENU_TAG As String = "CrosshairAddinMenu"
Private Const DEFAULT_BASE_COLOR As Long = 10872468
Private Const DEFAULT_PATTERN_COLOR As Long = 10872468
Private Const VISIBLE_RANGE_BUFFER_SCREENS As Long = 20
Private Const VISIBLE_RANGE_REFRESH_MARGIN_SCREENS As Long = 1
Private Const MIN_BUFFER_ROWS As Long = 480
Private Const MIN_BUFFER_COLUMNS As Long = 160
Private Const MIN_REFRESH_MARGIN_ROWS As Long = 20
Private Const MIN_REFRESH_MARGIN_COLUMNS As Long = 8
Private Const WATCH_INTERVAL_SECONDS As Long = 1

Private gEnabled As Boolean
Private gInitialized As Boolean
Private gUpdating As Boolean
Private gOwnsStatusBar As Boolean
Private gWatchScheduled As Boolean
Private gNextWatchTime As Date
Private gLastHighlightRange As Range
Private gLastHighlightTopRow As Long
Private gLastHighlightBottomRow As Long
Private gLastHighlightLeftCol As Long
Private gLastHighlightRightCol As Long
Private gLastSelectionKey As String
Private gBaseColor As Long
Private gPatternColor As Long

Public Sub Auto_Open()
    InitializeCrosshair
End Sub

Public Sub InitializeCrosshair()
    On Error GoTo Fail

    gEnabled = True
    gInitialized = True
    LoadColorSettings
    BuildCrosshairMenu

    ApplyToCurrentSelection
    ScheduleCrosshairWatcher
    Exit Sub

Fail:
    ShowStatus "Initialize failed: " & Err.Description
End Sub

Public Sub ToggleCrosshair()
    EnsureInitialized

    If gEnabled Then
        DisableCrosshair
    Else
        EnableCrosshair
    End If
End Sub

Public Sub EnableCrosshair()
    EnsureInitialized

    gEnabled = True
    ClearStatus
    ApplyToCurrentSelection
    ScheduleCrosshairWatcher
End Sub

Public Sub DisableCrosshair()
    EnsureInitialized

    gEnabled = False
    CancelCrosshairWatcher
    CleanupAllOpenWorkbooks
    ShowStatus "Disabled"
End Sub

Public Sub CleanupCrosshair()
    CancelCrosshairWatcher
    CleanupAllOpenWorkbooks
    ClearStatus
End Sub

Public Sub ChooseCrosshairColor()
    Dim choice As Variant
    Dim prompt As String
    Dim textValue As String

    EnsureInitialized
    LoadColorSettings

    prompt = "Choose crosshair color:" & vbCrLf & _
             "1  Green" & vbCrLf & _
             "2  Blue" & vbCrLf & _
             "3  Purple" & vbCrLf & _
             "4  Pink" & vbCrLf & _
             "5  Orange" & vbCrLf & _
             "6  Yellow" & vbCrLf & _
             "7  Red" & vbCrLf & _
             "8  Gray" & vbCrLf & _
             vbCrLf & _
             "Or enter a custom hex color." & vbCrLf & _
             vbCrLf & _
             "Current color: " & CurrentCrosshairColorText()

    choice = Application.InputBox(prompt, "Crosshair Color", "1", Type:=2)
    If choice = False Then Exit Sub

    textValue = Trim$(CStr(choice))
    If Len(textValue) = 0 Then Exit Sub

    Select Case LCase$(textValue)
        Case "1", "green"
            SetDefaultCrosshairGreen
        Case "2", "blue"
            SetCrosshairColor RGB(198, 237, 252)
        Case "3", "purple"
            SetCrosshairColor RGB(235, 208, 231)
        Case "4", "pink"
            SetCrosshairColor RGB(246, 233, 216)
        Case "5", "orange"
            SetCrosshairColor RGB(255, 228, 144)
        Case "6", "yellow"
            SetCrosshairColor RGB(255, 254, 199)
        Case "7", "red"
            SetCrosshairColor RGB(255, 199, 206)
        Case "8", "gray", "grey"
            SetCrosshairColor RGB(215, 216, 218)
        Case Else
            If Not SetCrosshairColorFromHex(textValue) Then
                ShowStatus "Color not recognized"
                Exit Sub
            End If
    End Select

    SaveColorSettings
    ApplyToCurrentSelection
    ScheduleCrosshairWatcher
End Sub

Public Sub ShowCrosshairColorPicker()
    ChooseCrosshairColor
End Sub

Public Sub ResetCrosshairColor()
    EnsureInitialized

    SetDefaultCrosshairGreen
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorGreen()
    SetDefaultCrosshairGreen
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorBlue()
    SetCrosshairColor RGB(198, 237, 252)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorPurple()
    SetCrosshairColor RGB(235, 208, 231)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorPink()
    SetCrosshairColor RGB(246, 233, 216)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorOrange()
    SetCrosshairColor RGB(255, 228, 144)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorYellow()
    SetCrosshairColor RGB(255, 254, 199)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorRed()
    SetCrosshairColor RGB(255, 199, 206)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub SetCrosshairColorGray()
    SetCrosshairColor RGB(215, 216, 218)
    SaveColorSettings
    ApplyToCurrentSelection
End Sub

Public Sub TerminateCrosshair()
    On Error Resume Next

    CancelCrosshairWatcher
    CleanupLastHighlight
    DeleteCrosshairMenu
    gInitialized = False
    gEnabled = False
    ClearStatus
End Sub

Public Sub RebuildCrosshairMenu()
    BuildCrosshairMenu
End Sub

Public Sub RemoveCrosshairMenu()
    DeleteCrosshairMenu
End Sub

Public Function CrosshairIsEnabled() As Boolean
    CrosshairIsEnabled = gEnabled
End Function

Public Sub Crosshair_WatchVisibleRange()
    gWatchScheduled = False

    On Error GoTo SafeExit
    If Not gEnabled Then Exit Sub
    If Not gUpdating Then RefreshCrosshairForVisibleRange

SafeExit:
    If gEnabled Then ScheduleCrosshairWatcher
End Sub

Public Sub Crosshair_HandleSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If gUpdating Then Exit Sub
    If Not gEnabled Then Exit Sub
    If Not (TypeOf Sh Is Worksheet) Then Exit Sub
    If Target Is Nothing Then Exit Sub

    ApplyCrosshairToSelection Sh, Target
End Sub

Public Sub Crosshair_HandleSheetActivate(ByVal Sh As Object)
    If gUpdating Then Exit Sub
    If Not gEnabled Then Exit Sub
    If Not (TypeOf Sh Is Worksheet) Then Exit Sub

    ApplyToCurrentSelection
End Sub

Public Sub Crosshair_HandleSheetDeactivate(ByVal Sh As Object)
    If gUpdating Then Exit Sub
    If Not (TypeOf Sh Is Worksheet) Then Exit Sub

    CleanupLastHighlightForWorksheet Sh
End Sub

Public Sub Crosshair_HandleWorkbookBeforeClose(ByVal Wb As Workbook)
    CleanupLastHighlightForWorkbook Wb
End Sub

Public Sub Crosshair_HandleWorkbookBeforeSave(ByVal Wb As Workbook)
    CleanupLastHighlightForWorkbook Wb
End Sub

Private Sub EnsureInitialized()
    If Not gInitialized Then
        gInitialized = True
    End If
End Sub

Private Sub ScheduleCrosshairWatcher()
    On Error Resume Next
    If Not gEnabled Then Exit Sub
    If gWatchScheduled Then Exit Sub

    Err.Clear
    gNextWatchTime = DateAdd("s", WATCH_INTERVAL_SECONDS, Now)
    Application.OnTime EarliestTime:=gNextWatchTime, _
                       Procedure:=CrosshairWatcherProcedureName(), _
                       Schedule:=True
    gWatchScheduled = (Err.Number = 0)
    Err.Clear
End Sub

Private Sub CancelCrosshairWatcher()
    On Error Resume Next
    If gWatchScheduled Then
        Application.OnTime EarliestTime:=gNextWatchTime, _
                           Procedure:=CrosshairWatcherProcedureName(), _
                           Schedule:=False
    End If

    gWatchScheduled = False
    Err.Clear
End Sub

Private Function CrosshairWatcherProcedureName() As String
    CrosshairWatcherProcedureName = "'" & ThisWorkbook.Name & "'!Crosshair_WatchVisibleRange"
End Function

Private Function CrosshairSelectionKey(ByVal target As Range) As String
    On Error GoTo Fail
    CrosshairSelectionKey = target.Address(RowAbsolute:=True, _
                                           ColumnAbsolute:=True, _
                                           ReferenceStyle:=xlA1, _
                                           External:=True)
    Exit Function

Fail:
    CrosshairSelectionKey = vbNullString
End Function

Private Function CanPreserveHighlight(ByVal ws As Worksheet, ByVal selectionKey As String) As Boolean
    On Error GoTo SafeExit
    If Len(selectionKey) = 0 Then Exit Function
    If selectionKey <> gLastSelectionKey Then Exit Function
    If gLastHighlightRange Is Nothing Then Exit Function
    If gLastHighlightTopRow = 0 Then Exit Function
    If Not (gLastHighlightRange.Worksheet Is ws) Then Exit Function

    CanPreserveHighlight = True

SafeExit:
End Function

Private Sub RefreshCrosshairForVisibleRange()
    Dim ws As Worksheet
    Dim target As Range

    On Error GoTo SafeExit
    If Not (TypeOf ActiveSheet Is Worksheet) Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet
    If ws.ProtectContents Then Exit Sub

    Set target = Selection
    If target Is Nothing Then Exit Sub
    If Not (target.Worksheet Is ws) Then Exit Sub

    If gLastHighlightRange Is Nothing Then
        ApplyCrosshairToSelection ws, target
    ElseIf Not (gLastHighlightRange.Worksheet Is ws) Then
        ApplyCrosshairToSelection ws, target
    ElseIf VisibleRangeNeedsRefresh(ws) Then
        ApplyCrosshairToSelection ws, target, True
    End If

SafeExit:
End Sub

Private Function VisibleRangeNeedsRefresh(ByVal ws As Worksheet) As Boolean
    Dim topRow As Long
    Dim bottomRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim visibleRows As Long
    Dim visibleColumns As Long
    Dim rowMargin As Long
    Dim columnMargin As Long

    If gLastHighlightTopRow = 0 Then
        VisibleRangeNeedsRefresh = True
        Exit Function
    End If

    If Not GetVisibleBounds(ws, topRow, bottomRow, leftCol, rightCol) Then Exit Function

    visibleRows = bottomRow - topRow + 1
    visibleColumns = rightCol - leftCol + 1
    rowMargin = MaxLong(MIN_REFRESH_MARGIN_ROWS, visibleRows * VISIBLE_RANGE_REFRESH_MARGIN_SCREENS)
    columnMargin = MaxLong(MIN_REFRESH_MARGIN_COLUMNS, visibleColumns * VISIBLE_RANGE_REFRESH_MARGIN_SCREENS)

    If topRow < gLastHighlightTopRow Or bottomRow > gLastHighlightBottomRow Then
        VisibleRangeNeedsRefresh = True
    ElseIf leftCol < gLastHighlightLeftCol Or rightCol > gLastHighlightRightCol Then
        VisibleRangeNeedsRefresh = True
    ElseIf gLastHighlightTopRow > 1 And topRow <= gLastHighlightTopRow + rowMargin Then
        VisibleRangeNeedsRefresh = True
    ElseIf gLastHighlightBottomRow < ws.Rows.Count And bottomRow >= gLastHighlightBottomRow - rowMargin Then
        VisibleRangeNeedsRefresh = True
    ElseIf gLastHighlightLeftCol > 1 And leftCol <= gLastHighlightLeftCol + columnMargin Then
        VisibleRangeNeedsRefresh = True
    ElseIf gLastHighlightRightCol < ws.Columns.Count And rightCol >= gLastHighlightRightCol - columnMargin Then
        VisibleRangeNeedsRefresh = True
    End If
End Function

Private Sub ApplyToCurrentSelection()
    On Error GoTo SafeExit

    If Not (TypeOf ActiveSheet Is Worksheet) Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    ApplyCrosshairToSelection ActiveSheet, Selection

SafeExit:
End Sub

Private Sub ApplyCrosshairToSelection(ByVal ws As Worksheet, ByVal target As Range, _
                                      Optional ByVal preserveExisting As Boolean = False)
    Dim wasSaved As Boolean
    Dim highlightRange As Range
    Dim formulaText As String
    Dim fc As FormatCondition
    Dim topRow As Long
    Dim bottomRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim selectionKey As String
    Dim canPreserve As Boolean

    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If Not (target.Worksheet Is ws) Then Exit Sub

    If ws.ProtectContents Then
        ShowStatus "Protected sheet skipped"
        Exit Sub
    End If

    wasSaved = WorkbookWasSaved(ws.Parent)
    gUpdating = True
    selectionKey = CrosshairSelectionKey(target)
    canPreserve = preserveExisting And CanPreserveHighlight(ws, selectionKey)

    If canPreserve Then
        Set highlightRange = BuildCrosshairExpansionRange(ws, target, topRow, bottomRow, leftCol, rightCol)
    Else
        CleanupLastHighlight
        Set highlightRange = BuildCrosshairRange(ws, target, topRow, bottomRow, leftCol, rightCol)
    End If

    If Not (highlightRange Is Nothing) Then
        formulaText = CrosshairFormula(target)
        Set fc = highlightRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaText)

        ApplyCrosshairStyle fc

        If canPreserve Then
            AddToRange gLastHighlightRange, highlightRange
            MergeHighlightBounds topRow, bottomRow, leftCol, rightCol
        Else
            Set gLastHighlightRange = highlightRange
            RememberHighlightBounds topRow, bottomRow, leftCol, rightCol
        End If
    End If
    gLastSelectionKey = selectionKey

    RestoreWorkbookSavedState ws.Parent, wasSaved
    ClearStatus

SafeExit:
    gUpdating = False
    If gEnabled Then ScheduleCrosshairWatcher
    Exit Sub

Fail:
    RestoreWorkbookSavedState ws.Parent, wasSaved
    ShowStatus "Update failed: " & Err.Description
    Resume SafeExit
End Sub

Private Sub ApplyCrosshairStyle(ByVal fc As FormatCondition)
    LoadColorSettings

    With fc
        .Interior.Pattern = xlSolid
        .Interior.Color = gBaseColor
        .StopIfTrue = False
        .SetFirstPriority
    End With
End Sub

Private Sub SetCrosshairColor(ByVal accentColor As Long)
    gBaseColor = accentColor
    gPatternColor = accentColor
End Sub

Private Sub SetDefaultCrosshairGreen()
    gBaseColor = DEFAULT_BASE_COLOR
    gPatternColor = DEFAULT_PATTERN_COLOR
End Sub

Private Function CurrentCrosshairColorText() As String
    Select Case gBaseColor
        Case DEFAULT_BASE_COLOR
            CurrentCrosshairColorText = "1 Green"
        Case RGB(198, 237, 252)
            CurrentCrosshairColorText = "2 Blue"
        Case RGB(235, 208, 231)
            CurrentCrosshairColorText = "3 Purple"
        Case RGB(246, 233, 216)
            CurrentCrosshairColorText = "4 Pink"
        Case RGB(255, 228, 144)
            CurrentCrosshairColorText = "5 Orange"
        Case RGB(255, 254, 199)
            CurrentCrosshairColorText = "6 Yellow"
        Case RGB(255, 199, 206)
            CurrentCrosshairColorText = "7 Red"
        Case RGB(215, 216, 218)
            CurrentCrosshairColorText = "8 Gray"
        Case Else
            CurrentCrosshairColorText = ColorToHexText(gBaseColor)
    End Select
End Function

Private Function ColorToHexText(ByVal colorValue As Long) As String
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    redPart = colorValue Mod 256
    greenPart = (colorValue \ 256) Mod 256
    bluePart = (colorValue \ 65536) Mod 256

    ColorToHexText = "#" & Right$("0" & Hex$(redPart), 2) & _
                     Right$("0" & Hex$(greenPart), 2) & _
                     Right$("0" & Hex$(bluePart), 2)
End Function

Private Function SetCrosshairColorFromHex(ByVal hexText As String) As Boolean
    Dim cleanHex As String
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    On Error GoTo Fail

    cleanHex = Trim$(hexText)
    If Left$(cleanHex, 1) = "#" Then cleanHex = Mid$(cleanHex, 2)
    If Left$(LCase$(cleanHex), 2) = "0x" Then cleanHex = Mid$(cleanHex, 3)
    If Len(cleanHex) <> 6 Then GoTo Fail

    redPart = CLng("&H" & Mid$(cleanHex, 1, 2))
    greenPart = CLng("&H" & Mid$(cleanHex, 3, 2))
    bluePart = CLng("&H" & Mid$(cleanHex, 5, 2))

    SetCrosshairColor RGB(redPart, greenPart, bluePart)
    SetCrosshairColorFromHex = True
    Exit Function

Fail:
    SetCrosshairColorFromHex = False
End Function

Private Sub LoadColorSettings()
    On Error GoTo Defaults

    If gBaseColor = 0 Then
        gBaseColor = CLng(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "BaseColor", CStr(DEFAULT_BASE_COLOR)))
    End If

    If gPatternColor = 0 Then
        gPatternColor = CLng(GetSetting(SETTINGS_APP, SETTINGS_SECTION, "PatternColor", CStr(DEFAULT_PATTERN_COLOR)))
    End If

    Exit Sub

Defaults:
    gBaseColor = DEFAULT_BASE_COLOR
    gPatternColor = DEFAULT_PATTERN_COLOR
End Sub

Private Sub SaveColorSettings()
    On Error Resume Next

    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "BaseColor", CStr(gBaseColor)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "PatternColor", CStr(gPatternColor)
End Sub

Private Sub BuildCrosshairMenu()
    Dim menuBar As CommandBar

    On Error GoTo SafeExit

    DeleteCrosshairMenu

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    AddCrosshairCommandButton menuBar, MenuTextEnable(), "EnableCrosshair", False
    AddCrosshairCommandButton menuBar, MenuTextDisable(), "DisableCrosshair", False
    AddCrosshairCommandButton menuBar, MenuTextChooseColor(), "ChooseCrosshairColor", False
    AddCrosshairCommandButton menuBar, MenuTextResetColor(), "ResetCrosshairColor", False
    AddCrosshairCommandButton menuBar, MenuTextCleanHighlight(), "CleanupCrosshair", False
    AddCrosshairCommandButton menuBar, MenuTextRebuildMenu(), "RebuildCrosshairMenu", False

SafeExit:
End Sub

Private Sub AddCrosshairCommandButton(ByVal menuBar As CommandBar, _
                                      ByVal captionText As String, _
                                      ByVal macroName As String, _
                                      ByVal beginGroupValue As Boolean)
    Dim button As CommandBarButton

    On Error Resume Next

    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With button
        .Caption = captionText
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
        .Style = msoButtonCaption
        .TooltipText = captionText
        .DescriptionText = captionText
        .BeginGroup = beginGroupValue
        .Tag = MENU_TAG & "." & macroName
    End With
End Sub

Private Sub DeleteCrosshairMenu()
    Dim menuBar As CommandBar
    Dim i As Long

    On Error Resume Next

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    For i = menuBar.Controls.Count To 1 Step -1
        If IsCrosshairCommand(menuBar.Controls(i)) Then
            menuBar.Controls(i).Delete
        End If
    Next i
End Sub

Private Function IsCrosshairCommand(ByVal control As CommandBarControl) As Boolean
    Dim captionText As String

    On Error Resume Next

    If control.Tag = MENU_TAG Then
        IsCrosshairCommand = True
        Exit Function
    End If

    captionText = control.Caption
    If Left$(captionText, Len(CrosshairCaption())) = CrosshairCaption() Then
        IsCrosshairCommand = True
        Exit Function
    End If

    If captionText = CrosshairCaption() Then
        IsCrosshairCommand = True
    End If
End Function

Private Function CrosshairCaption() As String
    CrosshairCaption = ChrW(&H5341) & ChrW(&H5B57) & ChrW(&H5149) & ChrW(&H6A19)
End Function

Private Function MenuTextEnable() As String
    MenuTextEnable = CrosshairCaption() & " " & ChrW(&H555F) & ChrW(&H7528)
End Function

Private Function MenuTextDisable() As String
    MenuTextDisable = CrosshairCaption() & " " & ChrW(&H505C) & ChrW(&H7528)
End Function

Private Function MenuTextChooseColor() As String
    MenuTextChooseColor = CrosshairCaption() & " " & ChrW(&H9078) & ChrW(&H64C7) & ChrW(&H984F) & ChrW(&H8272)
End Function

Private Function MenuTextResetColor() As String
    MenuTextResetColor = CrosshairCaption() & " " & ChrW(&H91CD) & ChrW(&H7F6E) & ChrW(&H984F) & ChrW(&H8272)
End Function

Private Function MenuTextCleanHighlight() As String
    MenuTextCleanHighlight = CrosshairCaption() & " " & ChrW(&H6E05) & ChrW(&H7406) & ChrW(&H9AD8) & ChrW(&H4EAE)
End Function

Private Function MenuTextRebuildMenu() As String
    MenuTextRebuildMenu = CrosshairCaption() & " " & ChrW(&H91CD) & ChrW(&H5EFA) & ChrW(&H6309) & ChrW(&H9215)
End Function

Private Function BuildCrosshairRange(ByVal ws As Worksheet, ByVal target As Range, _
                                     ByRef topRow As Long, ByRef bottomRow As Long, _
                                     ByRef leftCol As Long, ByRef rightCol As Long) As Range
    If Not GetVisibleBounds(ws, topRow, bottomRow, leftCol, rightCol) Then
        Exit Function
    End If
    ExpandVisibleBounds ws, topRow, bottomRow, leftCol, rightCol

    Set BuildCrosshairRange = BuildCrosshairRangeWithinBounds(ws, target, topRow, bottomRow, leftCol, rightCol)
End Function

Private Function BuildCrosshairExpansionRange(ByVal ws As Worksheet, ByVal target As Range, _
                                              ByRef topRow As Long, ByRef bottomRow As Long, _
                                              ByRef leftCol As Long, ByRef rightCol As Long) As Range
    Dim overlapTop As Long
    Dim overlapBottom As Long
    Dim result As Range

    If Not GetVisibleBounds(ws, topRow, bottomRow, leftCol, rightCol) Then
        Exit Function
    End If
    ExpandVisibleBounds ws, topRow, bottomRow, leftCol, rightCol

    If topRow < gLastHighlightTopRow Then
        AddToRange result, BuildCrosshairRangeWithinBounds(ws, target, topRow, gLastHighlightTopRow - 1, leftCol, rightCol)
    End If

    If bottomRow > gLastHighlightBottomRow Then
        AddToRange result, BuildCrosshairRangeWithinBounds(ws, target, gLastHighlightBottomRow + 1, bottomRow, leftCol, rightCol)
    End If

    overlapTop = MaxLong(topRow, gLastHighlightTopRow)
    overlapBottom = MinLong(bottomRow, gLastHighlightBottomRow)

    If overlapTop <= overlapBottom Then
        If leftCol < gLastHighlightLeftCol Then
            AddToRange result, BuildCrosshairRangeWithinBounds(ws, target, overlapTop, overlapBottom, leftCol, gLastHighlightLeftCol - 1)
        End If

        If rightCol > gLastHighlightRightCol Then
            AddToRange result, BuildCrosshairRangeWithinBounds(ws, target, overlapTop, overlapBottom, gLastHighlightRightCol + 1, rightCol)
        End If
    End If

    Set BuildCrosshairExpansionRange = result
End Function

Private Function BuildCrosshairRangeWithinBounds(ByVal ws As Worksheet, ByVal target As Range, _
                                                 ByVal topRow As Long, ByVal bottomRow As Long, _
                                                 ByVal leftCol As Long, ByVal rightCol As Long) As Range
    Dim area As Range
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long
    Dim result As Range

    If topRow > bottomRow Then Exit Function
    If leftCol > rightCol Then Exit Function

    For Each area In target.Areas
        rowStart = MaxLong(area.Row, topRow)
        rowEnd = MinLong(area.Row + area.Rows.Count - 1, bottomRow)
        colStart = MaxLong(area.Column, leftCol)
        colEnd = MinLong(area.Column + area.Columns.Count - 1, rightCol)

        If rowStart <= rowEnd Then
            AddRangeSkippingMerged result, ws.Range(ws.Cells(rowStart, leftCol), ws.Cells(rowEnd, rightCol))
        End If

        If colStart <= colEnd Then
            AddRangeSkippingMerged result, ws.Range(ws.Cells(topRow, colStart), ws.Cells(bottomRow, colEnd))
        End If
    Next area

    Set BuildCrosshairRangeWithinBounds = result
End Function

Private Function GetVisibleBounds(ByVal ws As Worksheet, _
                                  ByRef topRow As Long, ByRef bottomRow As Long, _
                                  ByRef leftCol As Long, ByRef rightCol As Long) As Boolean
    Dim visible As Range
    Dim area As Range

    On Error GoTo Fail
    Set visible = ActiveWindow.VisibleRange
    If visible Is Nothing Then Exit Function
    If Not (visible.Worksheet Is ws) Then Exit Function

    topRow = ws.Rows.Count
    bottomRow = 1
    leftCol = ws.Columns.Count
    rightCol = 1

    For Each area In visible.Areas
        topRow = MinLong(topRow, area.Row)
        bottomRow = MaxLong(bottomRow, area.Row + area.Rows.Count - 1)
        leftCol = MinLong(leftCol, area.Column)
        rightCol = MaxLong(rightCol, area.Column + area.Columns.Count - 1)
    Next area

    GetVisibleBounds = (topRow <= bottomRow And leftCol <= rightCol)
    Exit Function

Fail:
    GetVisibleBounds = False
End Function

Private Sub ExpandVisibleBounds(ByVal ws As Worksheet, _
                                ByRef topRow As Long, ByRef bottomRow As Long, _
                                ByRef leftCol As Long, ByRef rightCol As Long)
    Dim visibleRows As Long
    Dim visibleColumns As Long
    Dim rowPadding As Long
    Dim columnPadding As Long

    visibleRows = bottomRow - topRow + 1
    visibleColumns = rightCol - leftCol + 1
    rowPadding = MaxLong(MIN_BUFFER_ROWS, visibleRows * VISIBLE_RANGE_BUFFER_SCREENS)
    columnPadding = MaxLong(MIN_BUFFER_COLUMNS, visibleColumns * VISIBLE_RANGE_BUFFER_SCREENS)

    topRow = MaxLong(1, topRow - rowPadding)
    bottomRow = MinLong(ws.Rows.Count, bottomRow + rowPadding)
    leftCol = MaxLong(1, leftCol - columnPadding)
    rightCol = MinLong(ws.Columns.Count, rightCol + columnPadding)
End Sub

Private Sub RememberHighlightBounds(ByVal topRow As Long, ByVal bottomRow As Long, _
                                    ByVal leftCol As Long, ByVal rightCol As Long)
    gLastHighlightTopRow = topRow
    gLastHighlightBottomRow = bottomRow
    gLastHighlightLeftCol = leftCol
    gLastHighlightRightCol = rightCol
End Sub

Private Sub MergeHighlightBounds(ByVal topRow As Long, ByVal bottomRow As Long, _
                                 ByVal leftCol As Long, ByVal rightCol As Long)
    If gLastHighlightTopRow = 0 Then
        RememberHighlightBounds topRow, bottomRow, leftCol, rightCol
    Else
        gLastHighlightTopRow = MinLong(gLastHighlightTopRow, topRow)
        gLastHighlightBottomRow = MaxLong(gLastHighlightBottomRow, bottomRow)
        gLastHighlightLeftCol = MinLong(gLastHighlightLeftCol, leftCol)
        gLastHighlightRightCol = MaxLong(gLastHighlightRightCol, rightCol)
    End If
End Sub

Private Sub ResetHighlightBounds()
    gLastHighlightTopRow = 0
    gLastHighlightBottomRow = 0
    gLastHighlightLeftCol = 0
    gLastHighlightRightCol = 0
    gLastSelectionKey = vbNullString
End Sub

Private Function CrosshairFormula(ByVal target As Range) As String
    Dim exclusionFormula As String

    exclusionFormula = SelectionExclusionFormula(target)
    CrosshairFormula = "=AND(LEN(""" & CROSSHAIR_MARKER & """)>0," & _
                       "NOT(" & exclusionFormula & "))"
End Function

Private Function SelectionExclusionFormula(ByVal target As Range) As String
    Dim area As Range
    Dim expressionList As String
    Dim expressionPart As String
    Dim areaCount As Long

    On Error GoTo Fallback

    For Each area In target.Areas
        expressionPart = RangeExclusionFormula(area)
        If Len(expressionPart) > 0 Then
            If Len(expressionList) > 0 Then expressionList = expressionList & ","
            expressionList = expressionList & expressionPart
            areaCount = areaCount + 1
        End If
    Next area

    If Len(expressionList) = 0 Then GoTo Fallback

    If areaCount = 1 Then
        SelectionExclusionFormula = expressionList
    Else
        SelectionExclusionFormula = "OR(" & expressionList & ")"
    End If
    Exit Function

Fallback:
    SelectionExclusionFormula = ActiveCellExclusionFormula()
End Function

Private Function RangeExclusionFormula(ByVal selectedArea As Range) As String
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rowTest As String
    Dim columnTest As String

    firstRow = selectedArea.Row
    lastRow = selectedArea.Row + selectedArea.Rows.Count - 1
    firstCol = selectedArea.Column
    lastCol = selectedArea.Column + selectedArea.Columns.Count - 1

    If firstRow = lastRow Then
        rowTest = "ROW()=" & CStr(firstRow)
    Else
        rowTest = "ROW()>=" & CStr(firstRow) & ",ROW()<=" & CStr(lastRow)
    End If

    If firstCol = lastCol Then
        columnTest = "COLUMN()=" & CStr(firstCol)
    Else
        columnTest = "COLUMN()>=" & CStr(firstCol) & ",COLUMN()<=" & CStr(lastCol)
    End If

    RangeExclusionFormula = "AND(" & rowTest & "," & columnTest & ")"
End Function

Private Function ActiveCellExclusionFormula() As String
    ActiveCellExclusionFormula = "AND(ROW()=" & CStr(ActiveCell.Row) & _
                                 ",COLUMN()=" & CStr(ActiveCell.Column) & ")"
End Function

Private Sub AddToRange(ByRef baseRange As Range, ByVal rangeToAdd As Range)
    If rangeToAdd Is Nothing Then Exit Sub

    If baseRange Is Nothing Then
        Set baseRange = rangeToAdd
    Else
        Set baseRange = Union(baseRange, rangeToAdd)
    End If
End Sub

Private Sub AddRangeSkippingMerged(ByRef baseRange As Range, ByVal rangeToAdd As Range)
    Dim area As Range
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim segmentStart As Long

    On Error GoTo SafeExit
    If rangeToAdd Is Nothing Then Exit Sub

    For Each area In rangeToAdd.Areas
        Set ws = area.Worksheet
        firstCol = area.Column
        lastCol = area.Column + area.Columns.Count - 1

        For rowIndex = area.Row To area.Row + area.Rows.Count - 1
            segmentStart = 0

            For colIndex = firstCol To lastCol
                If ws.Cells(rowIndex, colIndex).MergeCells Then
                    If segmentStart > 0 Then
                        AddToRange baseRange, ws.Range(ws.Cells(rowIndex, segmentStart), ws.Cells(rowIndex, colIndex - 1))
                        segmentStart = 0
                    End If
                ElseIf segmentStart = 0 Then
                    segmentStart = colIndex
                End If
            Next colIndex

            If segmentStart > 0 Then
                AddToRange baseRange, ws.Range(ws.Cells(rowIndex, segmentStart), ws.Cells(rowIndex, lastCol))
            End If
        Next rowIndex
    Next area

SafeExit:
End Sub

Private Sub CleanupAllOpenWorkbooks()
    Dim wb As Workbook

    On Error Resume Next
    For Each wb In Application.Workbooks
        CleanupWorkbook wb
    Next wb
End Sub

Private Sub CleanupWorkbook(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim wasSaved As Boolean

    On Error GoTo SafeExit
    If wb Is Nothing Then Exit Sub

    wasSaved = WorkbookWasSaved(wb)
    gUpdating = True
    CleanupLastHighlightForWorkbook wb

    For Each ws In wb.Worksheets
        If Not ws.ProtectContents Then
            DeleteCrosshairRules ws
        End If
    Next ws

    RestoreWorkbookSavedState wb, wasSaved

SafeExit:
    gUpdating = False
End Sub

Private Sub CleanupWorksheet(ByVal ws As Worksheet)
    Dim wasSaved As Boolean

    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then Exit Sub

    wasSaved = WorkbookWasSaved(ws.Parent)
    gUpdating = True
    CleanupLastHighlightForWorksheet ws
    DeleteCrosshairRules ws
    RestoreWorkbookSavedState ws.Parent, wasSaved

SafeExit:
    gUpdating = False
End Sub

Private Sub CleanupLastHighlight()
    On Error Resume Next

    If Not (gLastHighlightRange Is Nothing) Then
        DeleteCrosshairRulesInRange gLastHighlightRange
        Set gLastHighlightRange = Nothing
    End If
    ResetHighlightBounds
End Sub

Private Sub CleanupLastHighlightForWorkbook(ByVal wb As Workbook)
    On Error Resume Next

    If gLastHighlightRange Is Nothing Then Exit Sub
    If gLastHighlightRange.Worksheet.Parent Is wb Then
        CleanupLastHighlight
    End If
End Sub

Private Sub CleanupLastHighlightForWorksheet(ByVal ws As Worksheet)
    On Error Resume Next

    If gLastHighlightRange Is Nothing Then Exit Sub
    If gLastHighlightRange.Worksheet Is ws Then
        CleanupLastHighlight
    End If
End Sub

Private Sub DeleteCrosshairRules(ByVal ws As Worksheet)
    Dim i As Long
    Dim formulaText As String
    Dim fc As Object

    On Error Resume Next

    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions.Item(i)
        formulaText = vbNullString
        formulaText = CStr(fc.Formula1)

        If InStr(1, formulaText, CROSSHAIR_MARKER, vbTextCompare) > 0 Then
            fc.Delete
        End If
    Next i
End Sub

Private Sub DeleteCrosshairRulesInRange(ByVal targetRange As Range)
    Dim i As Long
    Dim formulaText As String
    Dim fc As Object

    On Error Resume Next

    For i = targetRange.FormatConditions.Count To 1 Step -1
        Set fc = targetRange.FormatConditions.Item(i)
        formulaText = vbNullString
        formulaText = CStr(fc.Formula1)

        If InStr(1, formulaText, CROSSHAIR_MARKER, vbTextCompare) > 0 Then
            fc.Delete
        End If
    Next i
End Sub

Private Function WorkbookWasSaved(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    WorkbookWasSaved = wb.Saved
End Function

Private Sub RestoreWorkbookSavedState(ByVal wb As Workbook, ByVal wasSaved As Boolean)
    On Error Resume Next
    If wasSaved Then wb.Saved = True
End Sub

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        MinLong = a
    Else
        MinLong = b
    End If
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        MaxLong = a
    Else
        MaxLong = b
    End If
End Function

Private Sub ShowStatus(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = CROSSHAIR_STATUS_PREFIX & message
    gOwnsStatusBar = True
End Sub

Private Sub ClearStatus()
    On Error Resume Next
    If gOwnsStatusBar Then
        Application.StatusBar = False
        gOwnsStatusBar = False
    End If
End Sub
