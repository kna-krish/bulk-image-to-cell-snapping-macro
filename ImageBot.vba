#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub InteractiveMasterImageBotWithLogs()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim cell As Range
    Dim pic As Shape
    Dim initialCount As Long
    Dim foundShape As Boolean
    
    ' Variables for user settings
    Dim startRowInput As String, endRowInput As String
    Dim columnInput As String
    Dim startRow As Long, endRow As Long
    Dim targetColumn As Long
    Dim targetRowHeight As Double, targetColumnWidth As Double
    Dim padding As Double
    
    ' Variables for logging mechanics
    Dim logText As String
    Dim logPath As String
    Dim fileNum As Integer
    Dim timeStamp As String
    
    Set ws = ActiveSheet
    padding = 4 ' Cell inner margin padding size
    
    ' --------------------------------------------------------------
    '   INTERACTIVE POP-UP PROMPTS FOR USER INPUT
    ' --------------------------------------------------------------
    ' 1. Get Starting Row
    startRowInput = InputBox("Enter the STARTING row number:", "Image Bot Configuration", "2")
    If startRowInput = "" Then Exit Sub ' User pressed Cancel
    startRow = Val(startRowInput)
    
    ' 2. Get Ending Row
    endRowInput = InputBox("Enter the ENDING row number:", "Image Bot Configuration", "20")
    If endRowInput = "" Then Exit Sub
    endRow = Val(endRowInput)
    
    ' 3. Get Target Column Letter
    columnInput = InputBox("Enter the Target COLUMN letter (e.g., A, B, C):", "Image Bot Configuration", "A")
    If columnInput = "" Then Exit Sub
    
    ' Convert column letter to a functional number index (A=1, B=2, etc.)
    On Error Resume Next
    targetColumn = ws.Columns(columnInput).Column
    On Error GoTo 0
    If targetColumn = 0 Then
        MsgBox "Invalid column letter entered! Operation cancelled.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' 4. Get Custom Dimensions
    Dim heightInput As String, widthInput As String
    heightInput = InputBox("Enter desired ROW HEIGHT:", "Cell Dimensions", "92")
    If heightInput = "" Then Exit Sub
    targetRowHeight = Val(heightInput)
    
    widthInput = InputBox("Enter desired COLUMN WIDTH:", "Cell Dimensions", "11.55")
    If widthInput = "" Then Exit Sub
    targetColumnWidth = Val(widthInput)
    ' --------------------------------------------------------------
    
    ' Initialize the Logging System String
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    logPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\excel_centering_debug_" & timeStamp & ".txt"
    
    logText = "=== MASTER IMAGE BOT LOG: " & Now & " ===" & vbCrLf
    logText = logText & "Excel Version: " & Application.Version & " | OS: " & Application.OperatingSystem & vbCrLf
    logText = logText & "Target Range: Column " & UCase(columnInput) & " (Rows " & startRow & " to " & endRow & ")" & vbCrLf
    logText = logText & "Applied Dimensions: Height " & targetRowHeight & " | Width " & targetColumnWidth & vbCrLf
    logText = logText & "--------------------------------------------------------" & vbCrLf & vbCrLf
    
    ' Clear active filters to ensure target rows are completely visible
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
    
    ' Pre-size the selected target cell range boundaries completely
    ws.Range(ws.Cells(startRow, targetColumn), ws.Cells(endRow, targetColumn)).RowHeight = targetRowHeight
    ws.Range(ws.Cells(startRow, targetColumn), ws.Cells(endRow, targetColumn)).ColumnWidth = targetColumnWidth
    
    ' Loop through every cell sequentially based on inputs
    For rowNum = startRow To endRow
        Set cell = ws.Cells(rowNum, targetColumn)
        foundShape = False
        logText = logText & "Processing Row " & rowNum & "..." & vbCrLf
        
        cell.Select
        DoEvents
        Sleep 150
        
        initialCount = ws.Shapes.Count
        
        ' Run the clipboard bypass extraction method
        On Error Resume Next
        cell.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        ws.Paste Destination:=cell
        Application.CutCopyMode = False
        On Error GoTo 0
        
        DoEvents
        Sleep 250
        
        ' Scan for the newly pasted floating shape object
        If ws.Shapes.Count > initialCount Then
            Set pic = ws.Shapes(ws.Shapes.Count)
            foundShape = True
        Else
            ' Fallback positioning search rule
            For Each pic In ws.Shapes
                If pic.TopLeftCell.Row = rowNum And pic.TopLeftCell.Column = targetColumn Then
                    Set pic = pic
                    foundShape = True
                    Exit For
                End If
            Next pic
        End If
        
        ' Perform fitting calculations and apply centered alignment coordinates
        If foundShape Then
            On Error Resume Next
            pic.LockAspectRatio = msoTrue
            
            If (cell.Width - (padding * 2)) / pic.Width < (cell.Height - (padding * 2)) / pic.Height Then
                pic.Width = cell.Width - (padding * 2)
            Else
                pic.Height = cell.Height - (padding * 2)
            End If
            
            pic.Left = cell.Left + ((cell.Width - pic.Width) / 2)
            pic.Top = cell.Top + ((cell.Height - pic.Height) / 2)
            
            ' Bind configuration to scale and move dynamically with grid cells
            pic.Placement = xlMoveAndSize
            
            ' Clear the original background cell image value format
            cell.Value = ""
            logText = logText & "  [SUCCESS] Extracted, centered, and sized. Name: " & pic.Name & vbCrLf
            On Error GoTo 0
        Else
            logText = logText & "  [FAIL] No image object could be captured for this row." & vbCrLf
        End If
        
        logText = logText & "--------------------------------------------------------" & vbCrLf
        DoEvents
    Next rowNum
    
    ' Force Excel to redraw the graphics screen layout using the view-switch trick
    ActiveWindow.View = xlPageBreakPreview
    DoEvents
    Sleep 200
    ActiveWindow.View = xlNormalView
    DoEvents
    
    ' Write the log summary file onto the user's Desktop
    On Error Resume Next
    fileNum = FreeFile
    Open logPath For Output As #fileNum
    Print #fileNum, logText
    Close #fileNum
    On Error GoTo 0
    
    MsgBox "Automation complete!" & vbCrLf & vbCrLf & _
           "Images processed: Rows " & startRow & " to " & endRow & vbCrLf & _
           "Diagnostic log saved to Desktop as: " & vbCrLf & _
           "excel_centering_debug_" & timeStamp & ".txt", vbInformation, "Success"
End Sub
