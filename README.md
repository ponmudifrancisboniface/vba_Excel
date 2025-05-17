# vba_Excel
Consolidating of Multiple sheet
Option Explicit
Sub condmulwb_arraymethod()
    Dim WB As Workbook
    Dim NWB As Workbook
    Dim LastRow As Long
    Dim LastCol As Long
    Dim FileToOpen As Variant
    Dim NWBSh As Worksheet
    Dim SrcData As Variant
    Dim i As Integer
    Dim TargetRow As Long
    Dim FindResult As Range
    
    FileToOpen = Application.GetOpenFilename(Title:="Select The Files To Be Consolidated", filefilter:="Excel Files (*.xlsx), *.xlsx", MultiSelect:=True)
    
    If IsArray(FileToOpen) = False Then
        MsgBox "No files selected!", vbExclamation, "Error"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set NWB = Application.Workbooks.Add
    Set NWBSh = NWB.Sheets(1)
    TargetRow = 1
    
    For i = LBound(FileToOpen) To UBound(FileToOpen)
        Set WB = Application.Workbooks.Open(FileToOpen(i), ReadOnly:=True)
        
        With WB.ActiveSheet
            Set FindResult = .Cells.Find("*", , , , xlByRows, xlPrevious)
            If Not FindResult Is Nothing Then
                LastRow = FindResult.Row
                LastCol = .Cells.Find("*", , , , xlByColumns, xlPrevious).Column
                
                ' Load data into array
                If i = LBound(FileToOpen) Then
                    ' First file: Include headers
                    SrcData = .Range(.Cells(1, 1), .Cells(LastRow, LastCol)).Value
                Else
                    ' Next files: Skip headers
                    SrcData = .Range(.Cells(2, 1), .Cells(LastRow, LastCol)).Value
                End If
                
                ' Write array to destination
                NWBSh.Cells(TargetRow, 1).Resize(UBound(SrcData, 1), UBound(SrcData, 2)).Value = SrcData
                TargetRow = NWBSh.Cells(NWBSh.Rows.Count, 1).End(xlUp).Row + 1
            Else
                MsgBox "No data found in file: " & WB.Name, vbExclamation, "Empty File"
            End If
        End With
        
        WB.Close SaveChanges:=False
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Consolidation complete (Array Method)!", vbInformation, "Done"
End Sub

