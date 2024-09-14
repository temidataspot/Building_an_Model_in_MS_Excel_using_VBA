# Business Questions
The purpose of this analysis is to understand and answer some business questions regarding existing employees 
and new recruits, automate tools and gain insight into the raw dataset provided.

In this analysis, we focused on automating the number of employees across all sheets and automating the cost
of maintaining programs organized by GeoMine.

This analysis was carried out on the 23rd of May, 2024

# 1. How can we consolidate all sheets to reduce the number of sheets in the workbook?

**Steps:**
  - Identify related sheets that be consolidated
  - Using VBA to automate partial consolidation
    ```Sub ConsolidateSpecificSheets()
    Dim ws As Worksheet
    Dim wsMaster As Worksheet
    Dim wsList As Variant
    Dim rng As Range
    Dim lastRow As Long
    Dim i As Integer
    
    **Define the list of sheets to consolidate**
    wsList = Array("Sheet1", "Sheet2", "Sheet3") ' Modify this with your sheet names
    
    **Add a new sheet for consolidation**
    Set wsMaster = Sheets.Add
    wsMaster.Name = "Consolidated"
    
    **Loop through the list of specified sheets**
    For i = LBound(wsList) To UBound(wsList)
        Set ws = Sheets(wsList(i))
        lastRow = wsMaster.Cells(Rows.Count, 1).End(xlUp).Row + 1
        Set rng = ws.UsedRange
        rng.Copy wsMaster.Cells(lastRow, 1)
    Next i
    
    MsgBox "Specific sheets consolidated!"
    End Sub
