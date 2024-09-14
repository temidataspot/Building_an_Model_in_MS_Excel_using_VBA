# Business Questions
The purpose of this analysis is to understand and answer some business questions regarding existing employees 
and new recruits, automate tools and gain insight into the raw dataset provided.

In this analysis, we focused on automating the number of employees across all sheets and automating the cost
of maintaining programs organized by GeoMine.

This analysis was carried out on the 23rd of May, 2024

# 1. How can we consolidate all sheets to reduce the number of sheets in the workbook?

**Steps:**
  - Survey sheets to ensure data completion, and clean data by standardizing data types and formats
  - Fill blanks as needed, trim spaces
  - identify related sheets that can be consolidated
  - Using VBA to automate partial consolidation
    ```vbnet
    Sub ConsolidateSpecificSheets()
      Dim ws As Worksheet
      Dim wsMaster As Worksheet
      Dim wsList As Variant
      Dim rng As Range
      Dim lastRow As Long
      Dim i As Integer
    
    ' Define the list of sheets to consolidate
    wsList = Array("Sheet1", "Sheet2", "Sheet3") ' Modify this with your sheet names
    
    ' Add a new sheet for consolidation
    Set wsMaster = Sheets.Add
    wsMaster.Name = "Consolidated"
    
    ' Loop through the list of specified sheets
    For i = LBound(wsList) To UBound(wsList)
        Set ws = Sheets(wsList(i))
        lastRow = wsMaster.Cells(Rows.Count, 1).End(xlUp).Row + 1
        Set rng = ws.UsedRange
        rng.Copy wsMaster.Cells(lastRow, 1)
    Next i
    
    MsgBox "Specific sheets consolidated!"
    End Sub

**Output**
The number of sheets was reduced from 60 to 32 aiding easy referencing and analysis

# 2. How do we link the different employee data in the other sheets to the master data (Sheet 0) such that any updates to either the master data is reflected in the working sheets and vice versa?

**Steps:**
- Ensure data types and format are consistent across all sheets
- Use VLOOKUP to reference and link the sheets together. Relevant columns in the master data are matched with the worksheets
  ```excel
  =VLOOKUP(A2, Sheet2!$A$2:$B$100, 2, FALSE)

| Employee ID | Name        | Amount                                       |
|-------------|-------------|----------------------------------------------|
| 1001        | Alice Smith | =VLOOKUP(A2, Sheet2!$A$2:$B$100, 2, FALSE)   |
| 1002        | Bob Johnson | =VLOOKUP(A3, Sheet2!$A$2:$B$100, 2, FALSE)   |

**Output**

| Employee ID | Name        | Amount    |
|-------------|-------------|-----------|
| 1001        | Alice Smith | $55,000   |
| 1002        | Bob Johnson | $60,000   |

# 3. How can we create an input sheet that captures the information of new recruits and adds them to the master data? This new data also automatically populates the rest of the working sheet.

**Steps:**
