# Business Questions and Answers
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
- Create a new sheet named 'Input Sheet'
- Build tables to capture personal information, education, residence, and employment
- Install a submit button and assign a VBA code to it
- The VBA code for the submit button is linked to the master data which automatically populate in other worksheets

Below is the VBA code for the input sheet newly created
   ```vbnet
    Attribute VB_Name = "NewRecruit2"
  Sub SubmitData()
    Dim wsInput As Worksheet
    Dim wsMaster As Worksheet
    Dim tblMaster As ListObject
    Dim tblRecruitment As ListObject
    Dim newRow As ListRow
    Dim newRecruitmentRow As ListRow
    Dim recruitmentDate As String
    Dim lastRowRecruitment As Long

    ' Set references to the sheets
    Set wsInput = ThisWorkbook.Sheets("InputSheet")
    Set wsMaster = ThisWorkbook.Sheets("EmployeeFile")
    
    ' Ensure the sheets were set correctly
    If wsInput Is Nothing Then
        MsgBox "InputSheet not found!", vbCritical
        Exit Sub
    End If
    If wsMaster Is Nothing Then
        MsgBox "EmployeeFile not found!", vbCritical
        Exit Sub
    End If
    
    ' Set reference to the structured table in EmployeeFile
    Set tblMaster = wsMaster.ListObjects("MasterData")
    
    ' Add a new row to the master table
    Set newRow = tblMaster.ListRows.Add(AlwaysInsert:=True)
    
    ' Copy data from InputSheet to EmployeeFile, ensuring correct column placement
    newRow.Range(1, 2).Value = wsInput.Range("B2").Value ' Employee Name (Column B)
    newRow.Range(1, 3).Value = wsInput.Range("B3").Value ' Last Name (Column C)
    newRow.Range(1, 4).Value = wsInput.Range("B4").Value ' Second Name (Column D)
    newRow.Range(1, 5).Value = wsInput.Range("B5").Value ' Initials (Column E)
    newRow.Range(1, 10).Value = wsInput.Range("B6").Value ' Racial Group Description (Column J)
    newRow.Range(1, 8).Value = wsInput.Range("B7").Value ' Gender Description (Column H)
    newRow.Range(1, 26).Value = wsInput.Range("B11").Value ' Highest Qualification (Column Z)
    newRow.Range(1, 24).Value = wsInput.Range("B15").Value ' Province (Column X)
    newRow.Range(1, 23).Value = wsInput.Range("B16").Value ' Postal Code (Column W)
    newRow.Range(1, 22).Value = wsInput.Range("B17").Value ' Town (Column V)
    newRow.Range(1, 21).Value = wsInput.Range("B18").Value ' District Municipality (Column U)
    newRow.Range(1, 20).Value = wsInput.Range("B19").Value ' Local Municipality (Column T)
    newRow.Range(1, 19).Value = wsInput.Range("B20").Value ' Street Name (Column S)
    newRow.Range(1, 18).Value = wsInput.Range("B21").Value ' Complex (Column R)
    newRow.Range(1, 17).Value = wsInput.Range("B22").Value ' No (Column Q)
    newRow.Range(1, 7).Value = wsInput.Range("B26").Value ' Core or Non Core (Column G)
    newRow.Range(1, 15).Value = wsInput.Range("B27").Value ' Employment Equity - Occupational Levels (Column O)
    newRow.Range(1, 14).Value = wsInput.Range("B28").Value ' Paterson grade (Column N)
    newRow.Range(1, 25).Value = wsInput.Range("B29").Value ' Employment Contract Status (Column Y)
    newRow.Range(1, 6).Value = wsInput.Range("B30").Value ' Job Title (Column F)
    
    ' Set the recruitment date
    recruitmentDate = Date
    
    ' Ensure the recruitment table exists and set reference
    On Error Resume Next
    Set tblRecruitment = wsInput.ListObjects("RecruitmentTable")
    On Error GoTo 0
    
    If tblRecruitment Is Nothing Then
        MsgBox "RecruitmentTable not found!", vbCritical
        Exit Sub
    End If
    
    ' Find the first empty row in the recruitment table
    lastRowRecruitment = tblRecruitment.ListRows.Count + 1 ' +1 to account for the header row
    
    ' Copy new entry to recruitment table with the recruitment date
    tblRecruitment.ListRows.Add.Range(1, 1).Value = recruitmentDate ' Date column
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 2).Value = wsInput.Range("B2").Value ' Employee Name
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 3).Value = wsInput.Range("B3").Value ' Last Name
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 4).Value = wsInput.Range("B4").Value ' Second Name
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 5).Value = wsInput.Range("B30").Value ' Job Title
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 6).Value = wsInput.Range("B26").Value ' Core or Non Core
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 7).Value = wsInput.Range("B7").Value ' Gender Description
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 8).Value = wsInput.Range("B6").Value ' Racial Group Description
    tblRecruitment.ListRows(lastRowRecruitment).Range(1, 9).Value = wsInput.Range("B29").Value ' Employment Contract Status
    
    ' Clear the input fields after submission
    wsInput.Range("B2:B7, B11, B15:B22, B26:B30").ClearContents
    
    ' Ensure the new row matches the table formatting
    newRow.Range.Font.Size = tblMaster.HeaderRowRange.Font.Size
    newRow.Range.Font.Name = tblMaster.HeaderRowRange.Font.Name
    
    ' Display a message box to confirm submission
    MsgBox "Data submitted successfully!", vbInformation
  End Sub



