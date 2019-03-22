Attribute VB_Name = "MappingModule"
'Macro for mapping SAP GL-accounts and cost centers to HFM accounts of HFM-forms: P&L, COGS, SG&A, Category costs
Option Explicit 'require declarations for variables
Option Base 1 'first number in arrays is 1
Sub MappingMacro()
Dim OriginalCalcMode As Long
Dim DataTable As Worksheet, MappingTable As Worksheet
Dim DataTableLastRow As Long, MappingTableLastRow As Long
Dim HfmFormsArray(1 To 5, 1 To 3) As String
Dim HfmFormNumber As Long, DataTableCurrentRow As Long, MappingTableCurrentRow As Long 'counters of cycles
Dim ProfitCenterTypeCurrent As Long, ProfitCenterTypeLeftBorder As Long, ProfitCenterTypeRightBorder As Long 'pc types (1st number of cost center: 1 - make, 2 - psd, 3 - sell, 4 - shared services)
Dim GLaccountCurrent As Long, GLaccountLeftBorder As Long, GLaccountRightBorder As Long 'SAP GL-account full number
Dim CostCenterTypeCurrent As Long, CostCenterTypeLeftBorder As Long, CostCenterTypeRightBorder As Long 'cost center type (6th and 7th numbers of cost center code)
Dim HfmFormOnMappingTable As String 'HFM-form names of mapping table
Dim HfmFormAmount As Long 'amount of HFM forms

'settings for acceleration:
OriginalCalcMode = Application.Calculation 'save current status of calculation
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'initialization of the objects:
Set DataTable = Application.ThisWorkbook.Sheets("DataToMap") 'Application.ThisWorkbook is the property of object Application. It returns the current workbook where the macro code is running. It returned = MappingMacro.xlsm
Set MappingTable = Application.ThisWorkbook.Sheets("ZFM_ISRL_CUSTOM")

'initialization of the array:
HfmFormsArray(1, 1) = "INCOME STATEMENT": HfmFormsArray(1, 2) = "E":: HfmFormsArray(1, 3) = "F"
HfmFormsArray(2, 1) = "COST OF GOODS SOLD": HfmFormsArray(2, 2) = "G": HfmFormsArray(2, 3) = "H"
HfmFormsArray(3, 1) = "SPECIFICATION OVERHEAD QUARTERLY": HfmFormsArray(3, 2) = "I": HfmFormsArray(3, 3) = "J"
HfmFormsArray(4, 1) = "PERSONNEL COST ACTUAL QUARTERLY": HfmFormsArray(4, 2) = "K": HfmFormsArray(4, 3) = "L"
HfmFormsArray(5, 1) = "SPECIFICATION OF COSTS CATEGORIES YEAR": HfmFormsArray(5, 2) = "M": HfmFormsArray(5, 3) = "N"
HfmFormAmount = UBound(HfmFormsArray, 1) 'calculation of size of array

'search the last rows of the worksheet:
DataTableLastRow = DataTable.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
MappingTableLastRow = MappingTable.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
'Comments to the code above:
'DataTable and MapingTable are object variables that returns the worksheet objects that represents the worksheets "DataToMap" and "ZFM_ISRL_CUSTOM" respectively
'DataTable.Range() and MappingTable.Range() are the properties of specified worksheet objects (here - worksheets("DataToMap") and worksheets("ZFM_ISRL_CUSTOM"))
'The property <worksheet object>.Range() requires a parameter (single cell, single column, single row, range of cells or multiple ranges)
'ActiveSheet.Rows is the property of active worksheet object (other variants of syntax: Application.Rows, only Rows)
'The property ActiveSheet.Rows returns the range object that represents all the rows of the active worksheet (here - worksheets("DataToMap"))
'ActiveSheet.Rows.Count is the property of the range object ActiveSheet.Rows. It represents the total amount of objects in the collection. ActiveSheet.Rows.Count = 1048576 for xlsm Excel file format.
'The parameter for the property DataTable.Range() = "A1048576" (last most cell in the column "A")
'DataTable.Range("A1048576") is the range object that represents the cell A1048576
'DataTable.Range("A1048576").End() is the property of range object DataTable.Range("A1048576").
'The property <range object>.End() requires a parameter of the direction. The parameter xlUp is an equivalent of buttons END+Up_Arrow
'The property DataTable.Range("A1048576").End(xlUp) finds the next nearest filled cell in specified column (here - "A") when searching up from specified row (here - beginging from 1048576). The result is the single cell (range object)
'<found cell>.Row is the property of range object (singe found cell). It returns a number of the first row of the first area in the range object

'Comments to the code below:
'DataTable.Cells is the property of DataTable worksheet object. DataTable.Cells returns the range object that represents all cells of object DataTable
'DataTable.Cells.Item is the property of range object "all cells of object DataTable". DataTable.Cells.Item returns the range object at an offset to the specified range
'DataTable.Cells.Item(2, "A") returns the range object "single cell A2"

For HfmFormNumber = 1 To HfmFormAmount
    For DataTableCurrentRow = 2 To DataTableLastRow
    ProfitCenterTypeCurrent = Val(Left(DataTable.Cells.Item(DataTableCurrentRow, "C").Value, 1))
    GLaccountCurrent = Val(DataTable.Cells.Item(DataTableCurrentRow, "A").Value)
    CostCenterTypeCurrent = Val(Mid(DataTable.Cells.Item(DataTableCurrentRow, "C").Value, 6, 2))
        For MappingTableCurrentRow = 2 To MappingTableLastRow 'search on mapping table
            ProfitCenterTypeLeftBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "A").Value)
            ProfitCenterTypeRightBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "B").Value)
            GLaccountLeftBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "C").Value)
            GLaccountRightBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "D").Value)
            CostCenterTypeLeftBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "E").Value)
            CostCenterTypeRightBorder = Val(MappingTable.Cells.Item(MappingTableCurrentRow, "F").Value)
            HfmFormOnMappingTable = MappingTable.Cells.Item(MappingTableCurrentRow, "H").Value
            If (ProfitCenterTypeLeftBorder <= ProfitCenterTypeCurrent And ProfitCenterTypeRightBorder >= ProfitCenterTypeCurrent) _
                And (GLaccountLeftBorder <= GLaccountCurrent And GLaccountRightBorder >= GLaccountCurrent) _
                And (CostCenterTypeLeftBorder <= CostCenterTypeCurrent And CostCenterTypeRightBorder >= CostCenterTypeCurrent) _
                And (HfmFormsArray(HfmFormNumber, 1) = HfmFormOnMappingTable) Then
                    DataTable.Cells.Item(DataTableCurrentRow, HfmFormsArray(HfmFormNumber, 2)).Value = MappingTable.Cells.Item(MappingTableCurrentRow, "I").Value
                    DataTable.Cells.Item(DataTableCurrentRow, HfmFormsArray(HfmFormNumber, 3)).Value = MappingTable.Cells.Item(MappingTableCurrentRow, "J").Value
                Exit For 'stop searching on the mapping table because the desired combination PC_type+GL_acc+CC_type is found. The search by HFM forms will not be stopped
            End If
        Next MappingTableCurrentRow 'continue the search on the mapping table, because the right combination PC_type+GL_acc+CC_type is not found yet
    Next DataTableCurrentRow 'repeat the search for next combination PC_type+GL_acc+CC_type
Next HfmFormNumber 'repeat the search for next HFM-form

Application.ScreenUpdating = True 'return back the application settings
Application.Calculation = OriginalCalcMode 'return back the application settings

End Sub
