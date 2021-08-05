Attribute VB_Name = "DashboardUpdate"
Sub ForcastedProduction()
    Dim EBIT_Target As Workbook
    Dim EBITSheetName As String
    
    EBITSheetName = "EBIT (WO Issue)"
    Set EBIT_Target = Workbooks.Open("P:\_Departments\FC\05 - Accounting\01 Financial Accounting\04 Transfer Pricing\01 Vehicle\Profit Analysis\EBIT and TP Analysis\EA1 2021\Approved EBIT Target.xlsx", , True)
    ActiveWindow.Visible = False
    
   
    'Fill in V167 numbers by transposing
    Sheet1.Range("C32:C43").Value = WorksheetFunction.Transpose(EBIT_Target.Sheets(EBITSheetName).Range("B25:M25"))
    'Fill in C167 numbers by transposing
    Sheet1.Range("C61:C72").Value = WorksheetFunction.Transpose(EBIT_Target.Sheets(EBITSheetName).Range("B24:M24"))         '*******************Change the import file ranges if locations change in import file
    'Fill in X167 numbers by transposing
    Sheet1.Range("C90:C101").Value = WorksheetFunction.Transpose(EBIT_Target.Sheets(EBITSheetName).Range("B26:M26"))
    
    EBIT_Target.Close (False)


End Sub
Sub Forcasted_Material()
    
    Dim EBIT_Target As Workbook
    Dim SheetName As String
    Dim V167Row As Range
    Dim X167Row As Range
    Dim C167Row As Range
    Dim CostingMCColumn As Range
    
    
    Set EBIT_Target = Workbooks.Open("P:\_Departments\FC\05 - Accounting\01 Financial Accounting\04 Transfer Pricing\01 Vehicle\Profit Analysis\EBIT and TP Analysis\EA1 2021\Approved EBIT target.xlsx", , True)
    ActiveWindow.Visible = False
    SheetName = "EBIT (WO issue)"
    
    
    'V167 Material average per vehicle number
    Set V167Row = EBIT_Target.Sheets(SheetName).Range("40:49").Find("V167", LookAt:=xlWhole).EntireRow
    Set CostingMCColumn = EBIT_Target.Sheets(SheetName).Range("A:P").Find("Costing-MC", LookAt:=xlWhole).EntireColumn
    
    Sheet1.Range("R54").Value = Application.Intersect(V167Row, CostingMCColumn)
    
    'X167 Material average per vehicle number
    Set X167Row = EBIT_Target.Sheets(SheetName).Range("40:49").Find("X167", LookAt:=xlWhole).EntireRow
    
    Sheet1.Range("R112").Value = Application.Intersect(X167Row, CostingMCColumn)
    
    'C167 Material average per vehicle number
    Set C167Row = EBIT_Target.Sheets(SheetName).Range("40:49").Find("C167", LookAt:=xlWhole).EntireRow
    
    Sheet1.Range("R83").Value = Application.Intersect(C167Row, CostingMCColumn)
    
    'Close Material Cost workbook
    EBIT_Target.Close False
End Sub
Sub Forcasted_Sales_LOH_Profit_EBIT()           'This also inputs alot of data from Approved EBIT to the Profit Analysis Tab
    Application.AskToUpdateLinks = False
    Dim EBIT_Target As Workbook
    Dim FindEBITCell As String
    Dim FindProfitCell As String
    Dim EBITSheetName As String
    
    EBITSheetName = "EBIT (WO issue)" '****************************Change this if the sheet name changes
    
    Set EBIT_Target = Workbooks.Open("P:\_Departments\FC\05 - Accounting\01 Financial Accounting\04 Transfer Pricing\01 Vehicle\Profit Analysis\EBIT and TP Analysis\EA1 2021\Approved EBIT target.xlsx")
    
    
    
'This is the section that fills for Forcasted Sales (THIS DEPENDS ON THE CELL THAT SAYS "BREAK DOWN OF SALES")
    
    'C167
    Sheet1.Range("F61:F72").Value = WorksheetFunction.Transpose(Range(EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(1, 1), EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(1, 12)))
    
    'V167
    Sheet1.Range("F32:F43").Value = WorksheetFunction.Transpose(Range(EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(2, 1), EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(2, 12)))
    
    'X167
    Sheet1.Range("F90:F101").Value = WorksheetFunction.Transpose(Range(EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(3, 1), EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Break down of Sales", LookAt:=xlPart).Offset(3, 12)))
    
    
'This is the section that fills for Forcasted LOH

    Dim TPOHCell As Range
    Dim AverageCell As Range
    Dim TPOH_Value As Variant
    Dim CostingOH As Range
    Dim COstingOH_Value As Variant
    Dim CostingDestination As Range
    Dim TPDestination As Range

    'Finds the headers of TPOH, Average, and Costing-OH, grabs the entire colums/rows, then finds the intersection point.
    Set TPOHCell = EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("TP-OH", LookAt:=xlPart).EntireColumn
    Set AverageCell = EBIT_Target.Sheets(EBITSheetName).Range("A25:A80").Find("Average", LookAt:=xlPart).EntireRow
    Set CostingOHCell = EBIT_Target.Sheets(EBITSheetName).Range("A1:Y200").Find("Costing-OH", LookAt:=xlPart).EntireColumn

    'Assigns the value from the intesection to variables
    TPOH_Value = Application.Intersect(TPOHCell, AverageCell)
    COstingOH_Value = Application.Intersect(CostingOHCell, AverageCell)

    'Find destination for TP-OH and COsting-OH value in Calculation worksheet
    Set TPDestination = Sheet1.Range("A1:Y200").Find("LOH per Vehicle, TP", LookAt:=xlPart).Offset(0, 4)
    Set CostingDestination = TPDestination.Offset(2, 0)

    'Asssign values
    TPDestination.Value = TPOH_Value
    CostingDestination.Value = COstingOH_Value



'This is the section that fills for Forcasted EBIT
    EBIT_Target.Sheets(EBITSheetName).Activate
    FindEBITCell = EBIT_Target.Sheets(EBITSheetName).Range("A:A").Find("Adjusted EBIT", LookAt:=xlPart).Address(False, False)
    Sheet1.Range("AA4:AA15").Value = WorksheetFunction.Transpose(EBIT_Target.Sheets(EBITSheetName).Range(Range(FindEBITCell).Offset(0, 1), Range(FindEBITCell).Offset(0, 12)))

'This is the section that fills for Forcasted Profit
    FindProfitCell = EBIT_Target.Sheets(EBITSheetName).Range("A:A").Find("Billed Profit", LookAt:=xlPart).Address(False, False)
    Sheet1.Range("X4:X15").Value = WorksheetFunction.Transpose(EBIT_Target.Sheets(EBITSheetName).Range(Range(FindProfitCell).Offset(0, 1), Range(FindProfitCell).Offset(0, 12)))



'Close EBIT workbook
    EBIT_Target.Close


End Sub





Sub DashboardUpdateMain()
'This makes the code faster and reduces likely hood of warning popups that may interfear with the code
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

'On errors, the code will continue through the code trying to update as much as possible, a error box at the end will be displayed.
    'On Error Resume Next
    
    
    
'Jumps into different sub's (Main part of code)
    UpdateActuals.UpdateActualsMain
    ForcastedProduction
    Forcasted_Sales_LOH_Profit_EBIT 'Profit Analysis Forecast stuff too
    Forcasted_Material
    
    

'On errors, it will create a message box that says error, but if not, then it will say OP succeesfull
    If Err <> 0 Then
        MsgBox "I could not Update Dashboard. There was most likely a problem opening up import files. You need to find Nick.", , "Dashboard Update Failure"
    Else:
        MsgBox "Dashboard Updated Successfully!"
    End If
    On Error GoTo 0
    
'Returns the excel settings back to normal
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub


