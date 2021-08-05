Attribute VB_Name = "debitSupplier"

Global EmailBody As String


Sub PDFTemplate() 'Browse button for template

    Dim PDFFolder As FileDialog
    Set PDFFolder = Application.FileDialog(msoFileDialogFilePicker)
    With PDFFolder
        .Title = "Select PDF file to attach"
        .Filters.Add "PDF", "*.pdf", 1
        If .Show <> -1 Then GoTo NOSelection
        Sheet2.Range("Q2").Value = .SelectedItems(1)
    End With
NOSelection:


End Sub
Sub PDFSave() 'Browse button for Save as folder

    Dim PDFFolder As FileDialog
    Set PDFFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With PDFFolder
        .Title = "Select Folder to save PDF in"
        If .Show <> -1 Then GoTo NOSelection
        Sheet2.Range("Q3").Value = .SelectedItems(1)
    End With
NOSelection:


End Sub
Sub FilterForReadyRows() ' filters to look for suppliers are ready to have a form created

    With Sheet2.Range("7:7")
        'Reset Filter, I use on error resume next as a sandwhich because it will error if there is no filtered applied before the call
        On Error Resume Next
        Sheet2.ShowAllData
        On Error GoTo 0
        
   
        .AutoFilter 21, "<>0"
        
        'Filter to make sure we have not already created a form for the items. If there is text in these columns, that means we have already made a form for it.
        .AutoFilter 27, ""
        
        'Checks if there are no rows to make forms with. If none, then quit code
        If (Sheet2.Range("G999999").End(xlUp).Row = 7) Then
            MsgBox ("After filtering, There is no data ready for forms to be created. ")
            End
        End If
    End With
        
        


End Sub
Function getSupplierNumbersArrayList(thisRange) As ArrayList
    'The idea behind this function is that it will go through all the supplier numbers that result after the subroutine for filtering is done.
    'Then it will add all unique supplier numbers once, meaning that no duplicates will be added. This tells us which suppliers we will be creating forms for.

    Dim SupplierArrayList As ArrayList
    Set SupplierArrayList = New ArrayList
    
    Set thisRange = Sheet2.Range("G8", Range("G1000000").End(xlUp)).SpecialCells(xlCellTypeVisible)

    For Each c In thisRange
        If SupplierArrayList.Contains(c.Value) Then
            GoTo Skip
        
        Else:
           SupplierArrayList.Add (c.Value)
        End If
Skip:
    Next
    
    'This is what is returned
   Set getSupplierNumbersArrayList = SupplierArrayList
    
     
    

End Function
Sub FillDebitForm(PDFFile, PDFSaveFolder, PDFBreakdownName, supplierNumber, Amount, InvoiceNumber, GLAccount, InvoiceDate, CostCenter, SupplierName, SupplierAddress, Purpose, Description, JurisdictionCode, FormFileName)
'Declare Acrobat app and files
Dim pdfApp As Acrobat.AcroApp
Dim pdfDoc As Acrobat.AcroAVDoc
Dim SupportDoc As Acrobat.AcroPDDoc


Dim pdf_form As AFORMAUTLib.AFormApp
Dim pdf_for_fields As AFORMAUTLib.Fields

'Declare all fields
Dim FieldSupplierNumber As AFORMAUTLib.Field
Dim FieldAmount As AFORMAUTLib.Field
Dim FieldInvoiceNumber As AFORMAUTLib.Field
Dim FieldGLAccount As AFORMAUTLib.Field
Dim FieldInvoiceDate As AFORMAUTLib.Field
Dim FieldCostCenter As AFORMAUTLib.Field
Dim FieldPurpose As AFORMAUTLib.Field
Dim FieldSupplierName As AFORMAUTLib.Field
Dim FieldSupplierAddress As AFORMAUTLib.Field
Dim FieldDescription As AFORMAUTLib.Field
Dim FieldAdjustmentAmount As AFORMAUTLib.Field
Dim FieldDebitCheckBox As AFORMAUTLib.Field


PDFSaveAsFullPath = PDFSaveFolder & "\" & "FC-54.07 " & supplierNumber & ".pdf"

'Will delete an already existing for if there is one under the same file path
If (Dir(PDFSaveAsFullPath) <> "") Then
    Kill PDFSaveAsFullPath
End If


'Create Acrobat app instance and create document instance
Set pdfApp = CreateObject("AcroExch.App")
Set pdfDoc = CreateObject("AcroExch.AVDoc")




'Open Templete PDF

If pdfDoc.Open(PDFFile, "") = True Then
    pdfDoc.BringToFront
    pdfApp.Show

    
    'Create instance of an app that recognizes elemenets in form
    Set pdf_form = CreateObject("AFORMAUT.App")
    
    'Create variables that represent the fields in the templete
    Set FieldSupplierNumber = pdf_form.Fields("MBUSI Supplier Number")
    Set FieldAmount = pdf_form.Fields("Amount")
    Set FieldInvoiceNumber = pdf_form.Fields("Supplier Invoice Number")
    Set FieldGLAccount = pdf_form.Fields("GL Account")
    Set FieldInvoiceDate = pdf_form.Fields("Supplier Invoice Date")
    Set FieldCostCenter = pdf_form.Fields("4380")
    Set FieldPurpose = pdf_form.Fields("Purpose")
    Set FieldSupplierName = pdf_form.Fields("Supplier Name")
    Set FieldSupplierAddress = pdf_form.Fields("Address")
    Set FieldDescription = pdf_form.Fields("Description  Reason attach calculation documentationRow1")
    Set FieldAdjustmentAmount = pdf_form.Fields("AmountRow1_2")
    Set FieldDebitCheckBox = pdf_form.Fields("NON DOWNTIME DEBIT")
    
    'Set the value of those fields
    FieldSupplierNumber.Value = supplierNumber
    FieldAmount.Value = Amount
    FieldInvoiceNumber.Value = InvoiceNumber
    FieldGLAccount.Value = GLAccount
    FieldInvoiceDate.Value = InvoiceDate
    FieldCostCenter.Value = CostCenter
    FieldPurpose.Value = Purpose
    FieldSupplierName.Value = SupplierName
    FieldSupplierAddress.Value = SupplierAddress
    FieldDescription.Value = Description
    FieldAdjustmentAmount.Value = Amount
    FieldDebitCheckBox.Value = "On"
    
    Set SupportDoc = pdfDoc.GetPDDoc
    
    PDFSaveAsFullPath = PDFSaveFolder & "\" & "FC-54.07 " & supplierNumber & ".pdf"
    SupportDoc.Save PDSaveFull, PDFSaveAsFullPath
        
    

    
    
    'Close and clear memory
    pdfDoc.Close True
    SupportDoc.Close
        
    Set FieldSupplierNumber = Nothing
    Set FieldAmount = Nothing
    Set FieldInvoiceNumber = Nothing
    Set FieldGLAccount = Nothing
    Set FieldInvoiceDate = Nothing
    Set FieldCostCenter = Nothing
    Set FieldPurpose = Nothing
    Set FieldSupplierName = Nothing
    
    pdfApp.Exit
        
    'run the sub located in the payment module to combine pdf's
    payment.combinePDF PDFSaveAsFullPath, PDFBreakdownName
    
    'find the invoice filename and then combine with the pdf that was just combined above
    InvoiceFileName = Application.WorksheetFunction.VLookup(JurisdictionCode, Sheet1.Range("H:S"), 12, 0)
    InvoiceFilePath = "P:\_Departments\FC\11 - Cross functional Topics\Digitalization\Property Tax\TestPayment\Invoices\" & InvoiceFileName
    payment.combinePDF PDFSaveAsFullPath, InvoiceFilePath
     
    End If
    

End Sub
Sub FillPDFs(supplierNumber, thisRange, DebitFormTemplatePath, DebitFormSaveAsFolder, PDFBreakdownFolder)
'''This sub takes in a supplier number, looks in the filtered list for all line items with that  same supplier number, then takes the data from
'the Description, amount, and asset number columns and inputs those value on the next line on the breakdown template. Then for each supplier, the range will be exported as a PDF

    Dim CreatedFormRange As Range
    Set CreatedFormRange = Nothing
    
    'Creating variable to store the cumulative amount of taxes owed for all tools
    AmountTotal = 0
    
    'Fill Data based on supplier Numbers
    For Each c In thisRange
        If supplierNumber = c.Value Then
            AssetNumber = c.Offset(0, -5).Value
            ToolDescription = c.Offset(0, -4).Value
            Amount = c.Offset(0, 15).Value
            AmountTotal = AmountTotal + Amount
            SupplierName = c.Offset(0, 1).Value
            SupplierAddress = c.Offset(0, 2).Value
            InvoiceNumber = c.Offset(0, 10).Value
            GLAccount = "Placeholder GL"
            InvoiceDate = Date 'This grabs current date
            CostCenter = "4380-9999 placeholder"
            Purpose = c.Offset(0, 12).Value
            FormDescription = c.Offset(0, 8).Value & " - " & c.Offset(0, 11).Value & " - " & SupplierAddress
            JurisdictionCode = c.Offset(0, 9).Value
            FormFileName = c.Offset(0, 18).Value
            
            'Creates a variable range that stores all cells that will be marked "yes" in column "Form already created" so at the end when the form for that supplier is created, It will loop through the cells in this range and fill "yes"
            If Not CreatedFormRange Is Nothing Then
                Set CreatedFormRange = Application.Union(CreatedFormRange, c.Offset(0, 20)) 'the union function appends ranges to variable
            Else:
                Set CreatedFormRange = c.Offset(0, 20)
            End If
            
            'Fill breakdown template with the data
            With Sheet6
                If (.Range("B500").End(xlUp).Offset(1, 0) = "" And .Range("B500").End(xlUp).Offset(1, 3) = "" And .Range("B500").End(xlUp).Offset(1, 7) = "") Then
                    .Range("B500").End(xlUp).Offset(1, 0) = AssetNumber
                    .Range("E500").End(xlUp).Offset(1, 0) = ToolDescription
                    .Range("J500").End(xlUp).Offset(1, 0) = Amount
                'Catches if there is not all the information required per line item on breakdown template and stops macro
                Else:
                    Sheet6.Visible = xlSheetVisible
                    MsgBox (" THE PROCESS HAS STOOPED" & Chr(13) & Chr(13) & "ERROR: There is not a Asset Number, Description, and/or amount for each tool, thus we cannot make the tool breakdown summary for supplier " & supplierNumber & ". Look at the Breakdown sheet and make sure to fill in the correct value in the *Supplier Debits* Sheet." & Chr(13) & Chr(13) & "Fix and run again. No duplicates will be created.")
                    Sheet6.Activate
                    Call CreateEmail
                    End
                    
                End If
            End With
        End If

    Next
    'Sums up tax amounts on breakdown sheet
    With Sheet6.Range("J500").End(xlUp).Offset(2, 0)
        .Value = AmountTotal
        .Font.Bold = True
    End With
    
    'Print the breakdown as PDF
    Sheet6.Activate
    
    Dim PDFRange As Range
    Dim PDFBreakdownName As String
    Set PDFRange = Sheet6.Range("A1", Range("J500").End(xlUp).Offset(5, 1))
    
    PDFBreakdownName = PDFBreakdownFolder & "\" & supplierNumber & "_Breakdown.pdf"
    Sheet6.Visible = xlSheetVisible 'needs to be visible to do this pdf print
    PDFRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PDFBreakdownName
    
    Sheet2.Activate
    
    'Fill the Supplier Request debit form
    FillDebitForm DebitFormTemplatePath, DebitFormSaveAsFolder, PDFBreakdownName, supplierNumber, AmountTotal, InvoiceNumber, GLAccount, InvoiceDate, CostCenter, SupplierName, SupplierAddress, Purpose, FormDescription, JurisdictionCode, FormFileName
    
    'fills "yes" for the form already created column on all the cells in the Unioned range above
    For Each c In CreatedFormRange
        c.Value = "Yes"
    Next
        
Skip:
    
End Sub
Sub CreateEmail()
 Dim objOutlook As Object
    Dim objEmail As Object
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objEmail = objOutlook.CreateItem(0)
    
    With objEmail
        .To = "nicholas.malmquist@daimler.com"
        .CC = ""
        .Subject = "Test"
        .HTMLBody = EmailBody
        .Display
    End With
End Sub

''''The main macro that is run
Sub Main()
    Application.ScreenUpdating = False
    
     '''Pop up to make sure that person wants to run auto fill pdf
   If MsgBox("This will take control your computer! Are you sure you want to auto fill PDF's?", vbYesNo, "AutoFill PDF?") = vbYes Then
       If MsgBox("Are You Sure?", vbYesNo, "AutoFill PDF?") = vbNo Then Exit Sub
    Else: Exit Sub
    End If
    
    
    Dim SupplierArrayList As ArrayList
    Dim SupplierNumbersRange As Range
    EmailBody = "Nicholas,"
    'All file paths
    PDFTemplatePath = Sheet2.Range("Q2").Value
    PDFSaveAsFolder = Sheet2.Range("Q3").Value
    PDFBreakdownFolder = "P:\_Departments\FC\11 - Cross functional Topics\Digitalization\Property Tax\TestDebit\BreakdownTempFiles"

    'Sub routine above
    FilterForReadyRows
    
    Sheet2.Activate 'need the sheet to be activate to use end.offset in the line below
    Set SupplierNumbersRange = Sheet2.Range("G5", Range("G1000000").End(xlUp)).SpecialCells(xlCellTypeVisible) 'gives the range of supplier numbers that are ready for forms to be made
    
    'function above that will get each supplier number without duplicates from the range
    Set SupplierArrayList = getSupplierNumbersArrayList(SupplierNumbersRange)
    
    'sub that takes in the supplier number, and fills the breakdown template, ending with exporting breakdown template as PDF
    'The .Body property only will take one value and cannot be called multiple times to append text on email, so we have to append everything (by concatenation)to a single variable "EmailBody" then use that variable in the .Body property
    For Each supplierNumber In SupplierArrayList
        FillPDFs supplierNumber, SupplierNumbersRange, PDFTemplatePath, PDFSaveAsFolder, PDFBreakdownFolder
        Hyperlink = Replace(PDFSaveAsFolder & "\" & "FC-54.07 " & supplierNumber & ".pdf", " ", "%20") 'HTML href hyper link needs %20 instead of spaces to workx
        EmailBody = EmailBody & "<br><br><br><br>Please Sign this Supplier Debit Form from Property Tax and notify Rhonda Mccray when complete:<br> " & "<a href=" & Hyperlink & ">" & PDFSaveAsFolder & "\" & "FC-54.07 " & supplierNumber & ".pdf" & "</a>"
    Next

    Call CreateEmail
    
    


End Sub

