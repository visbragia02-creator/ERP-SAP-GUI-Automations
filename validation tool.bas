Attribute VB_Name = "ValidationTool_v1"
Sub ValidationTool()

    ' Standard validation flow for financial document uploads
    ' Optimized to ensure data integrity before system ingestion

    ' Step -2: Clear view
    UnhideAllColumns

    ' Step -1: Visual reset
    ClearLastValidation

    ' Step 0: Integrity Check
    CheckHeaderVsPostingKey

    ' Step 1: Document Classification
    InsertRowIndicators

    ' Step 2: Currency Formatting
    FormatCurrencyColumns

    ' Step 3: Mandatory Fields Validation
    CheckRequiredDataPoints

    ' Step 4: Accounting Logic - Balance Check
    CheckDebitCreditBalance

    ' Step 5: Advanced Logic - Cost Objects and Categorization
    ValidateCostObjects
    ValidateJournalCategories
    ValidateAccountFormats

    ' Step 6: Specific Compliance Rules (Sanitized)
    ValidateReversalReasons
    ValidatePartnerCodes
    ValidateCompanyCodes

    ' Final Step
    ShowCompletionMessage

End Sub

' --- CORE FUNCTIONS ---

Sub UnhideAllColumns()
    ActiveSheet.Columns.Hidden = False
End Sub

Sub ClearLastValidation()
    ' Resets background colors to clear previous error highlights
    ActiveSheet.Range("A4:CG2000").Interior.ColorIndex = xlNone
End Sub

Sub CheckHeaderVsPostingKey()
    Dim cell As Range
    For Each cell In Range("T4:T" & Cells(Rows.Count, "T").End(xlUp).Row)
        If Not IsEmpty(cell) And Not IsEmpty(cell.Offset(0, 1)) Then
            MsgBox "Validation Error: Line " & cell.Row & " contains both Header Text and Posting Key. " & _
                   "Only one entry type is allowed per line.", vbExclamation, "Data Integrity Alert"
            cell.Interior.Color = RGB(241, 101, 101)
            cell.Offset(0, 1).Interior.Color = RGB(241, 101, 101)
        End If
    Next cell
End Sub

Sub FormatCurrencyColumns()
    Dim i As Long
    For i = 5 To 9999
        ' Format columns Z, AA, AB for standard financial reporting
        If IsNumeric(Range("Z" & i).Value) And Not IsEmpty(Range("Z" & i)) Then
            Range("Z" & i).Value = Round(Range("Z" & i).Value, 2)
            Range("Z" & i).NumberFormat = "$#,##0.00"
        End If
    Next i
End Sub

Sub CheckDebitCreditBalance()
    Dim i As Long, j As Long
    Dim sumCredit As Double, sumDebit As Double
    Dim bStart As Long, bEnd As Long
