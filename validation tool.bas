Attribute VB_Name = "Módulo1"
Sub ValidationTool()

 

    ' Step -2: UnhideAllColumns: Unhide the columns

    UnhideAllColumns

 

    ' Step -1: changecolors:  Clear last validation

    changecolors

 

    ' Step 0: CheckCells: Check if Doc. Header Text and Posting Key are both filled in the same line, showing error if it occurs

    CheckCells

 

    ' Step 1: InsertingH_BasedOn_DocHeaderText:  Insert 'H' to the column A if there's content in column T (Doc. Header Text)

    InsertingH_BasedOn_DocHeaderText

 

    ' Step 21: EraseCells: Erases the L if amount in column Z is zero (updated by vini)

    EraseCells

 

    ' Step 2: InsertingL_BasedOn_PostingKey: Insert 'L' to the column A if there's content in column U (Posting Key)

    InsertingL_BasedOn_PostingKey

   

    ' Step 3: TwoNumbers_afterComma: Format columns Z, AA, and AB as currency with two decimal places

    TwoNumbers_afterComma

   

    ' Step 3.1: IfPostingKey_ThenZMustToBeFilled: Check if the posting key was filled then the column Z (Amount doc currency) must to be provided

    IfPostingKey_ThenZMustToBeFilled

   

    ' Step 4: CheckDebitCredit: Check the sum of credit and debit

    CheckDebitCredit

   

    ' Step 5: CheckRequiredColumns: Check required columns for 'H' cells in column A. If empty = highlight the cells

    CheckRequiredColumns

   

    ' Step 6: CheckHColumnsUtoCG: Checking if H lines are empty at U:CG

    CheckHColumnsUtoCG

   

    ' Step 7: ConvertColumnAHtoUpper: Column AH (Internal Order) and AS (Tax code) are converted to UpperCase

    ConvertColumnAHtoUpper

   

    ' Step 8: JournalCategoryExactly: Journal Category correction to "allocation/operatingcategory" (without spaces) and shows error if it founds something out of the parameters

    JournalCategoryExactly

   

    ' Step 9: AccountMustToBeFilledIn: Check if has Posting Key, then must to have account filled in

    AccountMustToBeFilledIn

   

    ' Step 10: CheckCostObjectForAccount04OperatingCategory: Account 04* must to have a cost object and ONLY one

    CheckCOexistence_or_MultipleCOs_forAccOperatingCategory

   

    ' Step 11: VerifyExternalCode: Check if the ExternalCode (column X) is filled, and shows that the Assignment (column V) must to be filled too - vini

    VerifyExternalCode

   

    ' Step 12: SBSAReversalReason: SB reversal reason must be 05, if SA then must to be reversal reason empty

    SBSAReversalReason

   

    ' Step 13: MustHavePartnerCode: 199* accounts must to have a PartnerCode

    MustHavePartnerCode

   

    ' Step 14: SBMustHaveReversalDate: SB must have reversal date

    SBMustHaveReversalDate

   

    ' Step 15: RUwith4digits: Checking if there are 4 digits na RU

    RUwith4digits

   

    ' Step 16: CheckSpecificAccounts515003000: This account must have PartnerCode 0 and a value on 'Assingment' column

    ' CheckSpecificAccounts515003000

   

    ' Step 17: ArePostingKeyFilled: Check if the PostingKeys column are filled

    ArePostingKeyFilled

   

    'Step 17.5:

    ReplaceDot_to_Bar

   

    ' Step 18: PostingPeriodComparingWith_PeriodSource_And_PostingDate: Checking if the posting period number is the same as the month of the PeriodSource and Posting Date

    'PostingPeriodComparingWith_PeriodSource_And_PostingDate

   

    ' Step 19: IsItCurrentMonth_Or_ClosingMonth: Posting Period only acceptable is current month or closing month

    'IsItCurrentMonth_Or_ClosingMonth

   

    ' Step 20: IsItTheLastDay_of_themonth: Checking if it is the last day of the month filled in "PeriodSource" column or "Posting date"

    'IsItTheLastDay_of_theMonth

   

    ' Step 22: NotOperatingCategoryAccount: Checks if the account is OperatingCategory, if not, it cannot have content in the columns AF, AJ, AK, AH

    NotOperatingCategoryAccount

   

    'Step 23: PartnerCode4Digits:  Adds 0 to PartnerCode with 3 digits, or gives error if doesn't have 3 or 4 digits

    PartnerCode4Digits

   

    'Step: check system RU

    CheckSystemRU

   

    'Step 25: Ref key check

    HighlightaccountAndRefKeyerrors

   

    ' ShowCompletionMessage: Final Step

    ShowCompletionMessage

   

 

End Sub

 

'Step -2:

Sub UnhideAllColumns()

    Dim ws As Worksheet

   

    'Set the ws

    Set ws = ActiveSheet

   

    ws.Columns.Hidden = False

 

End Sub

 

'Step -1: Clear last validation:

Sub changecolors()

    Dim ws As Worksheet

    Dim rng As Range

   

    'Set the ws

    Set ws = ActiveSheet

   

    'Define the range

    Set rng = ws.Range("A4:CG2000")

   

    'Clear corrections

    rng.Interior.ColorIndex = xlNone

   

   

End Sub

 

'Step 0:

Sub CheckCells()

    Dim rng As Range

    Dim cell As Range

 

    ' Set the range from row 4 to the last filled row in column T

    Set rng = Range("T4:T" & Cells(Rows.Count, "T").End(xlUp).Row)

 

    ' Loop through each cell in the range

    For Each cell In rng

        ' Check if the cell in column T and the corresponding cell in column U are filled

        If Not IsEmpty(cell) And Not IsEmpty(cell.Offset(0, 1)) Then

            ' If both cells are filled, display an error message and color the cells red

            MsgBox "Error: Cells " & cell.Address & " and " & cell.Offset(0, 1).Address & " are filled." & vbCrLf & _

            "Both can not be filled in the same line" & vbCrLf & _

            "The Cells were pinted red", vbExclamation, "Error!"

           

            cell.Interior.Color = RGB(241, 101, 101)

            cell.Offset(0, 1).Interior.Color = RGB(241, 101, 101)

        End If

    Next cell

End Sub

 

' Step 1:

Sub InsertingH_BasedOn_DocHeaderText()

    Dim i As Long

   

    ' Loop through each row starting from row 4 to the last possible row (9999)

    For i = 4 To 9999

        ' Check if the cell in column T is not empty

        If Not IsEmpty(Range("T" & i).value) Then

            ' Insert 'H' in the corresponding cell in column A

            Range("A" & i).value = "H"

        End If

    Next i

End Sub

 

' Step 2:

Sub InsertingL_BasedOn_PostingKey()

    Dim i As Long

   

    ' Loop through each row starting from row 4 to the last possible row (9999)

    For i = 4 To 9999

        ' Check if the cell in column U is not empty

        If Not IsEmpty(Range("U" & i).value) Then

            ' Insert 'L' in the corresponding cell in column A

            Range("A" & i).value = "L"

        End If

    Next i

End Sub

 

' Step 3:

Sub TwoNumbers_afterComma()

    ' Declaration of variables

    Dim i As Long

   

    ' Loop through each row starting from row 5 to the last possible row (9999)

    For i = 5 To 9999

        ' Convert and format cells in column Z

        If Not IsEmpty(Range("Z" & i).value) Then

            Range("Z" & i).value = Round(Range("Z" & i).value, 2)

            Range("Z" & i).NumberFormat = "$#,##0.00"

        End If

       

        ' Convert and format cells in column AA

        If Not IsEmpty(Range("AA" & i).value) Then

            Range("AA" & i).value = Round(Range("AA" & i).value, 2)

            Range("AA" & i).NumberFormat = "$#,##0.00"

        End If

       

        ' Convert and format cells in column AB

        If Not IsEmpty(Range("AB" & i).value) Then

            Range("AB" & i).value = Round(Range("AB" & i).value, 2)

            Range("AB" & i).NumberFormat = "$#,##0.00"

        End If

    Next i

End Sub

 

'Step 3.1:

Sub IfPostingKey_ThenZMustToBeFilled()

    Dim lastRow As Long

    Dim i As Long

    Dim errorFound As Boolean

 

    ' Find the last row in column U

    lastRow = Cells(Rows.Count, "U").End(xlUp).Row

 

    ' Initialize errorFound as False

    errorFound = False

 

    ' Loop through the cells in column U starting from row 5

    For i = 5 To lastRow

        ' Check if the cell in column U is not empty

        If Not IsEmpty(Range("U" & i).value) Then

            ' Check if the corresponding cell in column Z is empty or contains text

            If IsEmpty(Range("Z" & i).value) Or Not IsNumeric(Range("Z" & i).value) Then

                ' Color the cell in column Z in red

                Range("Z" & i).Interior.Color = RGB(241, 101, 101)

                ' Set errorFound to True

                errorFound = True

            End If

        End If

    Next i

 

    ' Display the error message if any errors were found

    If errorFound Then

        MsgBox "Enter the 'Amount in Document Currency' values in column Z corresponding to the posting Keys. The cells were painted red.", _

               vbExclamation , "Error"

    End If

End Sub

 

'Step 4:

Sub CheckDebitCredit()

    ' Declaration of variables

    Dim i As Long, j As Long

    Dim sum50 As Double, sum40 As Double

    Dim blockStart As Long, blockEnd As Long

    Dim debitMsg As String

 

    blockStart = 0

   

    ' Loop through each row starting from row 5 to the last possible row (9999)

    For i = 5 To 9999

        If IsEmpty(Range("U" & i).value) Then

            If blockStart <> 0 Then

                blockEnd = i - 1

                ' Calculate the sum of values in column Z for '50' and '40' in column U

                sum50 = 0

                sum40 = 0

                For j = blockStart To blockEnd

                    If Range("U" & j).value = "50" Then

                        sum50 = sum50 + Range("Z" & j).value

                    End If

                    If Range("U" & j).value = "40" Then

                        sum40 = sum40 + Range("Z" & j).value

                    End If

                Next j

                ' Check if the sums are equal and highlight cells if they are not

                If sum50 <> sum40 Then

                    For j = blockStart To blockEnd

                        Range("Z" & j).Interior.Color = RGB(241, 101, 101)

                    Next j

                    debitMsg = "The credit (50) and debit (40) values are not the same." & vbCrLf & _

                               "Credit sum (lines with '50'): " & Format(sum50, "$#,##0.00") & vbCrLf & _

                               "Debit sum (lines with '40'): " & Format(sum40, "$#,##0.00")

                    MsgBox debitMsg, vbExclamation, "Alert!"

                End If

                ' Reset block variables for the next block

                blockStart = 0

                blockEnd = 0

            End If

        Else

            If blockStart = 0 Then

                blockStart = i

            End If

        End If

    Next i

   

    ' Final check in case the last block reaches the end of the range

    If blockStart <> 0 Then

        blockEnd = 9999

        sum50 = 0

        sum40 = 0

        For j = blockStart To blockEnd

            If Range("U" & j).value = "50" Then

                sum50 = sum50 + Range("Z" & j).value

            End If

            If Range("U" & j).value = "40" Then

                sum40 = sum40 + Range("Z" & j).value

            End If

        Next j

        If sum50 <> sum40 Then

            For j = blockStart To blockEnd

                Range("Z" & j).Interior.Color = RGB(241, 101, 101)

            Next j

            debitMsg = "The credit (50) and debit (40) values are not the same." & vbCrLf & _

                       "Credit sum (lines with '50'): " & Format(sum50, "$#,##0.00") & vbCrLf & _

                       "Debit sum (lines with '40'): " & Format(sum40, "$#,##0.00")

            MsgBox debitMsg, vbExclamation, "Alert!"

        End If

    End If

End Sub

 

'Step 5:

Sub CheckRequiredColumns()

    ' Check required columns for 'H' cells in column A and highlight empty cells

    Dim i As Long, col As Variant

    Dim colsToCheck As Variant

    Dim missingContent As Boolean

 

    ' Define the columns to be checked

    colsToCheck = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "S", "T")

   

    For i = 4 To 9999

        If Range("A" & i).value = "H" Then

            For Each col In colsToCheck

                If IsEmpty(Range(col & i).value) Then

                    Range(col & i).Interior.Color = RGB(241, 101, 101)

                    missingContent = True

                End If

            Next col

        End If

    Next i

 

    ' Display a message if there are missing required fields

    If missingContent Then

        MsgBox "One or more required fields are missing and have been highlighted in red", vbExclamation, "Required Fields Missing!"

    End If

End Sub

 

'Step 6:

Sub CheckHColumnsUtoCG()

    Dim i As Long

    Dim col As Long

    Dim rangeToCheck As Range

    Dim cell As Range

    Dim errorMsg As String

    Dim hasError As Boolean

 

    ' Loop through each row starting from 4 to 9999

    For i = 4 To 9999

        ' Check if column A contains "H"

        If Range("A" & i).value = "H" Then

            ' Check if any cell in the range U to CG is not empty

            Set rangeToCheck = Range("U" & i & ":CG" & i)

            For Each cell In rangeToCheck

                If Not IsEmpty(cell.value) Then

                    ' Highlight the cell in red

                    cell.Interior.Color = RGB(241, 101, 101)

                    hasError = True

                    ' Build error message

                    If errorMsg = "" Then

                        errorMsg = "The following rows contain 'H' in column A but have non-empty cells in columns U to CG:" & vbCrLf

                    End If

                    errorMsg = errorMsg & "Row " & i & ", Cell " & cell.Address(False, False) & vbCrLf

                End If

            Next cell

        End If

    Next i

 

    ' Display error message if any discrepancies are found

    If errorMsg <> "" Then

        MsgBox errorMsg, vbExclamation, "Error!"

    End If

End Sub

 

'Step 7:

Sub ConvertColumnAHtoUpper()

 

    Dim i As Long

    Dim cellAH As Range

    Dim cellAS As Range

   

    'Loop through each row starting from 5 to 9999 in column AH and AS

    For i = 5 To 9999

        Set cellAH = Range("AH" & i)

        Set cellAS = Range("AS" & i)

        'Check if the cell in the column AH is not empty and convert to uppercase

        If Not IsEmpty(cellAH.value) Then

            cellAH.value = UCase(cellAH.value)

        End If

       

        'Check if the cell in the column AS is not empty and convert to uppercase

        If Not IsEmpty(cellAS.value) Then

            cellAS.value = UCase(cellAS.value)

        End If

    Next i

End Sub

 

'Step 8:

Sub JournalCategoryExactly()

    Dim i As Long

    Dim cell As Range

    Dim validOptions As Variant

    Dim isValid As Boolean

    Dim errorMsg As String

    Dim optionn As Variant  'Declare the variable option

 

    ' List of valid options (sanitized - OperatingCategory)

    validOptions = Array("Allocation / OperatingCategory", "Bank Accounting", "Derivatives", "Dividend / Minority Interest", "Fixed Assets / Lease", "Forex / EAFE", "Intercompany", "Pension", "Tax", "Volumetric Accounting", "Other")

 

    ' Initialize error message

    errorMsg = "The following rows contain invalid entries in column I:" & vbCrLf

 

    ' Loop through each row starting from 4 to 9999 in column I

    For i = 4 To 9999

        Set cell = Range("I" & i)

        If Not IsEmpty(cell.value) Then

            If LCase(Left(cell.value, 4)) = "allo" Then

                ' Replace the content with sanitized value

                cell.value = "Allocation / OperatingCategory"

            Else

                ' Check if the content is in the list of valid options

                isValid = False

                For Each optionn In validOptions

                    If cell.value = optionn Then

                        isValid = True

                        Exit For

                    End If

                Next optionn

                ' If the content is not valid, highlight the cell in red

                If Not isValid Then

                    cell.Interior.Color = RGB(241, 101, 101)

                    errorMsg = errorMsg & "Row " & i & ": " & cell.value & vbCrLf

                End If

            End If

        End If

    Next i

 

    ' Display error message if there are any invalid entries

    If errorMsg <> "The following rows contain invalid entries in column I:" & vbCrLf Then

        MsgBox errorMsg, vbExclamation, "Error!"

    End If

End Sub

 

'Step 9:

Sub AccountMustToBeFilledIn()

    Dim i As Long

    Dim cellU As Range

    Dim cellV As Range

    Dim hasError As Boolean

 

    ' Initialize error message

    errorMsg = "The following rows have entries in column U but no valid number in column V:" & vbCrLf

 

    ' Loop through each row starting from 5 to 9999 in columns U and V

    For i = 5 To 9999

        Set cellU = Range("U" & i)

        Set cellV = Range("V" & i)

       

        ' Check if the cell in column U is not empty

        If Not IsEmpty(cellU.value) Then

            ' Check if the corresponding cell in column V is a number

            If Not IsNumeric(cellV.value) Or IsEmpty(cellV.value) Then

                ' Highlight the cell in column V in red

                cellV.Interior.Color = RGB(241, 101, 101)

                hasError = True

            End If

        End If

    Next i

 

    ' Display error message if there are any invalid entries

    If hasError Then

        MsgBox "Please provide the account number. The cells without account number were pinted red", vbExclamation, "Error!"

    End If

End Sub

 

'Step 10:

Sub CheckCOexistence_or_MultipleCOs_forAccOperatingCategory()

    Dim ws As Worksheet

    Set ws = ActiveSheet

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row

   

    Dim i As Long

    Dim cellValue As String

    Dim costObjectCount As Integer

   

    For i = 5 To lastRow

        cellValue = ws.Cells(i, "V").value

       

        If Len(cellValue) = 8 And (Left(cellValue, 1) = "4" Or Left(cellValue, 1) = "5") Then

            costObjectCount = 0

           

            If ws.Cells(i, "AF").value <> "" Then costObjectCount = costObjectCount + 1

            If ws.Cells(i, "AH").value <> "" Then costObjectCount = costObjectCount + 1

            If ws.Cells(i, "AI").value <> "" Then costObjectCount = costObjectCount + 1

            If ws.Cells(i, "AJ").value <> "" Then costObjectCount = costObjectCount + 1

           

            If costObjectCount = 0 Then

                MsgBox "You must enter a Cost Object corresponding to the account 04* (OperatingCategory) entered in cell V" & i & "." & vbCrLf & _

                vbCrLf & _

                "The CO cells were painted red.", vbExclamation

                ws.Cells(i, "AF").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AH").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AI").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AJ").Interior.Color = RGB(241, 101, 101)

            ElseIf costObjectCount > 1 Then

                MsgBox "You must inform only ONE Cost Object corresponding to account 04* (OperatingCategory) in cell V" & i & "." & vbCrLf & _

                vbCrLf & _

                "The cells were painted red.", vbExclamation

                ws.Cells(i, "AF").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AH").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AI").Interior.Color = RGB(241, 101, 101)

                ws.Cells(i, "AJ").Interior.Color = RGB(241, 101, 101)

            End If

        End If

    Next i

End Sub

 

'Step 11:

Sub VerifyExternalCode()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim lastRow As Long

    Dim MustChange As Boolean

   

    Set ws = ActiveSheet

    MustChange = False

    lastRow = ws.Rows.Count

  

    For ActualLine = 4 To 9999

       

        ' Check if the cell in column X is not empty

        If Not IsEmpty(ws.Cells(ActualLine, 24).value) Then

           

            ' Color cells in column X and AZ

            ws.Cells(ActualLine, "X").Interior.Color = RGB(241, 101, 101)

            ws.Cells(ActualLine, "AZ").Interior.Color = RGB(250, 224, 129)

            MustChange = True

           

        End If

    Next ActualLine

  

    If MustChange Then

        MsgBox "Please fill the external code in the 'Assignment' column", vbExclamation, "Attention"

    End If

End Sub

 

'Step 12:

Sub SBSAReversalReason()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim lastRow As Long

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 4 to the last row

    For ActualLine = 4 To ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

       

        ' Check if the cell in column J is filled with "SB"

        If ws.Cells(ActualLine, "J").value = "SB" Then

           

            ' Check if the cell in column M is "8"

            If ws.Cells(ActualLine, "M").value = "8" Then

               

                ' Overwrite with "'08" in column M

                ws.Cells(ActualLine, "M").value = "'08"

            Else

                ' If the cell in column M is empty or contains anything else, overwrite with "'05"

                ws.Cells(ActualLine, "M").value = "'05"

            End If

       

        ' Check if the cell in column J is filled with "SA"

        ElseIf ws.Cells(ActualLine, "J").value = "SA" Then

           

            ' Check if the cell in column M is not empty

            If Not IsEmpty(ws.Cells(ActualLine, "M").value) Then

                ' Color the cell in column M yellow

                ws.Cells(ActualLine, "M").Interior.Color = RGB(250, 224, 129)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "The SA document (permanent entry) cannot be filled in with reverse reason (column M). The cell was painted yellow.", vbExclamation, "Error"

    End If

End Sub

 

'Step 13:

Sub MustHavePartnerCode()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim lastRow As Long

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 5 to 9999

    For ActualLine = 5 To 9999

       

        ' Check if the cell in column V starts with "199"

        If Left(ws.Cells(ActualLine, "V").value, 3) = "199" Then

           

            ' Check if the cell in column AP is empty or not a number

            If IsEmpty(ws.Cells(ActualLine, "AP").value) Or Not IsNumeric(ws.Cells(ActualLine, "AP").value) Then

               

                ' Color the cell in column AP red

                ws.Cells(ActualLine, "AP").Interior.Color = RGB(241, 101, 101)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "Account 199* must have a PartnerCode (Column AP). The cells without PartnerCode were pinted red", vbExclamation, "Error"

    End If

End Sub




'Step 14:

Sub SBMustHaveReversalDate()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 4 to 9999

    For ActualLine = 4 To 9999

       

        ' Check if the cell in column J contains the text "SB"

        If ws.Cells(ActualLine, "J").value = "SB" Then

           

            ' Check if the corresponding cell in column N is empty or does not contain a date

            If IsEmpty(ws.Cells(ActualLine, "N").value) Or Not IsDate(ws.Cells(ActualLine, "N").value) Then

                ' Color the cell in column N red

                ws.Cells(ActualLine, "N").Interior.Color = RGB(241, 101, 101)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "Accrual (SB) must have a Reversal Date.", vbExclamation, "Error"

    End If

End Sub

 

'Step 15:

Sub RUwith4digits()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim cellValue As String

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 4 to 9999

    For ActualLine = 4 To 9999

       

        ' Get the value of the cell in column B

        cellValue = ws.Cells(ActualLine, "B").value

       

        ' Check if the cell is not empty

        If cellValue <> "" Then

            ' Check the length of the cell value

            If Len(cellValue) < 4 Then

                ' Add leading zeros to make it 4 digits

                ws.Cells(ActualLine, "B").value = "'" & Right("0000" & cellValue, 4)

            ElseIf Len(cellValue) > 4 Then

                ' Color the cell in column B red

                ws.Cells(ActualLine, "B").Interior.Color = RGB(241, 101, 101)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "Company code error with 5 or more digits", vbExclamation, "Error"

    End If

End Sub




'Step 17:

Sub ArePostingKeyFilled()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim checkLine As Long

    Dim hasError As Boolean

    Dim message As String

   

    Set ws = ActiveSheet

   

    ' Initialize error flag and message

    hasError = False

    message = "Either you must fill in the posting keys in column U or delete some cells filled with L in column A."

   

    ' Loop through each line from row 4 to 9999

    For ActualLine = 4 To 9999

        ' Check if the cell in column A contains the letter "H"

        If ws.Cells(ActualLine, "A").value = "H" Then

            ' Loop through subsequent cells in column A to check for "L"

            For checkLine = ActualLine + 1 To 9999

                ' Exit loop if the cell in column A is empty or not "L"

                If ws.Cells(checkLine, "A").value <> "L" Then Exit For

               

                ' Check if the corresponding cell in column U is empty

                If IsEmpty(ws.Cells(checkLine, "U").value) Then

                    ' Color the cell in column U red

                    ws.Cells(checkLine, "U").Interior.Color = RGB(241, 101, 101)

                   

                    ' Set error flag to true

                    hasError = True

                End If

            Next checkLine

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox message, vbExclamation, "Error"

    End If

End Sub

 

'Step 17.5:

Sub ReplaceDot_to_Bar()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim cell As Range

   

   

    Set ws = ActiveSheet

   

    ' Find the last row with data in column A

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

   

    ' Loop through each row starting from row 4 to the last row

    For i = 4 To lastRow

        ' Check if the cell in column A contains "H"

        If ws.Cells(i, 1).value Like "*H*" Then

            ' Replace period with slash in column C

            Set cell = ws.Cells(i, 3)

            cell.value = Replace(cell.value, ".", "/")

           

            ' Replace period with slash in column D

            Set cell = ws.Cells(i, 4)

            cell.value = Replace(cell.value, ".", "/")

           

            ' Replace period with slash in column G

            Set cell = ws.Cells(i, 7)

            cell.value = Replace(cell.value, ".", "/")

        End If

    Next i

End Sub




'Step 18:

Sub PostingPeriodComparingWith_PeriodSource_And_PostingDate()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim monthE As Integer

    Dim monthG As Integer

    Dim monthD As Integer

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 4 to the last row

    For ActualLine = 4 To ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

       

        ' Check if the cell in column E is filled

        If Not IsEmpty(ws.Cells(ActualLine, "E").value) Then

            ' Get the month values from columns E, G, and D

            monthE = CInt(ws.Cells(ActualLine, "E").value)

            monthG = Month(ws.Cells(ActualLine, "G").value)

            monthD = Month(ws.Cells(ActualLine, "D").value)

           

            ' Check if the months are different

            If monthE <> monthG Or monthE <> monthD Then

                ' Color the cells in columns E, G, and D red

                ws.Cells(ActualLine, "E").Interior.Color = RGB(241, 101, 101)

                ws.Cells(ActualLine, "G").Interior.Color = RGB(241, 101, 101)

                ws.Cells(ActualLine, "D").Interior.Color = RGB(241, 101, 101)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "The month of the posting period is different from the 'Posting date' or 'PeriodSource'. The cells were painted red.", vbExclamation, "Error"

    End If

End Sub

 

' Step 19:

    Sub IsItCurrentMonth_Or_ClosingMonth()

    Dim ws As Worksheet

    Dim ActualLine As Long

    Dim currentMonth As Integer

    Dim closingMonth As Integer

    Dim monthE As Integer

    Dim hasError As Boolean

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Get the current month and closing month (previous month)

    currentMonth = Month(Date)

    closingMonth = currentMonth - 1

    If closingMonth = 0 Then closingMonth = 12 ' Adjust for January

   

    ' Initialize error flag

    hasError = False

   

    ' Loop through each line from row 4 to the last row

    For ActualLine = 4 To ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

       

        ' Check if the cell in column E is filled

        If Not IsEmpty(ws.Cells(ActualLine, "E").value) Then

            ' Get the month value from column E

            monthE = CInt(ws.Cells(ActualLine, "E").value)

           

            ' Check if the month is not the current month or the closing month

            If monthE <> currentMonth And monthE <> closingMonth Then

                ' Color the cell in column E red

                ws.Cells(ActualLine, "E").Interior.Color = RGB(241, 101, 101)

               

                ' Set error flag to true

                hasError = True

            End If

        End If

    Next ActualLine

   

    ' Display error message if necessary

    If hasError Then

        MsgBox "One of the cells in the posting period column is filled with a month different from the current or closing month (previous month). The cell was painted red.", vbExclamation, "Error"

    End If

End Sub

 

'Step 20

 

Sub AddLastDayOfMonth()

    Dim ws As Worksheet

    Dim lastDayD As Date

    Dim lastDayG As Date

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Get the last day of the month for the current date

    lastDayD = DateSerial(Year(Date), Month(Date) + 1, 0)

    lastDayG = DateSerial(Year(Date), Month(Date) + 1, 0)

   

    ' Add the last day of the month to cells D4 and G4

    ws.Cells(4, "D").value = lastDayD

    ws.Cells(4, "G").value = lastDayG

End Sub

 

'Step 21

 

Sub AddTodayDate()

    Dim ws As Worksheet

    Dim today As Date

   

    ' Set the active worksheet

    Set ws = ActiveSheet

   

    ' Get today's date

    today = Date

   

    ' Add today's date to cells D4 and G4

    ws.Cells(4, "D").value = today

    ws.Cells(4, "G").value = today

End Sub

 

'Step 22

 

Sub NotOperatingCategoryAccount()

    Dim ws          As Worksheet

    Dim lastRow     As Long

    Dim cell        As Range

    Dim msg         As String

    Dim isError     As Boolean

   

    ' Set the worksheet you are working on

    Set ws = ActiveSheet        ' Alterar "Sheet1" para o nome da sua planilha

   

    ' Find the last row in column V with data

    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row

   

    ' Loop through each cell in column V starting from row 5

    For Each cell In ws.Range("V5:V" & lastRow)

        isError = False

       

        ' Check if the cell content starts with 4 or 5 and is 8 digits long (OperatingCategory account). If it is not OperatingCategory account, then it should not have content in the cost object

        If Not (Left(cell.value, 1) = "4" Or Left(cell.value, 1) = "5") Or Len(cell.value) <> 8 Then

            ' Check if the corresponding cells in columns AF, AJ, AK, and AH have content

            If ws.Cells(cell.Row, "AF").value <> "" Then

                ws.Cells(cell.Row, "AF").Interior.Color = RGB(241, 101, 101)

                isError = True

            End If

            If ws.Cells(cell.Row, "AJ").value <> "" Then

                ws.Cells(cell.Row, "AJ").Interior.Color = RGB(241, 101, 101)

                isError = True

            End If

            If ws.Cells(cell.Row, "AK").value <> "" Then

                ws.Cells(cell.Row, "AK").Interior.Color = RGB(241, 101, 101)

                isError = True

            End If

            If ws.Cells(cell.Row, "AH").value <> "" Then

                ws.Cells(cell.Row, "AH").Interior.Color = RGB(241, 101, 101)

                isError = True

            End If

        End If

       

        ' If there was an error, display a message box

        If isError Then

            msg = msg & "Row " & cell.Row & ": The account filled does Not accept cost objects (columns AF, AJ, AK, And AH)" & vbCrLf

        End If

    Next cell

   

    ' Display the error message if there were any errors

    If msg <> "" Then

        MsgBox msg, vbExclamation, "Validation Errors"

    End If

End Sub

 

'Step 23

Sub PartnerCode4Digits()

    Dim ws          As Worksheet

    Dim lastRow     As Long

    Dim cell        As Range

    Dim cellValue   As String

    Dim isError     As Boolean

   

    ' Set the worksheet you are working on

    Set ws = ActiveSheet       ' Alterar "Sheet1" para o nome da sua planilha

   

    ' Find the last row in column AP with data

    lastRow = ws.Cells(ws.Rows.Count, "AP").End(xlUp).Row

   

    ' Loop through each cell in column AP starting from row 5

    For Each cell In ws.Range("AP5:AP1100" & lastRow)

        isError = False

       

        ' Check if the cell is not empty

        If Not IsEmpty(cell.value) Then

            cellValue = cell.value

           

            ' Check if the content is not "0"

            If cellValue <> "0" Then

                ' Check if the content has 3 digits

                If Len(cellValue) = 3 Then

                    cell.value = "'0" & cellValue

                    ' Check if the content does not have 4 digits

                ElseIf Len(cellValue) <> 4 Then

                    cell.Interior.Color = RGB(241, 101, 101)

                    isError = True

                End If

            End If

        End If

    Next cell

    ' If there was an error, display a message box

    If isError Then

        MsgBox "PartnerCode must have 4 digits", vbExclamation, "Error!"

    End If

End Sub













' =_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=

' =_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=

' =_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=

' =_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=_=

 

'Step Check System RU

Sub CheckSystemRU()

    Dim SystemGroupC As Variant

    Dim SystemGroupA As Variant

    Dim SystemGroupB As Variant

    Dim SystemGroupD As Variant

   

    ' === Sanitized: lists of RU codes removed to avoid exposure ===

    ' Populate these arrays externally (e.g., Config sheet) if needed.

    SystemGroupC = Array( _

        "[REDACTED]" _

    )

    SystemGroupA = Array( _

        "[REDACTED]" _

    )

    SystemGroupB = Array( _

        "[REDACTED]" _

    )

    SystemGroupD = Array( _

        "[REDACTED]" _

    )

   

    Dim ws As Worksheet

    Dim i As Long

    Dim cellValue As String

   

    Set ws = ActiveSheet

   

    ' Loop through the cells in column B starting from row 4

    For i = 4 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

        If Not IsEmpty(ws.Cells(i, 2)) Then

            ' Format cell value to 4 digits

            cellValue = Format(ws.Cells(i, 2).value, "0000")

           

            ' Check if the value is in any of the lists and call the corresponding sub

            If IsInArray(cellValue, SystemGroupC) Then

                HandleSystemGroupC ws.Cells(i, 2)  ' Organized checks related to SystemGroupC

            ElseIf IsInArray(cellValue, SystemGroupA) Then

                HandleSystemGroupA ws.Cells(i, 2)  ' Organized checks related to SystemGroupA

            ElseIf IsInArray(cellValue, SystemGroupB) Then

                HandleSystemGroupB ws.Cells(i, 2)  ' Organized checks related to SystemGroupB

            ElseIf IsInArray(cellValue, SystemGroupD) Then

                HandleSystemGroupD ws.Cells(i, 2)  ' Organized checks related to SystemGroupD

            End If

        End If

    Next i

End Sub

 

' Function to check if a value is in an array

Function IsInArray(value As String, numArray As Variant) As Boolean

    Dim i As Long

    ' Loop through the array

    For i = LBound(numArray) To UBound(numArray)

        ' Check if value is found in the array

        If numArray(i) = value Then

            IsInArray = True

            Exit Function

        End If

    Next i

    ' Value not found in array

    IsInArray = False

End Function

 

Sub HandleSystemGroupC(cell As Range)

    Dim startRow As Long, endRow As Long

    DefineRowRange cell.Worksheet, cell.Row, startRow, endRow ' Call DefineRowRange subroutine to find startRow and endRow. The cell.Row is sent to the rowNum parameter

End Sub

 

Sub HandleSystemGroupA(cell As Range)

    Dim startRow As Long, endRow As Long

    DefineRowRange cell.Worksheet, cell.Row, startRow, endRow ' Call DefineRowRange subroutine to find startRow and endRow. The cell.Row is sent to the rowNum parameter

End Sub

 

Sub HandleSystemGroupB(cell As Range)

    Dim startRow As Long, endRow As Long

    DefineRowRange cell.Worksheet, cell.Row, startRow, endRow ' Call DefineRowRange subroutine to find startRow and endRow. The cell.Row is sent to the rowNum parameter

End Sub

 

Sub HandleSystemGroupD(cell As Range)

    Dim startRow As Long, endRow As Long

    DefineRowRange cell.Worksheet, cell.Row, startRow, endRow ' Call DefineRowRange subroutine to find startRow and endRow. The cell.Row is sent para o parâmetro rowNum

   

    ' Print startRow and endRow values (to test if the code is working)

    ' MsgBox "startRow: " & startRow & ", endRow: " & endRow

   

    CheckCC_SystemGroupD cell, startRow, endRow ' Call CheckCC_SystemGroupD subroutine with cell and row parameters

    ReversalReasonAMP cell, startRow, endRow ' Call ReversalReasonAMP subroutine with cell and row parameters

   

End Sub

 

' ===================================================================================================================

 

' Subroutine to define startRow and endRow based on column A data

' Sets the "block" (H and L's) of that RU which will be analysed

Sub DefineRowRange(ws As Worksheet, rowNum As Long, ByRef startRow As Long, ByRef endRow As Long)

    Dim rowCount As Long

    Dim i As Long

   

    ' This line finds the last used row in column A

    rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

   

    ' Find the first cell below the current cell in column B that does not contain "L"

    For i = rowNum + 1 To rowCount

        If Left(ws.Cells(i, 1).value, 1) <> "L" Then

            endRow = i - 1

            Exit For

        End If

    Next i

   

    ' If no cell was found different from "L", set endRow to the last row in column A

    If endRow = 0 Then

        endRow = rowCount

    End If

   

    ' Determine the starting row for checking in column A

    startRow = rowNum + 1

End Sub

 

' Subroutine to check specific conditions in column A and AF

' If the account is OperatingCategory (starts with 4* or 5* and has 8 digits, then the CC must have 6 digits)

Sub CheckCC_SystemGroupD(cell As Range, startRow As Long, endRow As Long)

    Dim ws As Worksheet

    Dim i As Long

   

    Set ws = cell.Worksheet

   

    ' Check each row between startRow-1 and endRow in column A

    For i = startRow - 1 To endRow

        If Left(ws.Cells(i, 1).value, 1) = "L" Then

            ' Check if the corresponding cell in column V starts with 4 or 5 and has 8 digits

            If (Left(ws.Cells(i, "V").value, 1) = "4" Or Left(ws.Cells(i, "V").value, 1) = "5") And Len(ws.Cells(i, "V").value) = 8 Then

           

                ' Check if cells in columns AH, AI, and AJ are empty

                If IsEmpty(ws.Cells(i, "AH").value) And IsEmpty(ws.Cells(i, "AI").value) And IsEmpty(ws.Cells(i, "AJ").value) Then

               

                    ' Check if the corresponding cell in column AF (Cost Center) is empty or not 6 digits (if yes, then show error)

                    If IsEmpty(ws.Cells(i, "AF").value) Or Not IsNumeric(ws.Cells(i, "AF").value) Or Len(ws.Cells(i, "AF").value) <> 6 Then

               

                        ' Color the cell in column AF purple

                        ws.Cells(i, "AF").Interior.Color = RGB(128, 0, 128) ' purple color

                        MsgBox "Error: The CostCenter in cell AF" & i & " must be filled or it doesn't have 6 digits. The cell was painted purple", vbExclamation, "Error"

                       

                    End If

                End If

            End If

        End If

    Next i

End Sub

 

' Subroutine to check specific conditions in column J and update column M

Sub ReversalReasonAMP(cell As Range, startRow As Long, endRow As Long)

    Dim ws As Worksheet

    Set ws = cell.Worksheet

   

    ' Print startRow and endRow values (to test if the code is working)

    'MsgBox "startRow: " & startRow & ", endRow: " & endRow

   

    ' Check the first cell of the RowRange in column J

    If ws.Cells(startRow - 1, "J").value = "SB" Then

        ' Update the corresponding cell in column M to "Y3"

        ws.Cells(startRow - 1, "M").value = "Y3"

    End If

End Sub

 

Sub HighlightaccountAndRefKeyerrors()

    Dim ws As Worksheet

    Dim cell As Range

    Dim errorFound As Boolean

    Dim errorMessage As String

   

    ' Set the worksheet

    Set ws = ActiveSheet

   

    ' Initialize errorFound flag

    errorFound = False

   

    ' Loop through each row

    For Each cell In ws.Range("V2:V" & ws.Cells(ws.Rows.Count, "V").End(xlUp).Row)

        If Left(cell.value, 1) = "6" Or Left(cell.value, 1) = "2" Then

            If cell.Offset(0, 18).value <> "" And cell.Offset(0, 19).value <> "" Then

                ' Highlight cells in red

                cell.Interior.Color = RGB(255, 0, 0)

                cell.Offset(0, 18).Interior.Color = RGB(255, 0, 0)

                cell.Offset(0, 19).Interior.Color = RGB(255, 0, 0)

               

                ' Set errorFound flag

                errorFound = True

            End If

        End If

    Next cell

   

    ' Show error message if errors were found

    If errorFound Then

        errorMessage = "Accounts starting with 6 or 2 cannot have a reference key field. The cells were painted in red."

        MsgBox errorMessage, vbExclamation, "Error"

    End If

End Sub




'Final Step

Sub ShowCompletionMessage()

    MsgBox "Validation completed.", vbInformation, "Status"

 

End Sub

