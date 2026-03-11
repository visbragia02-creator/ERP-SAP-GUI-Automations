Attribute VB_Name = "Módulo2"
Sub Standard_Template_Converter()
    Dim ws As Worksheet
    Dim NewTemplateSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    Dim suffix As Integer

    ' Set the original sheet
    Set ws = ActiveSheet

    ' Optimize column widths on the original sheet
    ws.Columns.AutoFit

    ' Determine the new sheet name
    sheetName = "Converted_Template"
    suffix = 1
    Do While SheetExists(sheetName)
        sheetName = "Converted_Template_" & suffix
        suffix = suffix + 1
    Loop

    ' Adds the new sheet with the determined name
    Sheets.Add(After:=ws).Name = sheetName
    Set NewTemplateSheet = Sheets(sheetName)

    ' Creates the Template Structure
    NewTemplateSheet.Range("A1:K1").Merge
    NewTemplateSheet.Range("L1:AB1").Merge

    NewTemplateSheet.Cells(1, 1).value = "SECTION - HEADER"
    NewTemplateSheet.Cells(1, 1).HorizontalAlignment = xlCenter

    NewTemplateSheet.Cells(1, 21).value = "SECTION - LINE"
    NewTemplateSheet.Cells(1, 21).HorizontalAlignment = xlCenter

    ' ==========================================
    ' Column names (Technical / System Names)
    ' ==========================================
    NewTemplateSheet.Range("B2").value = "[SYS_HDR_FIELD_01]"
    NewTemplateSheet.Range("C2").value = "[SYS_HDR_FIELD_02]"
    NewTemplateSheet.Range("D2").value = "[SYS_HDR_FIELD_03]"
    NewTemplateSheet.Range("E2").value = "[SYS_HDR_FIELD_04]"
    NewTemplateSheet.Range("F2").value = "[SYS_HDR_FIELD_05]"
    NewTemplateSheet.Range("G2").value = "[SYS_HDR_FIELD_06]"
    NewTemplateSheet.Range("H2").value = "[SYS_HDR_FIELD_07]"
    NewTemplateSheet.Range("I2").value = "[SYS_HDR_FIELD_08]"
    NewTemplateSheet.Range("J2").value = "[SYS_HDR_FIELD_09]"
    NewTemplateSheet.Range("K2").value = "[SYS_HDR_FIELD_10]"

    NewTemplateSheet.Range("L2").value = "[SYS_LINE_FIELD_01]"
    NewTemplateSheet.Range("M2").value = "[SYS_LINE_FIELD_02]"
    NewTemplateSheet.Range("N2").value = "[SYS_LINE_FIELD_03]"
    NewTemplateSheet.Range("O2").value = "[SYS_LINE_FIELD_04]"
    NewTemplateSheet.Range("P2").value = "[SYS_LINE_FIELD_05]"
    NewTemplateSheet.Range("Q2").value = "[SYS_LINE_FIELD_06]"
    NewTemplateSheet.Range("R2").value = "[SYS_LINE_FIELD_07]"
    NewTemplateSheet.Range("S2").value = "[SYS_LINE_FIELD_08]"
    NewTemplateSheet.Range("T2").value = "[SYS_LINE_FIELD_09]"
    NewTemplateSheet.Range("U2").value = "[SYS_LINE_FIELD_10]"
    NewTemplateSheet.Range("V2").value = "[SYS_LINE_FIELD_11]"
    NewTemplateSheet.Range("W2").value = "[SYS_LINE_FIELD_12]"
    NewTemplateSheet.Range("X2").value = "[SYS_LINE_FIELD_13]"
    NewTemplateSheet.Range("Y2").value = "[SYS_LINE_FIELD_14]"
    NewTemplateSheet.Range("Z2").value = "[SYS_LINE_FIELD_15]"
    NewTemplateSheet.Range("AA2").value = "[SYS_LINE_FIELD_16]"
    NewTemplateSheet.Range("AB2").value = "[SYS_LINE_FIELD_17]"

    ' ==========================================
    ' Friendly column labels (User-facing)
    ' ==========================================
    NewTemplateSheet.Range("A3").value = "Record Type" ' Kept to maintain the H/L logic context
    NewTemplateSheet.Range("B3").value = "[Header Description 01]"
    NewTemplateSheet.Range("C3").value = "[Header Description 02]"
    NewTemplateSheet.Range("D3").value = "[Header Description 03]"
    NewTemplateSheet.Range("E3").value = "[Header Description 04]"
    NewTemplateSheet.Range("F3").value = "[Header Description 05]"
    NewTemplateSheet.Range("G3").value = "[Header Description 06]"
    NewTemplateSheet.Range("H3").value = "[Header Description 07]"
    NewTemplateSheet.Range("I3").value = "[Header Description 08]"
    NewTemplateSheet.Range("J3").value = "[Header Description 09]"
    NewTemplateSheet.Range("K3").value = "[Header Description 10]"

    NewTemplateSheet.Range("L3").value = "[Line Item Description 01]"
    NewTemplateSheet.Range("M3").value = "[Line Item Description 02]"
    NewTemplateSheet.Range("N3").value = "[Line Item Description 03]"
    NewTemplateSheet.Range("O3").value = "[Line Item Description 04]"
    NewTemplateSheet.Range("P3").value = "[Line Item Description 05]"
    NewTemplateSheet.Range("Q3").value = "[Line Item Description 06]"
    NewTemplateSheet.Range("R3").value = "[Line Item Description 07]"
    NewTemplateSheet.Range("S3").value = "[Line Item Description 08]"
    NewTemplateSheet.Range("T3").value = "[Line Item Description 09]"
    NewTemplateSheet.Range("U3").value = "[Line Item Description 10]"
    NewTemplateSheet.Range("V3").value = "[Line Item Description 11]"
    NewTemplateSheet.Range("W3").value = "[Line Item Description 12]"
    NewTemplateSheet.Range("X3").value = "[Line Item Description 13]"
    NewTemplateSheet.Range("Y3").value = "[Line Item Description 14]"
    NewTemplateSheet.Range("Z3").value = "[Line Item Description 15]"
    NewTemplateSheet.Range("AA3").value = "[Line Item Description 16]"
    NewTemplateSheet.Range("AB3").value = "[Line Item Description 17]"

    ' Draw borders
    With NewTemplateSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Change colors, font bold
    With NewTemplateSheet.Range("A1:K3")
        .Interior.Color = RGB(171, 247, 211)
        .Font.Color = RGB(0, 122, 55)
        .Font.Bold = True
    End With

    With NewTemplateSheet.Range("L1:AB3")
        .Interior.Color = RGB(255, 229, 117)
        .Font.Color = RGB(116, 94, 0)
        .Font.Bold = True
    End With

    ' ==========================================
    ' Data Mapping & Routing
    ' ==========================================
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    NewTemplateSheet.Cells(4, 1).value = "H" ' Adds "H" (Header flag) in cell A4
    
    For i = 2 To lastRow
        NewTemplateSheet.Cells(i + 3, 1).value = "L"          ' Adds "L" (Line flag) in column A starting from row 5
        NewTemplateSheet.Cells(i + 3, 12).value = "'" & ws.Cells(i, 1).Text  ' Col A -> Col L
        NewTemplateSheet.Cells(i + 3, 13).value = "'" & ws.Cells(i, 2).Text  ' Col B -> Col M
        NewTemplateSheet.Cells(i + 3, 14).value = "'" & ws.Cells(i, 3).Text  ' Col C -> Col N
        NewTemplateSheet.Cells(i + 3, 15).value = "'" & ws.Cells(i, 4).Text  ' Col D -> Col O
        NewTemplateSheet.Cells(i + 3, 16).value = "'" & ws.Cells(i, 5).Text  ' Col E -> Col P
        NewTemplateSheet.Cells(i + 3, 17).value = "'" & ws.Cells(i, 6).Text  ' Col F -> Col Q
        NewTemplateSheet.Cells(i + 3, 18).value = "'" & ws.Cells(i, 7).Text  ' Col G -> Col R
        NewTemplateSheet.Cells(i + 3, 19).value = "'" & ws.Cells(i, 8).Text  ' Col H -> Col S
        NewTemplateSheet.Cells(i + 3, 20).value = "'" & ws.Cells(i, 9).Text  ' Col I -> Col T
        NewTemplateSheet.Cells(i + 3, 28).value = "'" & ws.Cells(i, 10).Text ' Col J -> Col AB
        NewTemplateSheet.Cells(i + 3, 26).value = ws.Cells(i, 11).Text       ' Col K -> Col Z
    Next i

    ' Optimize column widths on the new sheet
    NewTemplateSheet.Columns.AutoFit

    MsgBox "Template conversion executed successfully!", vbInformation, "Success"
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
