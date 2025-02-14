# Excel Checklist Automation with VBA

This repository contains VBA macros that automate various tasks in Excel, specifically for creating and managing checklists. The macros allow for the following functionalities:

1. **Linking Checkboxes to Cells**
2. **Centering Checkboxes in Cells**
3. **Setting Row Height**

## VBA Script

```vb
Sub LinkCheckBoxes()

    Dim chk As CheckBox
    Dim lcol As Long

    lcol = 0

    ' Loop through all the checkboxes on the active sheet
    For Each chk In ActiveSheet.CheckBoxes
        With chk
            ' Link the checkbox to the cell to the right of it (adjusted by lcol)
            .LinkedCell = .TopLeftCell.Offset(0, lcol).Address
        End With
    Next chk

End Sub


Sub CenterCheckBoxes()

    Dim chk As CheckBox
    Dim cell As Range

    ' Loop through all the checkboxes in the active sheet
    For Each chk In ActiveSheet.CheckBoxes
        Set cell = chk.TopLeftCell
        
        ' Adjust the position of the checkbox to center it in the cell
        chk.Left = cell.Left + (cell.Width - chk.Width) / 2
        chk.Top = cell.Top + (cell.Height - chk.Height) / 2
    Next chk

End Sub


Sub SetRowHeight()

    Dim rowHeight As Double
    rowHeight = 18 ' Set the desired row height (in points)

    ' For all rows in the active sheet
    Rows.RowHeight = rowHeight
    
    ' OR, for specific rows (e.g., rows 1 to 10)
    ' Rows("1:10").RowHeight = rowHeight

End Sub
```
****************************************************************************************************************************************
# Formula for 'Progress' column
```
=IF(COUNTA(K12B5:F2) = 0, "", REPT("|", COUNTIF(B2:F2, TRUE) / COUNTA(B2:F2) * 100) & " " & TEXT(COUNTIF(B2:F2, TRUE) / COUNTA(B2:F2), "0%"))
```
Conditional Formulas Based on Progress Ranges
These formulas check the percentage of TRUE values in the range B2:F2 and return TRUE or FALSE based on predefined percentage ranges.
```
=COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) = 0
```
```
=AND(COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) > 0, COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) <= 0.25)
```
```
=AND(COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) > 0.25, COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) <= 0.5)
```
```
=AND(COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) > 0.5, COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) <= 0.8)
```
```
=COUNTIF(B2:F2, TRUE)/COUNTA(B2:F2) > 0.81
```
