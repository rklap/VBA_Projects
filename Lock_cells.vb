Sub ChangeToAbsoluteReference()
    Dim cell As Range
    
    ' Turn off calculations and screen updating
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ' Loop through each cell in the selection
    For Each cell In Selection
        ' Check if the cell contains a formula
        If cell.HasFormula Then
            ' Change the formula to absolute reference
            cell.Formula = Application.ConvertFormula(cell.Formula, xlA1, xlA1, xlAbsolute)
        End If
    Next cell
    
    ' Turn calculations and screen updating back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
