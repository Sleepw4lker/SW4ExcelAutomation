function Set-ExcelColumnAutoFit {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App
    )
    
    begin {

    }
    
    process {

        # returns true if successful
        # https://www.thespreadsheetguru.com/the-code-vault/2014/3/25/vba-code-to-autofit-columns
        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlcelltype
        $App.ActiveWorkbook.Worksheets | Foreach-Object {
            [void]($_.Cells.Specialcells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeVisible).Entirecolumn.Autofit())
        }
        
    }
    
    end {
    }
}