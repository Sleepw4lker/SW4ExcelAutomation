function Get-ExcelTableStyles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App   
    )
    
    begin {

    }
    
    process {

        # https://docs.microsoft.com/de-de/office/vba/api/excel.styles
        $App.ActiveWorkbook.TableStyles | ForEach-Object { 
            $_.Name
        }
    }
    
    end {
    }
}