function Get-ExcelStyles {
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
        $App.ActiveWorkbook.Styles | ForEach-Object { 
            $_.Name
        }
        
    }
    
    end {
    }
}