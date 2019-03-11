function Set-ExcelStyleWorkbook {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.Workbook]
        $SourceWorkbook,
        
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.Workbook]
        $TargetWorkbook
    )
    
    begin {

    }
    
    process {

        # https://docs.microsoft.com/en-us/office/vba/api/excel.styles.merge
        $SourceWorkbook.Styles.Merge($TargetWorkbook)

    }
    
    end {
    }
}