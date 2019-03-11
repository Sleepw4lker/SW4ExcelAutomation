<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Close-ExcelWorkbook {

    # You must pass a "Excel.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App
    )

    process {

        Write-Verbose -Message "Closing current Workbook"

        # Check version of Excel installed and discard changes
        If ($(Get-ExcelVersion) -eq 14) {
            $App.ActiveWorkbook.Close([ref]$False)
        }
        Else {
            # Office 2013 or newer
            $App.ActiveWorkbook.Close($False)  
        }

    }

}