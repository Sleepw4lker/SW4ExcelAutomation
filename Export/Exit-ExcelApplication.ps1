<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Exit-ExcelApplication {

    # You must pass a "Excel.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]

        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App
    )

    process {

        If ($App.Application.Documents.Count -gt 0) {
            Close-ExcelWorkbook -App $App
        }

        Write-Verbose -Message "Exiting Excel Application"

        $App.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($App)

    }

}