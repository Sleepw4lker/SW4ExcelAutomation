<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Open-ExcelWorkbookAsText {

    # You must pass a "Excel.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App,

        # To-Do: Verify against allowed Extensions
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File,

        [Parameter(Mandatory=$False)]
        [Switch]
        $ReadOnly = $False
    )

    process {

        Write-Verbose -Message "Opening Workbook $File. Read-Only: $ReadOnly"
        
        # Arrrrghhhh
        Start-Sleep -Seconds 1

        # Use [Type]::Missing for parameters that you want used with their default value.
        #$DefaultValue = [Type]::Missing

        # https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.opentext
        [void]$App.Workbooks.OpenText(
            $File
        )

    }

}