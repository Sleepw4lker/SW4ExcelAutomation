<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Save-ExcelWorkbook {

    # You must pass a "Excel.Application" Object

    [cmdletbinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App,

        # To-Do: Verify against allowed Extensions
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path ((New-Object System.IO.FileInfo $_).Directory.FullName)})]
        [String]
        $File
    )

    process {

        Write-Verbose -Message "Saving Workbook as $File"

        # https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
        $App.ActiveWorkbook.SaveAs(
            $File,
            [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
        )

    }

}