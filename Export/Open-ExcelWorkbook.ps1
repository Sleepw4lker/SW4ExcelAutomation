<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function Open-ExcelWorkbook {

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
        $DefaultValue = [Type]::Missing

        # https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
        [void]$App.Workbooks.Open(
            $File,	                # FileName. String. The file name of the workbook to be opened.
            $DefaultValue,	        # UpdateLinks. Specifies the way external references (links) in the file, 
                                    # such as the reference to a range in the Budget.xls workbook in the following 
                                    # formula =SUM([Budget.xls]Annual!C10:C25), are updated. 
            $ReadOnly.IsPresent     # ReadOnly.  	True to open the workbook in read-only mode.

        )

    }

}