<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Get-ExcelVersion {

    [cmdletbinding()]
    param()

    process {

        # We both check if Excel is installed at all and if so, which Version
        Try { 
            $ExcelVersion = (Get-ItemProperty HKLM:\Software\Classes\Excel.Application\CurVer).'(default)'
        }
        Catch {
            Throw "Excel seems not to be installed"
        }

        $ExcelVersion = $ExcelVersion.TrimStart("Excel.Application.")
        [int]$ExcelVersion = [convert]::ToInt32($ExcelVersion, 10)

        $ExcelVersion

    }

}