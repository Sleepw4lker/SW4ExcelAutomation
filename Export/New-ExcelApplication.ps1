<#
    .SYNOPSIS
        To Do: Documentation for this function
#>
Function New-ExcelApplication {

    # Returns a "Excel.Application" Object

    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [ValidateRange(8,64)]
        [int]
        $MinimumVersion
    )

    process {

        # We both check if Excel is installed at all and if so, which Version
        $ExcelVersion = Get-ExcelVersion

        If ($MinimumVersion -and $ExcelVersion -lt $MinimumVersion) {
            Throw "No compatible Version of Microsoft Excel installed. Needing $MinimumVersion, whereas the installed Version is $ExcelVersion"
        }

        Write-Verbose -Message "Spawning a new Excel Application Instance"

        Try {
            $App = New-Object -ComObject Excel.Application
        }
        Catch {
            Throw "Unable to open Microsoft Excel"
        }

        # Checking if the -Verbose Argument was given.
        # In this case, we also make the Application Window visible.
        If ([System.Management.Automation.ActionPreference]::SilentlyContinue -ne $VerbosePreference) {
            $App.Visible = $True
        }
        Else {
            $App.Visible = $False
        }

        # I hate such lousy Workarounds. But Excel seems to sometimes reject RPC Calls 
        # if we directly return the Object after launching the App.
        # This must be enough for now.
        Start-Sleep -Seconds 5

        $App
    }
}