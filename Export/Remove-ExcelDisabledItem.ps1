<#
    .SYNOPSIS
        To Do: Documentation for this function
#>

Function Remove-ExcelDisabledItem {

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path $_})]
        [String]
        $File
    )

    process {

        # Kudos to
        # https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-documents-in-the-ms-office-list-of-disabled-fil

        $ExcelVersion = Get-ExcelVersion

        # Converts the File Name string to UTF16 Hex
        $FileNameInUTF16Hex = ""
        [System.Text.Encoding]::ASCII.GetBytes($File.ToLower()) | ForEach-Object { 
            $FileNameInUTF16Hex += "{0:X2}00" -f $_
        }

        Try {
            # Tests to see if the Disabled items registry key exists
            $DisabledItemsRegistryKey = (Get-Item "HKCU:\Software\Microsoft\Office\${ExcelVersion}.0\Excel\Resiliency\DisabledItems\")
        }
        Catch {
            # Nothing yet
        }

        If ($NULL -ne $DisabledItemsRegistryKey) {

            # Cycles through all the properties and deletes it if it contains the file name.
            Foreach ($DisabledItem in $DisabledItemsRegistryKey.Property) {

                $Value = ""

                ($DisabledItemsRegistryKey | Get-ItemProperty).$DisabledItem | ForEach-Object {
                    $Value += "{0:X2}" -f $_
                }

                If ($Value.Contains($FileNameInUTF16Hex)) {

                    Write-Verbose "Removing $File from the List of Disabled Items."

                    $DisabledItemsRegistryKey | Remove-ItemProperty -name $DisabledItem

                }
            }
        }
    }
}