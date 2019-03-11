function Set-ExcelCellTableStyle {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App,

        [Parameter(Mandatory=$True)]
        [ValidateScript({$(Get-ExcelTableStyles) -contains $_})]
        [String]
        $TableStyle,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $StartCell = "A1"
    )
    
    begin {

    }
    
    process {

        # https://stackoverflow.com/questions/21557916/excel-macro-select-all-cells-with-data-and-format-as-table
        # https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba
        # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xllistobjectsourcetype?view=excel-pia
        # https://docs.microsoft.com/en-us/office/vba/api/excel.listobjects.add
        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlcelltype
        
        $Range1 = $App.ActiveWorkbook.ActiveSheet.Range($StartCell)
        $Range2 = $Range1.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeLastCell)
        $Range3 = $App.Range($Range1, $Range2)
        
        $ListObject = $App.ActiveSheet.ListObjects.Add(
            [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
            $Range3, 
            [Type]::Missing, 
            [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
        )
        $ListObject.TableStyle = $TableStyle

    }
    
    end {
    }
}