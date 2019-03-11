function Get-ExcelCurrentSelection {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Microsoft.Office.Interop.Excel.ApplicationClass]
        $App,

        [Parameter(Mandatory=$False)]
        [ValidateNotNullOrEmpty()]
        [String]
        $StartCell = "A1"      
    )
    
    begin {

    }
    
    process {

        # returns true if successful
        # http://www.informit.com/articles/article.aspx?p=2021718&seqNum=14
        # https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet.range
        $App.ActiveWorkbook.ActiveSheet.Range($StartCell).CurrentRegion.Select()
        #ActiveCell.Value
    }
    
    end {

    }
}

