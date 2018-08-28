class ExcelObject {
    [string] $filepath;
    [object] $excelDoc;
    $workbook;
    
    # Init blank document
    ExcelObject() {
        $this.cleanup();
        $this.excelDoc = $this.openExcelDoc($this.filepath);
    }

    # Initialize object based on filepath
    ExcelObject([string] $filepath) {
        $this.cleanup();
        $this.filepath = $filepath;
        $this.excelDoc = $this.newExcelDoc("new-employees.xlsx");
        $this.addWorkbook();
        $this.renameWorksheet();
        $this.updateCell();
    }

    # Get Excel object ready
    [object] initExcel() {
        return NEW-Object -comobject Excel.Application;
    }

    # Create a new Excel document
    [object] newExcelDoc([string]$filename) {
        return $this.initExcel();
    }


    # Open Excel document using Excel object
    [object] openExcelDoc([string]$filepath) {
        $file = $this.initExcel();
        return $file.workbooks.open($filepath);
    }
    
    [string] readCell($col, [int]$row) {
        $worksheet=$this.excelDoc.Activesheet;
        return $worksheet.Range("$col$row").text;
    }

    addWorkbook() {
        $this.workbook = $this.excelDoc.Workbooks.Add();
        $this.workbook.Worksheets.Item(3).Delete();
    }

    renameWorksheet() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        $uregwksht.Name = 'The name you choose';
    }

    updateCell() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        $row = 1 
        $Column = 1 
        $uregwksht.Cells.Item($row,$column)= 'Title'
    }

    mergeCells() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        #merging a few cells on the top row to make the title look nicer 
        $MergeCells = $uregwksht.Range("A1:G1") 
        $MergeCells.Select() 
        $MergeCells.MergeCells = $true 
        $uregwksht.Cells(1, 1).HorizontalAlignment = -4108
    }

    formatCells() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        # If you want to give a nicer format to your title ( a specific font , height, ... ) can be done in this way.
        $uregwksht.Cells.Item(1,1).Font.Size = 18 
        $uregwksht.Cells.Item(1,1).Font.Bold=$True 
        $uregwksht.Cells.Item(1,1).Font.Name = "Cambria" 
        $uregwksht.Cells.Item(1,1).Font.ThemeFont = 1 
        $uregwksht.Cells.Item(1,1).Font.ThemeColor = 4 
        $uregwksht.Cells.Item(1,1).Font.ColorIndex = 55 
        $uregwksht.Cells.Item(1,1).Font.Color = 8210719
    }

    createColumnHeaders() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        #create the column headers 
        $uregwksht.Cells.Item(3,1) = 'Date';
        $uregwksht.Cells.Item(3,2) = 'Hour';
        $uregwksht.Cells.Item(3,3) = 'Name';
    }

    loadCsvData([string] $sourcepath) {
        $records = Import-Csv -Path $sourcepath;

        #seeing i used row 1 for the title then left a blank row & use row 3 for the column headers 
        # i chose to start with the data from row 4 hence the $i is set to 4 
        $i = 4

        # the .appendix to $record refers to the column header in the csv file 
        foreach($record in $records) 
        { 
            $this.excelDoc.cells.item($i,1) = $record.date;
            $this.excelDoc.cells.item($i,2) = $record.hour;
            $this.excelDoc.cells.item($i,3) = $record.name;
            $i++;
        }
    }

    autosizeColumns() {
        $uregwksht= $this.workbook.Worksheets.Item(1);
        #adjusting the column width so all data's properly visible 
        $usedRange = $uregwksht.UsedRange;
        $usedRange.EntireColumn.AutoFit() | Out-Null;
    }

    # Using a word document as a template
    [void] replaceWord([string]$target, [string]$word) {
        $this.searchWord($target, $word);
    }

    [void] print() {
        $this.excelDoc.printout();
    }
    [void] save() {
        $this.excelDoc.save();
        $this.excelDoc.close();
        $this.excelDoc.quit();
    }

    [void] save([string]$fileName) {
        $this.excelDoc.Saveas($filename);
        $this.excelDoc.close();
        $this.excelDoc.quit();
    }

    [void] cleanup() {
        Stop-Process -Name "EXCEL";
    }
}
