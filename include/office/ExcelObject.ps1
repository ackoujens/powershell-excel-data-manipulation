class ExcelObject {
    [string] $filepath;
    [object] $excelDoc;
    
    # Initialize object based on filepath
    ExcelObject([string] $filepath) {
        $this.cleanup();
        $this.filepath = $filepath;
        $this.excelDoc = $this.openExcelDoc($this.filepath);
    }

    # Get Word object ready
    [object] initExcel() {
        return NEW-Object -comobject Excel.Application;
    }
    
    # Open Word document using Word object
    [object] openExcelDoc([string]$filepath) {
        $file = $this.initExcel();
        return $file.workbooks.open($filepath);
    }
    
    [string] readCell($col, [int]$row) {
        $worksheet=$this.excelDoc.Activesheet;
        return $worksheet.Range("$col$row").text;
    }

    # Using a word document as a template
    [void] replaceWord([string]$target, [string]$word) {
        $this.searchWord($target, $word);
    }

    [void] print() {
        $this.excelDoc.printout();
    }
    [void] save() {
        $this.excelDoc.save()
        $this.excelDoc.close()
    }

    [void] save([string]$fileName) {
        $this.excelDoc.Saveas($filename);
        $this.excelDoc.close();
    }

    [void] cleanup() {
        Stop-Process -Name "EXCEL";
    }
}
