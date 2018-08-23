class WordObject {
    [string] $filepath;
    [object] $wordDoc;
    
    # Initialize object based on filepath
    WordObject([string] $filepath) {
        $this.cleanup();
        $this.filepath = $filepath;
        $this.wordDoc = $this.openWordDoc($this.filepath);
    }

    # Get Word object ready
    [object] initWord() {
        return NEW-Object -comobject Word.Application;
    }
    
    # Open Word document using Word object
    [object] openWordDoc([string]$filepath) {
        $file = $this.initWord();
        return $file.documents.open($filepath);
    }
    
    [void] searchWord([string]$target, [string]$word) {
        $FindReplace=$this.wordDoc.ActiveWindow.Selection.Find
        $matchCase = $false;
        $matchWholeWord = $true;
        $matchWildCards = $false;
        $matchSoundsLike = $false;
        $matchAllWordForms = $false;
        $forward = $true;
        $format = $false;
        $matchKashida = $false;
        $matchDiacritics = $false;
        $matchAlefHamza = $false;
        $matchControl = $false;
      
        # TODO Ommit or not to ommit? (Problem: var not used in code)
        # $read_only = $false;
        # $visible = $true;
      
        $replace = 2;
        $wrap = 1;
        $FindReplace.Execute($target, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $word, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl) |out-null
      }

    # Using a word document as a template
    [void] replaceWord([string]$target, [string]$word) {
        $this.searchWord($target, $word);
    }

    [void] saveDocument([string]$fileName) {
        $this.wordDoc.Saveas($filename);
        $this.wordDoc.close();
    }

    [void] cleanup() {
        Stop-Process -Name "WINWORD";
    }
}
