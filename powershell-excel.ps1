function OpenWordDoc($Filename) {
  $Word=NEW-Object -comobject Word.Application
  return $Word.documents.open($Filename)
}

function SearchAWord($Document,$findtext,$replacewithtext) {
  $FindReplace=$Document.ActiveWindow.Selection.Find
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
  $read_only = $false;
  $visible = $true;
  $replace = 2;
  $wrap = 1;
  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)
}

function SaveAsWordDoc($Document,$FileName) {
  $Document.Saveas([REF]$Filename)
  $Document.close()
}

function OpenExcelBook($FileName) {
  $Excel=new-object -ComObject Excel.Application
  return $Excel.workbooks.open($Filename)
}

function SaveExcelBook($Workbook) {
  $Workbook.save()
  $Workbook.close()
}

function ReadCellData($Workbook,$Cell) {
  $Worksheet=$Workbook.Activesheet
  return $Worksheet.Range($Cell).text
}

# Open Excel sheet
$Workbook=OpenExcelBook -FileName "$PSScriptRoot\test-sheet.xlsx"
$Row=2

do {
  $Data=ReadCellData -Workbook $Workbook -Cell "A$Row"

  if ($Data.length -ne 0) {
    $Doc=OpenWordDoc -Filename "$PSScriptRoot\test-document.docx"
    SearchAWord -Document $Doc -findtext '***FirstName***' -replacewithtext $Data

    $Data=ReadCellData -Workbook $Workbook -Cell "B$Row"
    SearchAWord -Document $Doc -findtext '***LastName***' -replacewithtext $Data

    $Data=ReadCellData -Workbook $Workbook -Cell "C$Row"
    SearchAWord -Document $Doc -findtext '***StreetAddress***' -replacewithtext $Data

    $Data=ReadCellData -Workbook $Workbook -Cell "D$Row"
    SearchAWord -Document $Doc -findtext '***City***' -replacewithtext $Data

    $Data=ReadCellData -Workbook $Workbook -Cell "E$Row"
    SearchAWord -Document $Doc -findtext '***State***' -replacewithtext $Data

    $Data=ReadCellData -Workbook $Workbook -Cell "F$Row"
    SearchAWord -Document $Doc -findtext '***Country***' -replacewithtext $Data

    $SaveName="$FirstName-$LastName.docx"
    SaveAsWordDoc -document $Doc -Filename $Savename

    $Row++
  }

  SaveExcelBook -workbook $Workbook
} while ($Data.length -ne 0)
