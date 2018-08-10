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

  # TODO Ommit or not to ommit? (Problem: var not used in code)
  # $read_only = $false;
  # $visible = $true;

  $replace = 2;
  $wrap = 1;
  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl) |out-null
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
$Workbook=OpenExcelBook -FileName "$PSScriptRoot\overzicht_nieuwe_medewerkers.xlsx"
$Row=2

do {
  write-host "Parsing row " $Row
  $name=ReadCellData -Workbook $Workbook -Cell "D$Row"

  if ($name.length -ne 0) {
    $Doc=OpenWordDoc -Filename "$PSScriptRoot\document-nieuwe-werknemer.docx"

    # Name
    $firstName=$name.Split(" ")[0]
    $lastName=$name.Substring($firstName.length+1)
    SearchAWord -Document $Doc -findtext '***firstName***' -replacewithtext $firstName
    SearchAWord -Document $Doc -findtext '***lastName***' -replacewithtext $lastName

    # Team
    $team=ReadCellData -Workbook $Workbook -Cell "C$Row"
    $teamNum=$team.Split(" ")[0]
    $teamName=$team.Split(" ")[1]
    SearchAWord -Document $Doc -findtext '***teamNum***' -replacewithtext $teamNum
    SearchAWord -Document $Doc -findtext '***teamName***' -replacewithtext $teamName

    # Delivery details
    $deviceNum=ReadCellData -Workbook $Workbook -Cell "M$Row"
    $sessionDate=ReadCellData -Workbook $Workbook -Cell "T$Row"
    SearchAWord -Document $Doc -findtext '***deviceNum***' -replacewithtext $deviceNum
    SearchAWord -Document $Doc -findtext '***sessionDate***' -replacewithtext $sessionDate

    # TODO Replace empty space in signature field with actual signature

    # Document creation
    $SaveName="$PSScriptRoot\output\$FirstName-$LastName.docx"
    SaveAsWordDoc -document $Doc -Filename $Savename

    # TODO Send to printer action

    $Row++

    # DEBUG
    write-host "ROW: " $Row - "FIRST NAME: " $FirstName - "LAST NAME: " $LastName
  }

} while ($Row -ne 10)
SaveExcelBook -workbook $Workbook
