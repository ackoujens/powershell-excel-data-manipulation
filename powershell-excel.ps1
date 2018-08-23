# Includes
. .\include\office\WordObject.ps1

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
    # Init WordObject
    $wordObject = [WordObject]::new("$PSScriptRoot\document-nieuwe-werknemer.docx");

    # Name
    $firstName=$name.Split(" ")[0]
    $lastName=$name.Substring($firstName.length+1)

    # SearchAWord -Document $Doc -findtext '***firstName***' -replacewithtext $firstName
    $wordObject.replaceWord("***firstName***", $firstName);
    
    # SearchAWord -Document $Doc -findtext '***lastName***' -replacewithtext $lastName
    $wordObject.replaceWord("***lastName***", $lastName);

    # Team
    $team=ReadCellData -Workbook $Workbook -Cell "C$Row"
    $teamNum=$team.Split(" ")[0]
    $teamName=$team.Split(" ")[1]
    
    # SearchAWord -Document $Doc -findtext '***teamNum***' -replacewithtext $teamNum
    $wordObject.replaceWord("***teamNum***", $teamNum);

    # SearchAWord -Document $Doc -findtext '***teamName***' -replacewithtext $teamName
    $wordObject.replaceWord("***teamName***", $teamName);

    # Delivery details
    $deviceNum=ReadCellData -Workbook $Workbook -Cell "M$Row"
    $sessionDate=ReadCellData -Workbook $Workbook -Cell "T$Row"
    
    # SearchAWord -Document $Doc -findtext '***deviceNum***' -replacewithtext $deviceNum
    $wordObject.replaceWord("***deviceNum***", $deviceNum);
    
    # SearchAWord -Document $Doc -findtext '***sessionDate***' -replacewithtext $sessionDate
    $wordObject.replaceWord("***sessionDate***", $sessionDate);

    # TODO Replace empty space in signature field with actual signature

    # Document creation
    # $SaveName="$PSScriptRoot\output\$FirstName-$LastName.docx"
    # SaveAsWordDoc -document $Doc -Filename $Savename
    $path = $PSScriptRoot + "\output\" + $FirstName + "-" + $LastName + ".docx";
    $wordObject.saveDocument($path);

    # TODO Send to printer action

    $Row++

    # DEBUG
    write-host "ROW: " $Row - "FIRST NAME: " $FirstName - "LAST NAME: " $LastName
  }

} while ($Row -ne 2)
SaveExcelBook -workbook $Workbook
