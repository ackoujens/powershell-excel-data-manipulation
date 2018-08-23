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
    $wordObject.replaceWord("***firstName***", $firstName);
    $wordObject.replaceWord("***lastName***", $lastName);

    # Team
    $team=ReadCellData -Workbook $Workbook -Cell "C$Row"
    $teamNum=$team.Split(" ")[0]
    $teamName=$team.Split(" ")[1]
    $wordObject.replaceWord("***teamNum***", $teamNum);
    $wordObject.replaceWord("***teamName***", $teamName);

    # Delivery details
    $deviceNum=ReadCellData -Workbook $Workbook -Cell "M$Row"
    $sessionDate=ReadCellData -Workbook $Workbook -Cell "T$Row"
    $wordObject.replaceWord("***deviceNum***", $deviceNum);
    $wordObject.replaceWord("***sessionDate***", $sessionDate);

    # TODO Replace empty space in signature field with actual signature

    # Print
    $wordObject.print();

    # Save
    $path = $PSScriptRoot + "\output\" + $FirstName + "-" + $LastName + ".docx";
    # FIXME: Removing this line makes the script fail, complaining there is no method "save"
    $wordObject.save($path);

    $Row++

    # DEBUG
    write-host "ROW: " $Row - "FIRST NAME: " $FirstName - "LAST NAME: " $LastName
  }

} while ($Row -ne 2)
SaveExcelBook -workbook $Workbook
