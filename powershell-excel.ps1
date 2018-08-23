# Reset vars
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

# Includes
. .\include\office\WordObject.ps1
. .\include\office\ExcelObject.ps1

# Open Excel sheet
$excelObject = [ExcelObject]::new("$PSScriptRoot\overzicht_nieuwe_medewerkers.xlsx");

$Row=2

do {
  write-host "------------------------"
  write-host "Parsing row" $row
  write-host "------------------------"

  $name=$excelObject.readCell("D", $row);
  
  if ($name.length -ne 0) {

    # Init WordObject
    $wordObject = [WordObject]::new("$PSScriptRoot\document-nieuwe-werknemer.docx");

    # Name
    $firstName=$name.Split(" ")[0]
    $lastName=$name.Substring($firstName.length+1)
    $wordObject.replaceWord("***firstName***", $firstName);
    $wordObject.replaceWord("***lastName***", $lastName);

    # Team
    $team=$excelObject.readCell("C", $row);
    $teamNum=$team.Split(" ")[0]
    $teamName=$team.Split(" ")[1]
    $wordObject.replaceWord("***teamNum***", $teamNum);
    $wordObject.replaceWord("***teamName***", $teamName);

    # Delivery details
    $deviceNum=$excelObject.readCell("M", $row);
    $sessionDate=$excelObject.readCell("T", $row);
    $wordObject.replaceWord("***deviceNum***", $deviceNum);
    $wordObject.replaceWord("***sessionDate***", $sessionDate);

    # Signature
    # TODO: Replace empty space in signature field with actual signature
    # TODO: Pick random from multiple signature files, located in "signatures" folder

    # Print
    $wordObject.print();

    # Save
    $path = $PSScriptRoot + "\output\" + $firstName + "-" + $lastName + ".docx";
    # FIXME: Removing this line makes the script fail, complaining there is no method "save"
    $wordObject.save($path);

    $Row++

    # DEBUG
    write-host "Row: "        $row
    write-host "First name: " $firstName
    write-host "Last name: "  $lastName
  }

} while ($name.length -ne 0)
$excelObject.save();