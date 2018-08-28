# Reset vars
# TODO: Include in reset procedure
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host;

# Includes
. .\SQLObject.ps1

Class Employee {
    [SQLObject] $sqlObject;

    Employee() {
        $this.sqlObject = [SQLObject]::new("a253pclu02sql1\instomz", "WgkOvl_Interface");
    }

    [object] getEmployee([int] $employeeNumber) {
        $this.sqlObject.setQuery("SELECT * FROM Wgk_Werknemer WHERE PunteerNr = $employeeNumber");
        # write-host $result.Voornaam lives in $result.Gemeente;
        return $this.sqlObject.execQuery();
    }

    [object] getEmployee([string] $username) {
        $this.sqlObject.setQuery("SELECT * FROM Wgk_Werknemer WHERE Gebruikersnaam LIKE '%$username'");
        return $this.sqlObject.execQuery();
    }

    [object] getEmployees([string] $gemeente) {
        $this.sqlObject.setQuery("SELECT * FROM Wgk_Werknemer WHERE Gemeente LIKE '%$gemeente'");
        return $this.sqlObject.execQuery();
    }

}

$employee = [Employee]::new();










function Check-WGKInterface{
	[cmdletbinding()]
    PARAM(
        $EmployeeID,
        [Parameter(ValueFromPipeline)]
		$sAMAccountName
    )
    
# FIXME: a253pclu02sql1
# $SQLServer = "a253pclu01sql1\instomz"
$SQLServer = "a253pclu02sql1\instomz"

$SQLDBName = "WgkOvl_Interface"

if($sAMAccountName -ne $null){
$SqlQuery = @"
SELECT * FROM Wgk_Werknemer WHERE Gebruikersnaam LIKE '%$sAMAccountName'
"@
} else {
$SqlQuery = @"
SELECT * FROM Wgk_Werknemer WHERE PunteerNr = $EmployeeID
"@
}

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SQLDBName;Integrated Security = true;"

$Sql = New-Object System.Data.SqlClient.SqlCommand

$Sql.CommandText = $SqlQuery
$Sql.Connection = $SqlConnection


$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $Sql

$DataSet = New-Object System.Data.DataSet
$dSet = $SqlAdapter.Fill($DataSet)

$Result = $DataSet.Tables
$SqlConnection.Close()

$list = @()
$Result |ForEach-Object{
        $hash  = @{
            WerknemerID = $_.WerknemerID
            PunteerNr = $_.PunteerNr
            Naam = $_.Naam
            Voornaam = $_.Voornaam
            DatumInDienst = $_.DatumInDienst
            DatumUitDienst = $_.DatumUitDienst
            Geslacht = $_.Geslacht
            GeboorteDatum = $_.GeboorteDatum
            GeboortePlaats = $_.GeboortePlaats
            Straat = $_.Straat
            HuisNummer = $_.Huisnummer
            Bus = $_.Bus
            PostNummer = $_.PostNr
            Gemeente = $_.Gemeente
            PriveTelNr = $_.PriveTelefoonNr
            PriveGSMNr = $_.PriveGsmNr
            PriveFaxNr = $_.PriveFaxNr
            WGKGsmNr = $_.WGKGsmNr
            WGKEmail = $_.WGKEmail
            PriveEmail = $_.PriveEmail
            Gebruikersnaam = $_.Gebruikersnaam
            Adres = ""
            }
        if($_.Bus){
            $hash.Adres = $_.Straat + " " + $_.Huisnummer + " bus: " + $_.Bus + " - " + $_.PostNr + ", " + $_.Gemeente
        } else{
            $hash.Adres = $_.Straat + " " + $_.Huisnummer + " - " + $_.PostNr + ", " + $_.Gemeente
        }
        $list += New-Object psobject -Property $hash
    }
$list  #| select werknemerid, voornaam, naam, datumindienst, datumuitdienst
}






function Invoke-SQL {
    param(
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "MasterData",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables

}



function Get-AfdelingsCode{
    [cmdletbinding()]
    PARAM(
        $afdeling,
        $afdelingsnummer
    )
	$SQLServer = "A253PCLU01SQL3\INSTSHARED" 
	$SQLDBName = "Telecom"
	if($afdeling){
$SQLQuery = @"
SELECT CODE, ADRES, CODE_LOCATIE FROM WgkOvl_AfdAdm_Afdeling WHERE NAAM = '$afdeling'
"@
	}elseif($afdelingsnummer){
	$SQLQuery = @"
SELECT CODE, ADRES, CODE_LOCATIE FROM WgkOvl_AfdAdm_Afdeling WHERE AFD_NR = $afdelingsnummer
"@
	}


	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = true" #

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection

	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter

	$SqlCmd.CommandText = $SQLQuery
	$SqlAdapter.SelectCommand = $SqlCmd

	$DataSet = New-Object System.Data.DataSet
	$dSet = $SqlAdapter.Fill($DataSet)
	
	$SqlConnection.Close()
	
	$DataSet.Tables
	
}

function Check-UsersToDisable{
    <#
        .SYNOPSIS
        Check on users to disable in Active Directory
    
        .DESCRIPTION
        This wil check eXtend interface on users to be disabled
    
        .NOTES
        For examples type:
            Get-Help Check-usersToDisable -examples
    
        .EXAMPLE
        To get a list of users to be disabled today:
    
            Check-usersToDisable
    
        .EXAMPLE
        To get a list of users to be disabled on a certain date:
        
            Check-UsersToDisable -date 22/05/2017
    
        .EXAMPLE    
        To get a list of past week:
        
            Check-UsersToDisable -week
    
    #>
        [cmdletbinding()]
        PARAM(
            $date = (get-date).adddays(-1),
            [switch]$week
        )
        $SQLServer = "a253pclu02sql1\instomz"
        $SQLDBName = "WgkOvl_Interface"
    
        if($week){
        $dateFirst = "'" +  $(get-date($date) -Format "yyyy-MM-dd 00:00:00.000") + "'"
        $dateSecond = "'" +  $(get-date($date).adddays(-6) -Format "yyyy-MM-dd 00:00:00.000") + "'"
        $SqlQuery = "SELECT WerknemerID FROM Wgk_Werknemer WHERE DatumUitDienst BETWEEN $dateFirst AND $dateSecond"
        }else{
        $date = "'" + $(get-date($date) -Format "yyyy-MM-dd 00:00:00.000") + "'"
        $SqlQuery = "SELECT WerknemerID FROM Wgk_Werknemer WHERE DatumUitDienst = $date"
        }
        
    
    
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SQLDBName;Integrated Security = true;"
    
        $Sql = New-Object System.Data.SqlClient.SqlCommand
    
        $Sql.CommandText = $SqlQuery
        $Sql.Connection = $SqlConnection
    
    
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $Sql
    
        $DataSet = New-Object System.Data.DataSet
        $dSet = $SqlAdapter.Fill($DataSet)
    
        $Result = $DataSet.Tables
        $SqlConnection.Close()
        $Result
    }

    function Check-AccountsToExpire{
        [cmdletbinding()]
        PARAM(
            $script:date = (get-date).adddays(-1)
        )
        $SQLServer = "a253pclu02sql1\instomz"
        $SQLDBName = "WgkOvl_Interface"
    
        $date = "'" + $(get-date($date) -Format "yyyy-MM-dd 00:00:00.000") + "'"
        $SqlQueryWerknemer = "SELECT WerknemerID, Voornaam, Naam, DatumUitDienst FROM Wgk_Werknemer WHERE DatumUitDienst = $date"
        
        #start query werknemer
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$SQLDBName;Integrated Security = true;"
    
        $Sql = New-Object System.Data.SqlClient.SqlCommand
        $Sql.CommandText = $SqlQueryWerknemer
        $Sql.Connection = $SqlConnection
        
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $Sql
        
        $DataSetWerknemer = New-Object System.Data.DataSet
        try{
            $SqlAdapter.Fill($DataSetWerknemer) | out-null -ErrorAction Stop
        } catch {
            $result = $null
        }
        $SqlConnection.Close()
        #end query Werknemer
        
        #start query contractbeweging
        $list = $DataSetWerknemer.Tables.WerknemerID
        $IDs = ([string]$list).replace(" ",",")
        $SqlQueryContractBeweging = "SELECT DISTINCT WerknemerID FROM Wgk_ContractBeweging where WerknemerID in ($IDs) and (EindDatumContract is null or EindDatumContract > $date)"
    
        $Sql.CommandText = $SqlQueryContractBeweging
        $Sql.Connection = $SqlConnection
    
        $SqlAdapterContractBeweging = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $Sql
    
        $DataSetContractBeweging = New-Object System.Data.DataSet
        try{
            $SqlAdapter.Fill($DataSetContractBeweging) | Out-Null -ErrorAction Stop
        } catch {
            $result = $null
        }
        $SqlConnection.Close()
        #end queryContractBeweging
    
    
        $script:filePath = "C:\expireUsers$(get-date((get-date).adddays(-1)) -Format "yyyyMMdd").txt"
        
        $werknemers = New-Object System.Collections.ArrayList
        $DataSetWerknemer.tables.rows | ForEach-Object{$item = New-Object psobject
            $item | Add-Member NoteProperty -Name WerknemerId -Value $_.WerknemerId
            $item | Add-Member NoteProperty -Name Voornaam -Value $_.Voornaam
            $item | Add-Member NoteProperty -Name Naam -Value $_.Naam
            $item | Add-Member NoteProperty -Name DatumUitDienst -Value $(if($_.DatumUitDienst){get-date($_.DatumUitDienst) -Format "dd MMMM yyyy"})
            $item | Add-Member NoteProperty -Name AccountExpires -Value $(if($_.DatumUitDienst){get-date($_.DatumUitDienst).adddays(1) -Format "dd MMMM yyyy"})
            $Werknemers += $item
        }
    
        $bewegingen = New-Object System.Collections.ArrayList
            $DataSetContractBeweging.tables.rows | ForEach-Object{$item = New-Object psobject
            $item | Add-Member NoteProperty -Name WerknemerId -Value $_.WerknemerId
            $bewegingen += $item
        }
        $script:accounts = $werknemers | Where-Object {$_.werknemerid -ne $bewegingen.werknemerid}
        [string]$script:Result = $werknemers | Where-Object {$_.werknemerid -ne $bewegingen.werknemerid} | Format-Table -a | out-string
        
        
        if($Result){
            $($werknemers | Where-Object {$_.werknemerid -ne $bewegingen.werknemerid}).werknemerid | Out-File -FilePath $filePath -Encoding UTF8
            Return $Result
        }
    }