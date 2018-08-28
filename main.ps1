<#
 # New Employees Prep Generator
 #>

# Reset vars
# TODO: Include in reset procedure
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host;

# Subroutines
. compile-employee-data.ps1
. compile-device-documents.ps1