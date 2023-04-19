#=================================================
#   Import Functions
#   https://github.com/RamblingCookieMonster/PSStackExchange/blob/master/PSStackExchange/PSStackExchange.psm1
#=================================================
$PublicFunctions  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue )
$PrivateFunctions = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -Recurse -ErrorAction SilentlyContinue )

foreach ($Import in @($PublicFunctions + $PrivateFunctions)) {
    Try {
        Write-Host "Importing function $($Import.FullName)"
        . $Import.FullName
    }
    Catch {
        Write-Error -Message "Failed to import function $($Import.FullName): $_"
    }
}

Export-ModuleMember -Function $PublicFunctions.BaseName

#=================================================
#   Export-ModuleMember
#=================================================
Export-ModuleMember -Function * -Alias *