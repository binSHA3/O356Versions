#=================================================
#   Import Functions
#   https://github.com/RamblingCookieMonster/PSStackExchange/blob/master/PSStackExchange/PSStackExchange.psm1
#=================================================
$PublicFunctions  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue )
$PrivateFunctions = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -Recurse -ErrorAction SilentlyContinue )

foreach ($Import in @($PublicFunctions + $PrivateFunctions)) {
    Try {
        . $Import.FullName
    }
    Catch {
        Write-Error -Message "Failed to import function $($Import.FullName): $_"
    }
}

Export-ModuleMember -Function $OSDPublicFunctions.BaseName
# New-Alias -Name Get-OSBuilder -Value Get-OSDBuilder -Force -ErrorAction SilentlyContinue
# New-Alias -Name New-OSBMediaISO -Value New-OSDBuilderISO -Force -ErrorAction SilentlyContinue
# New-Alias -Name New-OSBMediaUSB -Value New-OSDBuilderUSB -Force -ErrorAction SilentlyContinue
# New-Alias -Name Show-OSBMediaInfo -Value Show-OSDBuilderInfo -Force -ErrorAction SilentlyContinue
# New-Alias -Name Get-DownOSDBuilder -Value Save-OSDBuilderDownload -Force -ErrorAction SilentlyContinue

#=================================================
#   Export-ModuleMember
#=================================================
Export-ModuleMember -Function * -Alias *