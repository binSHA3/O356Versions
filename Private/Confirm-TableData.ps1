function Confirm-TableData {
    param (
        [Parameter(Mandatory)]
        [Object] $VersionHistory,

        [Parameter(Mandatory)]
        [Object] $SupportedVersions
    )

    $ExpectVerHisHeaders = @('Current Channel','Monthly Enterprise Channel','Release date','Semi-Annual Enterprise Channel','Semi-Annual Enterprise Channel (Preview)','Year')
    $ExpectedSupVerHeaders = @('Build','Channel','End of service','Latest release date','Version','Version availability date')

    $CurVerHisHeaders = $VersionHistory |
        Get-Member |
        Where-Object {$_.MemberType -eq 'NoteProperty'} |
        Select-Object -ExpandProperty Name

    $CurSupVerHeaders = $SupportedVersions |
        Get-Member |
        Where-Object {$_.MemberType -eq 'NoteProperty'} |
        Select-Object -ExpandProperty Name

    #If arrays equal, will return $null
    $supVerCompare = Compare-Object -ReferenceObject $CurSupVerHeaders -DifferenceObject $ExpectedSupVerHeaders
    $missingHeaders = ($supVerCompare | Where-Object {$_.SideIndicator -eq '=>'} | Select -ExpandProperty InputObject) -Join(',')
    $extraHeaders = ($supVerCompare | Where-Object {$_.SideIndicator -eq '<='} | Select -ExpandProperty InputObject) -Join(',')
    if ($supVerCompare){
        Write-Host "ERROR Unexpected headers in Supporting Versions table:"
        Write-Host "`tMissing headers: $missingHeaders"
        Write-Host "`tExtra headers: $extraHeaders"
    } else {
        Write-Host """Supporting Versions"" table headers are as expected."
    } 
}