function Format-SupportedVersionsTable {
    [cmdletbinding()]
    param (
        [Object] $TableObject
    )

    foreach ($row in $TableObject) {
        $row | Add-Member -Name "VersionInt" -Value ([int] $row.Version) -MemberType NoteProperty
        $row | Add-Member -Name "FullBuild" -Value ([version] "16.0.$($row.'Build')") -MemberType NoteProperty
        $row | Add-Member -Name "LatestRelease" -Value ([datetime]::parseexact($row.'Latest Release Date', 'MMMM d, yyyy', $null)) -MemberType NoteProperty
        $row | Add-Member -Name "VersionAvailability" -Value ([datetime]::parseexact($row.'Version Availability Date', 'MMMM d, yyyy', $null)) -MemberType NoteProperty
        Try {
            $row | Add-Member -Name "EndOfService" -Value ([datetime]::parseexact($row.'End of service', 'MMMM d, yyyy', $null)) -MemberType NoteProperty
        } Catch {
            $row | Add-Member -Name "EndOfService" -Value $null -MemberType NoteProperty
        }
        $row | Add-Member -Name "ChannelShortName" -Value  $row.Channel.Replace('Channel', '').Replace('Enterprise', '').Replace('-', '').Replace('(', '').Replace(')', '').Replace(' ', '') -MemberType NoteProperty
        $row
    }
}