function Format-VersionHistoryTable {
    [cmdletbinding()]
    param (
        [Object] $TableObject
    )

    [String[]] $channels = $tableobject[0].psobject.Properties.Name | Where-Object { $_ -like "*Channel*" }
    [String] $lastYear = $null

    foreach ($row in $TableObject) {

        if ($row.Year){ 
            $lastYear = $row.Year
        } else {
            $row.Year = $lastYear
        }

        $relaseDateAry = $row.'Release date'.Split(' ')
        $month = [array]::indexof([cultureinfo]::CurrentCulture.DateTimeFormat.MonthNames, $relaseDateAry[0]) + 1
        $day = $relaseDateAry[1]
        $releaseDateStr = "$($day.PadLeft(2,'0'))-$($month.ToString().PadLeft(2,'0'))-$($row.Year)"
        $releaseDate = [datetime]::parseexact($releaseDateStr, 'dd-MM-yyyy', $null)

        foreach ($channel in $channels){
            $versionBuilds = $row.$channel.split(')')
            $release = 0
            foreach ($vb in $versionBuilds){
                $tempAry = $vb.Split('(')
                if ($tempAry[1]){
                    $channelBuild = $tempAry[1].Replace('Build ','').Replace(')','')
                    [PSCustomObject]([ordered]@{
                        Channel = $channel;
                        ChannelShortName = $channel.Replace('Channel', '').Replace('Enterprise', '').Replace('-', '').Replace('(', '').Replace(')', '').Replace(' ', '');
                        Version = $tempAry[0].Replace('Version ', '').Trim()
                        ChannelBuild = $channelBuild
                        FullBuild = ([version] "16.0.$($channelBuild)")
                        ReleaseDate = $releaseDate
                        Release = $release
                    })
                    $release++
                }              
            }
        }
    }
}
