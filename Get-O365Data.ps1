#Requires -modules htmlagilitypack

function main {
    $url = 'https://docs.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'
    $supVerHeaderXpth = '/html[1]/body[1]/div[2]/div[1]/section[1]/div[1]/div[1]/main[1]/div[3]/table[1]/thead[1]/tr'
    $supVerRowsXpth = '/html[1]/body[1]/div[2]/div[1]/section[1]/div[1]/div[1]/main[1]/div[3]/table[1]/tbody[1]/tr'
    $verHistHeaderXpth = '/html[1]/body[1]/div[2]/div[1]/section[1]/div[1]/div[1]/main[1]/div[3]/table[2]/thead[1]/tr'
    $verHistRowsXpth = '/html[1]/body[1]/div[2]/div[1]/section[1]/div[1]/div[1]/main[1]/div[3]/table[2]/tbody[1]/tr'

    $HtmlObj = ConvertFrom-Html (Get-WebHtml -Url $url)
    $SupportedVersions = Get-TableData -HtmlObj $HtmlObj -HeaderXpath $supVerHeaderXpth -RowXpath $supVerRowsXpth
    $VersionHistory = Get-TableData -HtmlObj $HtmlObj -HeaderXpath $verHistHeaderXpth -RowXpath $verHistRowsXpth
    Confirm-TableData -VersionHistory $VersionHistory -SupportedVersions $SupportedVersions
    $FormattedSupportedVersions = Format-SupportedVersionsTable -TableObject $SupportedVersions
    $FormattedVersionHistory = Format-VersionHistoryTable -TableObject $VersionHistory | Sort-Object ReleaseDateDate, Channel, FullBuild

    $finalObject = ([PSObject] ([ordered]@{
        CreationDate = Get-Date
        VersionHistory = $FormattedVersionHistory | Select-Object -Property *
        SupportedVersions = ($FormattedSupportedVersions | Select-Object -Property * -ExcludeProperty @('Build','End of Service','Latest release date','Version availability date'))
    }))

    $finalObject.VersionHistory | ogv
       
}

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

function Get-WebHtml {
    param (
        [string] $Url
    )

    Add-Type -AssemblyName System.Net.Http
    $httpClient = New-Object System.Net.Http.HttpClient
    try {
        $task = $httpClient.GetAsync($url)
        $task.wait()
        $res = $task.Result

        if ($res.isFaulted) {
            write-host $('Error: Status {0}, reason {1}.' -f [int]$res.Status, $res.Exception.Message)
        }
        return $res.Content.ReadAsStringAsync().Result
    } catch {
        write-host ('Error: {0}' -f $_)
    } finally {
        if($null -ne $res){
            $res.Dispose()
        }
    }   
}


function Get-TableData {
    [cmdletbinding()]
    param (
        [HtmlAgilityPack.HtmlNode] $HtmlObj,
        [String] $HeaderXpath,
        [String] $RowXpath
    )

    $headers = $HtmlObj.SelectNodes($HeaderXpath).InnerText.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
    $rows = $HtmlObj.SelectNodes($RowXpath) 

    $tblData =  foreach ($row in $rows) {
        $elements = $row.ChildNodes | ? NodeType -eq 'Element' | Select -ExpandProperty InnerText
        $rowHsh = [ordered]@{}
        for ($i=0; $i -lt $headers.Count; $i++) {
            $rowHsh.Add($headers[$i],$elements[$i])
        }
            
        [PSCustomObject]($rowHsh)
    }

    return $tblData
}

main
