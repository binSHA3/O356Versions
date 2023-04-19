function Get-O365VersionHistory {
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
                CreationDate      = Get-Date
                VersionHistory    = $FormattedVersionHistory | Select-Object -Property *
                SupportedVersions = ($FormattedSupportedVersions | Select-Object -Property * -ExcludeProperty @('Build', 'End of Service', 'Latest release date', 'Version availability date'))
            }))

    return $finalObject.VersionHistory
}

