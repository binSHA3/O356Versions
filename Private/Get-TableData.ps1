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
