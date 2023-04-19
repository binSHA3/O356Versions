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
