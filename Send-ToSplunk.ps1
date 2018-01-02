function Send-ToSplunk {

    [CmdletBinding()]

    param (

        [Parameter(ValueFromPipeline=$true)]
        [System.Object[]]
        $Message,

        [Parameter()]
        [string]
        $SplunkHecUri = 'http://splunk.example.com:8088/services/collector',

        [Parameter()]
        [string]
        $SplunkHecApiKey = 'DEADBEEF-DEAD-BEEF-DEAD-BEEFDEADBEEF',

        [Parameter()]
        [System.Collections.Hashtable]
        $SplunkHecRestHeaders = @{ Authorization = "Splunk $SplunkHecApiKey" }

    ) #param

    process {

        if (-not $Message) {

            $Message = New-Object -TypeName psobject -Property @{   ComputerName = $env:COMPUTERNAME
                                                                    UserName = $env:USERNAME
                                                                    Message = 'Hello World!' }

        } #if

        [datetime] $Epoch = (Get-Date -Date '01/01/1970')

        [datetime] $TimeNow = (Get-Date)

        [string] $EpochTime = $(((New-TimeSpan -Start $Epoch -End ([system.timezoneinfo]::ConvertTime(($TimeNow),([system.timezoneinfo]::UTC)))).TotalSeconds).ToString())

        $PowerShellVersion = $PSVersionTable.PSVersion.Major

        if ($PowerShellVersion -ge 3) {

            $Message | ForEach-Object {

                $JsonEvent = $_ | ConvertTo-Json -Compress

                $SplatArgs = @{ Uri = $SplunkHecUri
                                Headers = $SplunkHecRestHeaders
                                Method = 'Post'
                                Body = "{`"time`": `"$EpochTime`",`"host`": `"$($env:COMPUTERNAME)`",`"event`": $JsonEvent}" }

                Invoke-RestMethod @SplatArgs | Out-Null

            } #ForEach

        } elseif ($PowerShellVersion -lt 3) {

            Add-Type -AssemblyName System.Web

            $PsJs = New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer

            $Message | ForEach-Object {

                $JsonEvent = $PsJs.Serialize($_)

                [byte[]][char[]] $Body = "{`"time`": `"$EpochTime`",`"host`": `"$($env:COMPUTERNAME)`",`"event`": $JsonEvent}"

                $Request = [System.Net.HttpWebRequest]::CreateHttp("$SplunkHecUri")

                $Request.Method = 'POST'

                $Request.Headers.Add("Authorization","Splunk $SplunkHecApiKey")

                $Stream = $Request.GetRequestStream()

                $Stream.Write($Body, 0, $Body.Length)

                $Request.GetResponse() | Out-Null

            } #ForEach

        } #if

    } #process

} #function Send-ToSplunk