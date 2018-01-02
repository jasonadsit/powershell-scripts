function Get-DnsClientCache {

    [CmdletBinding()]

    param (

        [string]
        $IpConfigPath = 'C:\Windows\System32\ipconfig.exe',

        [string]
        $IpConfigArgs = '/displaydns',

        [string]
        $TempFile = (Join-Path -Path $env:TEMP -ChildPath $(([System.Guid]::NewGuid().Guid) + '.txt'))

    ) #param

    process {

        $SplatArgs = @{ FilePath = $IpConfigPath
                        ArgumentList = $IpConfigArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }

        Start-Process @SplatArgs

        $IpConfigOutput = Get-Content -Path $TempFile

        Remove-Item -Path $TempFile | Out-Null

        $DnsClientCache = @()

        $IpConfigOutput | Select-String -Pattern "Record Name" -Context 0,5 | ForEach-Object {

            $Record = New-Object -TypeName psobject -Property @{

                Name = (($_.Line -split ':')[1]).Trim()
                Type = (($_.Context.PostContext[0] -split ':')[1]).Trim()
                TTL = (($_.Context.PostContext[1] -split ':')[1]).Trim()
                Length = (($_.Context.PostContext[2] -split ':')[1]).Trim()
                Section = (($_.Context.PostContext[3] -split ':')[1]).Trim()
                HostRecord = (($_.Context.PostContext[4] -split ':')[1]).Trim()

            } #New-Object

            $DnsClientCache += $Record

        } #ForEach

        Write-Output -InputObject $DnsClientCache

    } #process

} #function Get-DnsClientCache