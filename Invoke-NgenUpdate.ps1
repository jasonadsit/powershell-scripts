function Invoke-NgenUpdate {

    [CmdletBinding()]

    param(

        [string]
        $NgenPath32v4 = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe',

        [string]
        $NgenPath64v4 = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\ngen.exe',

        [string]
        $NgenPath32v2 = 'C:\Windows\Microsoft.NET\Framework\v2.0.50727\ngen.exe',

        [string]
        $NgenPath64v2 = 'C:\Windows\Microsoft.NET\Framework64\v2.0.50727\ngen.exe',

        [string]
        $NgenArgs = 'update /force',

        [string]
        $TempFile = (Join-Path -Path $env:TEMP -ChildPath $(([System.Guid]::NewGuid().Guid) + '.txt'))

    ) #param

    process {

        $SplatArgs = @{ FilePath = $NgenPath32v4
                        ArgumentList = $NgenArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }
        Start-Process @SplatArgs

        $SplatArgs = @{ FilePath = $NgenPath64v4
                        ArgumentList = $NgenArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }
        Start-Process @SplatArgs

        $SplatArgs = @{ FilePath = $NgenPath32v2
                        ArgumentList = $NgenArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }
        Start-Process @SplatArgs

        $SplatArgs = @{ FilePath = $NgenPath64v2
                        ArgumentList = $NgenArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }
        Start-Process @SplatArgs

        Remove-Item -Path $TempFile | Out-Null

    } #process

} #function Invoke-NgenUpdate