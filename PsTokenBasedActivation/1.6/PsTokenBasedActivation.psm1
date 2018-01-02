function Update-TokenBasedActivationLicence {
    <#
    .SYNOPSIS
        Updates token-based activation licences.
    .DESCRIPTION
        Uses a combination of PowerShell, VBScript, and unicorn tears to update the token-based activation licences.
    .PARAMETER ComputerName
        The computer name to operate on. Defaults to localhost (actually... `$env:COMPUTERNAME)
        Accepts values from the pipeline.
    .PARAMETER CacheMyPin
        Switch to allow PIN caching. Still a work in progress.
        Uncomment lines 140-164,317,320,323,335,338,341,370,373,376,388,390,392 to use.
    .PARAMETER LicenceRoot
        Location that xrm-ms licences are stored. Defaults to $PSScriptRoot.
    .PARAMETER Win7
        Licence file for Windows 7. Replace with a valid licence file.
    .PARAMETER Svr2008r2
        Licence file for Windows Server 2008 R2. Replace with a valid licence file.
    .PARAMETER Svr2012
        Licence file for Windows Server 2012. Replace with a valid licence file.
    .PARAMETER Win81
        Licence file for Windows 8.1. Replace with a valid licence file.
    .PARAMETER Svr2012r2
        Licence file for Windows Server 2012 R2. Replace with a valid licence file.
    .PARAMETER Win10
        Licence file for Windows 10. Replace with a valid licence file.
    .PARAMETER Office2013ProPlus
        Licence file for Microsoft Office 2013 Pro Plus. Replace with a valid licence file.
    .PARAMETER CScript
        Path to cscript.exe.
    .PARAMETER Slmgr
        Path to slmgr.vbs.
    .PARAMETER BackupSlmgr
        Backup path to slmgr.vbs.
    .PARAMETER SlmgrWin7
        Path to slmgr-win7.vbs.
    .PARAMETER BackupSlmgrWin7
        Backup path to slmgr-win7.vbs.
    .PARAMETER Ospp
        Path to ospp.vbs.
    .PARAMETER BackupOspp
        Backup path to ospp.vbs.
    .PARAMETER BackupOsppHtm
        Backup path to ospp.htm.
    .EXAMPLE
        Update-TokenBasedActivationLicence -ComputerName host1
        Updates the token-based activation licences on host1.
    .EXAMPLE
        'host1','host2','host3' | Update-TokenBasedActivationLicence -Verbose
        Updates the token-based activation licences on host1, host2, and host3 showing all the gory details.
    .INPUTS
        System.Object
    .OUTPUTS
        None
    .NOTES
        ###################################################################
        Author:     @oregon-national-guard/systems-administration
        Version:    1.6
        ###################################################################
    .LINK
        https://github.com/oregon-national-guard
    .LINK
        https://github.com/oregon-national-guard/powershell/blob/master/LICENCE
    .LINK
        https://creativecommons.org/publicdomain/zero/1.0/
    #>
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [string[]]
        $ComputerName,
        [switch]
        $CacheMyPin,
        [string]
        $LicenceRoot = $PSScriptRoot,
        [string]
        $Win7 = 'win7.xrm-ms',
        [string]
        $Svr2008r2 = '2008r2.xrm-ms',
        [string]
        $Svr2012 = '2012.xrm-ms',
        [string]
        $Win81 = 'win81.xrm-ms',
        [string]
        $Svr2012r2 = '2012r2.xrm-ms',
        [string]
        $Win10 = 'win10.xrm-ms',
        [string]
        $Office2013ProPlus = 'office2013proplus.xrm-ms',
        [string]
        $CScript = 'C:\Windows\System32\cscript.exe',
        [string]
        $Slmgr = 'C:\WINDOWS\System32\slmgr.vbs',
        [string]
        $BackupSlmgr = "$(Join-Path -Path $LicenceRoot -ChildPath 'slmgr.vbs')",
        [string]
        $SlmgrWin7 = 'C:\WINDOWS\System32\slmgr-win7.vbs',
        [string]
        $BackupSlmgrWin7 = "$(Join-Path -Path $LicenceRoot -ChildPath 'slmgr-win7.vbs')",
        [string]
        $Ospp = 'C:\WINDOWS\System32\ospp.vbs',
        [string]
        $OsppHtm = 'C:\WINDOWS\System32\ospp.htm',
        [string]
        $BackupOspp = "$(Join-Path -Path $LicenceRoot -ChildPath 'ospp.vbs')",
        [string]
        $BackupOsppHtm = "$(Join-Path -Path $LicenceRoot -ChildPath 'ospp.htm')"
    ) #param
    begin {
        Write-Verbose -Message 'Creating TempFile'
        $TempFile = (Join-Path -Path $env:TEMP -ChildPath $(([System.Guid]::NewGuid().Guid) + '.txt'))
        New-Item -Path $TempFile -ItemType File | Out-Null
        Write-Verbose -Message (
            'Checking to see if ospp.vbs is available locally. (It might not be if running from a server.)'
        )
        if (-not (Test-Path -Path $Ospp)) {
            Write-Verbose -Message "Can't find ospp.vbs. Copying it from $BackupOspp"
            $SplatArgs = @{ Path = "$BackupOspp"
                            Destination = "$Ospp" }
            Copy-Item @SplatArgs | Out-Null
        } elseif (Test-Path -Path $Ospp) {
            Write-Verbose -Message "Found ospp.vbs, using local copy."
        } #if -not Ospp
        if (-not (Test-Path -Path $OsppHtm)) {
            $SplatArgs = @{ Path = "$BackupOspp"
                            Destination = "$OsppHtm" }
            Copy-Item @SplatArgs | Out-Null
        } elseif (Test-Path -Path $Ospp) {
            Write-Verbose -Message "Found ospp.htm, using local copy."
        } #if -not OsppHtm
        Write-Verbose -Message "Checking ComputerName."
        if (-not $ComputerName) {
            Write-Verbose -Message "No ComputerName supplied, assuming $env:COMPUTERNAME as ComputerName."
            $ComputerName = $env:COMPUTERNAME
        } #if -not ComputerName
        if ($CacheMyPin) {
            Write-Verbose -Message 'Not caching PIN because CacheMyPin logic is commented out.'
            Write-Verbose -Message 'Uncomment lines 140-164,317,320,323,335,338,341,370,373,376,388,390,392 to use the CacheMyPin parameter.'
            <#
            Write-Verbose -Message 'Preparing to cache PIN due to CacheMyPin parameter.'
            Clear-Host
            Write-Host -Object ''
            Write-Host -Object ''
            Write-Host -Object '    ####################################################' -ForegroundColor Magenta
            Write-Host -Object '    #                                                  #' -ForegroundColor Magenta
            Write-Host -Object '    #        You used the CacheMyPin parameter.        #' -ForegroundColor Magenta
            Write-Host -Object '    #      When you are prompted for credentials,      #' -ForegroundColor Magenta
            Write-Host -Object '    #       select a certificate for token-based       #' -ForegroundColor Magenta
            Write-Host -Object '    #    activation. Don't just enter a username and   #' -ForegroundColor Magenta
            Write-Host -Object '    #   password. You must use select a certificate.   #' -ForegroundColor Magenta
            Write-Host -Object '    #     Not fully tested. Use at your own risk!      #' -ForegroundColor Magenta
            Write-Host -Object '    #                                                  #' -ForegroundColor Magenta
            Write-Host -Object '    ####################################################' -ForegroundColor Magenta
            Write-Host -Object ''
            Write-Host -Object ''
            $AnyKey = Read-Host -Prompt 'Press any key to continiue'
            $AnyKey | Out-Null
            Remove-Variable -Name AnyKey
            Clear-Host
            $Creds = Get-Credential -Message 'Select a certificate and enter a PIN for activation.'
            [string] $Pin = $Creds.GetNetworkCredential().Password
            Write-Verbose -Message 'Your PIN has been cached.'
            #>
        } #if CacheMyPin
    } #begin
    process {
        $ComputerName | ForEach-Object {
            $EachComputer = $_
            Write-Verbose -Message "Check to see if we're running locally."
            if ($EachComputer -match $env:COMPUTERNAME) {
                $RunningLocally = $true
            } else {
                $RunningLocally = $false
            } #if local or remote
            Write-Verbose -Message 'Getting Office install state via WMI'
            if ($RunningLocally) {
                $SplatArgs = @{ Class = 'Win32_Product'
                                Property = 'IdentifyingNumber'
                                Filter = 'IdentifyingNumber="{90150000-0011-0000-0000-0000000FF1CE}"' }
                $Office = Get-WmiObject @SplatArgs
            } else {
                $SplatArgs = @{ Class = 'Win32_Product'
                                Property = 'IdentifyingNumber'
                                Filter = 'IdentifyingNumber="{90150000-0011-0000-0000-0000000FF1CE}"'
                                ComputerName = $EachComputer }
                $Office = Get-WmiObject @SplatArgs
            } #if
            Write-Verbose -Message 'Getting SoftwareLicensingTokenActivationLicense via WMI'
            if ($RunningLocally) {
                $WmiLicences = Get-WmiObject -Class SoftwareLicensingTokenActivationLicense
            } else {
                $SplatArgs = @{ Class = 'SoftwareLicensingTokenActivationLicense'
                                ComputerName = $EachComputer }
                $WmiLicences = Get-WmiObject @SplatArgs
            } #if
            Write-Verbose -Message 'Getting Win32_OperatingSystem via WMI'
            if ($RunningLocally) {
                $SplatArgs = @{ Class = 'Win32_OperatingSystem'
                                Property = 'ProductType,Version' }
                $Os = Get-WmiObject @SplatArgs
            } else {
                $SplatArgs = @{ Class = 'Win32_OperatingSystem'
                                Property = 'ProductType,Version'
                                ComputerName = $EachComputer }
                $Os = Get-WmiObject @SplatArgs
            } #if
            $Version = $Os.Version
            $ProductType = $Os.ProductType
            Write-Verbose -Message 'Checking for installed licences'
            if ($WmiLicences -ne $null) {
                Write-Verbose -Message 'Getting Windows licence.'
                $WindowsLicence = $WmiLicences |
                    Where-Object { ($_.Description).ToString() -match '[\s\S]Windows[\s\S]' }
                if ($WindowsLicence -ne $null) {
                    Write-Verbose -Message 'Windows licence found, uninstalling.'
                    $WindowsLicence | Invoke-WmiMethod -Name Uninstall | Out-Null
                } else {
                    Write-Verbose -Message 'No Windows licence installed.'
                } #if WindowsLicence
                Write-Verbose -Message 'Getting Office licence.'
                $OfficeLicence = $WmiLicences |
                    Where-Object { ($_.Description).ToString() -match '[\s\S]Office[\s\S]' }
                if ($OfficeLicence -ne $null) {
                    Write-Verbose -Message 'Office licence found, uninstalling.'
                    $OfficeLicence | Invoke-WmiMethod -Name Uninstall | Out-Null
                } else {
                    Write-Verbose -Message 'No Office licence installed.'
                } #if OfficeLicence
            } #if WmiLicences
            Write-Verbose -Message 'Starting OS detection'
            if (($Version -match '^6\.1') -and ($ProductType -eq '1')) {
                $OsVersion = 'Windows 7'
                $Licence = $Win7
                $SplatArgs = @{ Path = "$BackupSlmgrWin7"
                                Destination = "$SlmgrWin7"
                                Force = $true }
                Copy-Item @SplatArgs | Out-Null
                $Slmgr = $SlmgrWin7
            } elseif (($Version -match '^6\.1') -and ($ProductType -in @('2','3'))) {
                $OsVersion = 'Windows Server 2008 R2'
                $Licence = $Svr2008r2
                $SplatArgs = @{ Path = "$BackupSlmgrWin7"
                                Destination = "$SlmgrWin7"
                                Force = $true }
                Copy-Item @SplatArgs | Out-Null
                $Slmgr = $SlmgrWin7
            } elseif (($Version -match '^6\.2') -and ($ProductType -in @('2','3'))) {
                $OsVersion = 'Windows Server 2012'
                $Licence = $Svr2012
            } elseif (($Version -match '^6\.3') -and ($ProductType -eq '1')) {
                $OsVersion = 'Windows 8.1'
                $Licence = $Win81
            } elseif (($Version -match '^6\.3') -and ($ProductType -in @('2','3'))) {
                $OsVersion = 'Windows Server 2012 R2'
                $Licence = $Svr2012r2
            } elseif (($Version -match '^10\.') -and ($ProductType -eq '1')) {
                $OsVersion = 'Windows 10'
                $Licence = $Win10
            } #if Version and ProductType
            Write-Verbose -Message "$OsVersion detected, selecting $Licence for install."
            # Install Windows Licence
            $LicenceFile = (Join-Path -Path $LicenceRoot -ChildPath $Licence)
            if ($RunningLocally) {
                [string] $SlmgrIlcArgs = "$Slmgr //Nologo /ilc `"$LicenceFile`""
            } else {
                [string] $SlmgrIlcArgs = "$Slmgr //Nologo $EachComputer /ilc `"$LicenceFile`""
            } #if
            $SplatArgs = @{ FilePath = $CScript
                            ArgumentList = $SlmgrIlcArgs
                            Wait = $true
                            NoNewWindow = $true
                            RedirectStandardOutput = $TempFile
                            PassThru = $true }
            $SlmgrIlcProc = Start-Process @SplatArgs
            $SlmgrIlcOutput = Get-Content -Path $TempFile -Raw
            Remove-Item -Path $TempFile | Out-Null
            if ($SlmgrIlcProc.ExitCode -eq 0) {
                Write-Verbose -Message $SlmgrIlcOutput
            } elseif ($SlmgrIlcProc.ExitCode -ne 0) {
                Write-Error -Message $SlmgrIlcOutput
            } #if SlmgrIlcProc.ExitCode
            if ($Office -ne $null) {
                Write-Verbose -Message 'Office is installed. Installing Office licence.'
                $OfficeLicenceFile = (Join-Path -Path $LicenceRoot -ChildPath $Office2013ProPlus)
                if ($RunningLocally) {
                    [string] $OsppInslicArgs = "`"$Ospp`" //Nologo /inslic:`"$OfficeLicenceFile`""
                } else {
                    [string] $OsppInslicArgs = "`"$Ospp`" //Nologo /inslic:`"$OfficeLicenceFile`" $EachComputer"
                } #if
                $SplatArgs = @{ FilePath = $CScript
                                ArgumentList = $OsppInslicArgs
                                Wait = $true
                                NoNewWindow = $true
                                RedirectStandardOutput = $TempFile
                                PassThru = $true }
                $OsppInslicProc = Start-Process @SplatArgs
                $OsppInslicOutput = Get-Content -Path $TempFile -Raw
                Remove-Item -Path $TempFile | Out-Null
                if ($OsppInslicProc.ExitCode -eq 0) {
                    Write-Verbose -Message $OsppInslicOutput
                } elseif ($OsppInslicProc.ExitCode -ne 0) {
                    Write-Error -Message $OsppInslicOutput
                } #if $OsppInslicProc.ExitCode
            } #if Office
            Write-Verbose -Message 'Activate Windows'
            if ($RunningLocally) {
                [string] $SlmgrLtcArgs = "$Slmgr //Nologo /ltc"
                $SplatArgs = @{ FilePath = $CScript
                                ArgumentList = $SlmgrLtcArgs
                                Wait = $true
                                NoNewWindow = $true
                                RedirectStandardOutput = $TempFile }
                Start-Process @SplatArgs
                $SlmgrLtcOutput = Get-Content -Path $TempFile
                Remove-Item -Path $TempFile | Out-Null
                $CertThumbPrint = (($SlmgrLtcOutput | Select-String -Pattern 'Thumbprint')[0] -split ':')[1].Trim()
                #if (-not $CacheMyPin) {
                    Write-Verbose -Message 'Prompting for PIN to Activate Windows'
                    [string] $SlmgrFtaArgs = "$Slmgr //Nologo /fta $CertThumbPrint"
                #} elseif ($CacheMyPin) {
                    #Write-Verbose -Message 'Using cached PIN to Activate Windows'
                    #[string] $SlmgrFtaArgs = "$Slmgr //Nologo /fta $CertThumbPrint $Pin"
                #} #if ($CacheMyPin)
            } else {
                [string] $SlmgrLtcArgs = "$Slmgr //Nologo $EachComputer /ltc"
                $SplatArgs = @{ FilePath = $CScript
                                ArgumentList = $SlmgrLtcArgs
                                Wait = $true
                                NoNewWindow = $true
                                RedirectStandardOutput = $TempFile }
                Start-Process @SplatArgs
                $SlmgrLtcOutput = Get-Content -Path $TempFile
                Remove-Item -Path $TempFile | Out-Null
                $CertThumbPrint = (($SlmgrLtcOutput | Select-String -Pattern 'Thumbprint')[0] -split ':')[1].Trim()
                #if (-not $CacheMyPin) {
                    Write-Verbose -Message 'Prompting for PIN to Activate Windows'
                    [string] $SlmgrFtaArgs = "$Slmgr //Nologo $EachComputer /fta $CertThumbPrint"
                #} elseif ($CacheMyPin) {
                    #Write-Verbose -Message 'Using cached PIN to Activate Windows'
                    #[string] $SlmgrFtaArgs = "$Slmgr //Nologo $EachComputer /fta $CertThumbPrint $Pin"
                #} #if ($CacheMyPin)
            } #if
            $SplatArgs = @{ FilePath = $CScript
                            ArgumentList = $SlmgrFtaArgs
                            Wait = $true
                            NoNewWindow = $true
                            RedirectStandardOutput = $TempFile
                            PassThru = $true }
            $SlmgrFtaProc = Start-Process @SplatArgs
            $SlmgrFtaOutput = Get-Content -Path $TempFile -Raw
            Remove-Item -Path $TempFile | Out-Null
            if ($SlmgrFtaProc.ExitCode -eq 0) {
                Write-Verbose -Message $SlmgrFtaOutput
            } elseif ($SlmgrFtaProc.ExitCode -ne 0) {
                Write-Error -Message $SlmgrFtaOutput
            } #if SlmgrFtaProc.ExitCode
            if ($Office -ne $null) {
                Write-Verbose -Message 'Searching for valid certs to activate Office...'
                if ($RunningLocally) {
                    [string] $OsppDtokcertsArgs = "`"$Ospp`" //Nologo /dtokcerts"
                    $SplatArgs = @{ FilePath = $CScript
                                    ArgumentList = $OsppDtokcertsArgs
                                    Wait = $true
                                    NoNewWindow = $true
                                    RedirectStandardOutput = $TempFile }
                    Start-Process @SplatArgs
                    $OsppDtokcertsOutput = Get-Content -Path $TempFile
                    Remove-Item -Path $TempFile | Out-Null
                    $CertThumbPrint = (($OsppDtokcertsOutput | Select-String -Pattern 'Thumbprint')[0] -split ':')[1].Trim()
                    #if (-not $CacheMyPin) {
                        Write-Verbose -Message 'Prompting for PIN to Activate Office'
                        [string] $OsppTokactArgs = "`"$Ospp`" //Nologo /tokact:$CertThumbPrint"
                    #} elseif ($CacheMyPin) {
                        #Write-Verbose -Message 'Using cached PIN to Activate Office'
                        #[string] $OsppTokactArgs = "`"$Ospp`" //Nologo /tokact:$CertThumbPrint`:$Pin"
                    #} #if ($CacheMyPin)
                } else {
                    [string] $OsppDtokcertsArgs = "`"$Ospp`" //Nologo /dtokcerts $EachComputer"
                    $SplatArgs = @{ FilePath = $CScript
                                    ArgumentList = $OsppDtokcertsArgs
                                    Wait = $true
                                    NoNewWindow = $true
                                    RedirectStandardOutput = $TempFile }
                    Start-Process @SplatArgs
                    $OsppDtokcertsOutput = Get-Content -Path $TempFile
                    Remove-Item -Path $TempFile | Out-Null
                    $CertThumbPrint = (($OsppDtokcertsOutput | Select-String -Pattern 'Thumbprint')[0] -split ':')[1].Trim()
                    #if (-not $CacheMyPin) {
                        [string] $OsppTokactArgs = "`"$Ospp`" //Nologo /tokact:$CertThumbPrint $EachComputer"
                    #} elseif ($CacheMyPin) {
                        #[string] $OsppTokactArgs = "`"$Ospp`" //Nologo /tokact:$CertThumbPrint`:$Pin $EachComputer"
                    #} #if ($CacheMyPin)
                } #if
                $SplatArgs = @{ FilePath = $CScript
                                ArgumentList = $OsppTokactArgs
                                Wait = $true
                                NoNewWindow = $true
                                RedirectStandardOutput = $TempFile
                                PassThru = $true }
                $OsppTokactProc = Start-Process @SplatArgs
                $OsppTokactOutput = Get-Content -Path $TempFile -Raw
                Remove-Item -Path $TempFile | Out-Null
                if ($OsppTokactProc.ExitCode -eq 0) {
                    Write-Verbose -Message $OsppTokactOutput
                } elseif ($OsppTokactProc.ExitCode -ne 0) {
                    Write-Error -Message $OsppTokactOutput
                } #if OsppTokactProc.ExitCode
            } #if Office
        } #ForEach ComputerName
    } #process
    end {} #end
} #function Update-TokenBasedActivationLicence