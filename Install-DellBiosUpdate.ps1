#script Install-DellBiosUpdate
#
# Pre-load the BIOS password in a variable if running against a remote system.
#   $SecureBiosPw = Read-Host -Prompt 'Enter the BIOS Password' -AsSecureString
# Then pass it to 'Invoke-Command' with the 'ArgumentList' parameter.
#
[CmdletBinding()]
param (
    [System.Security.SecureString] $SecureBiosPw,
    [string] $BiosPw,
    [string] $CatalogArchiveName = 'CatalogPC.cab',
    [string] $CatalogFileName = 'CatalogPC.xml',
    [string] $DellDownloadRoot = 'https://downloads.dell.com',
    [string] $Flash64ZipName = 'Flash64W.zip',
    [string] $Flash64Uri = "$DellDownloadRoot/FOLDER04165397M/1/$Flash64ZipName",
    [string] $CatalogArchiveUri = "$DellDownloadRoot/catalog/$CatalogArchiveName",
    [string] $LocalBasePath = "$env:ProgramFiles\$(([System.Guid]::NewGuid().Guid))",
    [string] $TempFile = "$LocalBasePath\$(([System.Guid]::NewGuid().Guid)).txt",
    [string] $Flash64Zip = "$LocalBasePath\$Flash64ZipName",
    [string] $LocalCatalogArchivePath = "$LocalBasePath\$CatalogArchiveName",
    [string] $LocalCatalogFilePath = "$LocalBasePath\$CatalogFileName"
) #param
process {
    if (($PSBoundParameters.ContainsKey('SecureBiosPw'))) {
        $BiosPw = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureBiosPw))
    } elseif (-not ($PSBoundParameters.ContainsKey('SecureBiosPw'))) {
        if (($PSBoundParameters.ContainsKey('BiosPw'))) {
            Write-Verbose -Message 'Using the passed-in plaintext BIOS password.'
        } elseif (-not ($PSBoundParameters.ContainsKey('BiosPw'))) {
            $SecureBiosPw = Read-Host -Prompt 'Enter the BIOS Password' -AsSecureString
            $BiosPw = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureBiosPw))
        } #if ($PSBoundParameters.ContainsKey('BiosPw'))
    } #if ($PSBoundParameters.ContainsKey('SecureBiosPw'))
    if (-not (Test-Path -Path $LocalBasePath)) {
        New-Item -Path $LocalBasePath -ItemType Directory | Out-Null
    } #if (-not (Test-Path -Path $LocalBasePath))
    $TempDir = Get-Item -Path $LocalBasePath
    $WebClient = New-Object -TypeName System.Net.WebClient
    $SplatArgs = @{ Namespace = 'root/cimv2'
                    ClassName = 'Win32_ComputerSystem'
                    Property = 'Manufacturer','Model','SystemFamily','SystemSKUNumber' }
    $ComputerSystem = Get-CimInstance @SplatArgs
    $SplatArgs = @{ Namespace = 'root/cimv2'
                    ClassName = 'Win32_BIOS'
                    Property = 'SMBIOSBIOSVersion' }
    $CurrentBios = (Get-CimInstance @SplatArgs).SMBIOSBIOSVersion
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Verbose -Message "Downloading the Dell catalog archive."
    $WebClient.DownloadFile("$CatalogArchiveUri","$LocalCatalogArchivePath")
    while (-not (Test-Path -Path $LocalCatalogArchivePath)) {
        Start-Sleep -Seconds 1
    } #while
    $ExpandArgs = "`"$LocalCatalogArchivePath`" -F:$CatalogFileName `"$LocalCatalogFilePath`""
    $SplatArgs = @{ FilePath = 'C:\Windows\System32\expand.exe'
                    ArgumentList = $ExpandArgs
                    NoNewWindow = $true
                    Wait = $true
                    RedirectStandardOutput = $TempFile }
    Write-Verbose -Message "Expanding the archive."
    Start-Process @SplatArgs | Out-Null
    Remove-Item -Path $TempFile | Out-Null
    while (-not (Test-Path -Path $LocalCatalogFilePath)) {
        Start-Sleep -Seconds 1
    } #while
    $DriverCatalogXml = (Get-Content -Path $LocalCatalogFilePath -Raw) -as [xml]
    $DownloadBase = $DriverCatalogXml.Manifest.baseLocation
    Write-Verbose -Message 'Selecting BIOS update.'
    $InstallerFile = $DriverCatalogXml.Manifest.SoftwareComponent |
        Where-Object {
            $_.SupportedSystems.Brand.Model.systemID -eq $ComputerSystem.SystemSKUNumber -and
            $_.ComponentType.value -eq 'BIOS'
        } | Sort-Object -Property dateTime -Descending | Select-Object -First 1
    if (-not ($InstallerFile.dellVersion -eq $CurrentBios)) {
        Write-Verbose -Message "'$($InstallerFile.dellVersion)' is newer that the installed BIOS version, preparing for update."
        $InstallerUri = "https://$DownloadBase/$($InstallerFile.path)"
        $InstallerFileName = $InstallerUri.Split('/') | Select-Object -Last 1
        $InstallerLocalPath = Join-Path -Path $LocalBasePath -ChildPath $InstallerFileName
        Write-Verbose -Message "Downloading $InstallerFileName"
        $WebClient.DownloadFile("$InstallerUri","$InstallerLocalPath")
        while (-not (Test-Path -Path $InstallerLocalPath)) {
            Start-Sleep -Seconds 1
        } #while
        $WebClient.DownloadFile("$Flash64Uri","$Flash64Zip")
        while (-not (Test-Path -Path $Flash64Zip)) {
            Start-Sleep -Seconds 1
        } #while
        $Flash64Dir = Join-Path -Path $LocalBasePath -ChildPath 'Flash64Dir'
        $ProgPrefBak = $ProgressPreference
        $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        Get-Item -Path $Flash64Zip | Expand-Archive -DestinationPath $Flash64Dir
        $ProgressPreference = $ProgPrefBak
        $Flash64 = Get-ChildItem -Path $Flash64Dir -Filter *.exe
        $Flash64 | Unblock-File
        Write-Verbose -Message 'Suspending BitLocker'
        Suspend-BitLocker -MountPoint 'C:' -RebootCount 1 -Confirm:$false | Out-Null
        $BiosArgs = "/s /f /b=`"$InstallerLocalPath`" /p=`"$BiosPw`" /l=`"$TempFile`""
        $SplatArgs = @{ FilePath = "$($Flash64.FullName)"
                        ArgumentList = $BiosArgs
                        NoNewWindow = $true
                        Wait = $true }
        Write-Verbose -Message "Running $InstallerFileName"
        Start-Process @SplatArgs
        $BiosOutput = Get-Content -Path $TempFile -Raw
        Remove-Item -Path $TempFile | Out-Null
        if ($BiosOutput -match 'Password\ Validation\ failure') {
            $BiosStatus = $false
            $Reason = 'BIOS Flash Failed: Bad Password'
        } elseif (($BiosOutput -match 'BIOS\ flash\ started') -and ($BiosOutput -match 'BIOS\ flash\ finished')) {
            $BiosStatus = $true
            $Reason = 'BIOS Flash Succeeded: Rebooting'
        } #if
        if ($BiosStatus) {
            $EventObject = New-Object -TypeName psobject -Property @{
                TimeStamp = [datetime] (Get-Date)
                ComputerName = $env:COMPUTERNAME
                Bios = [string] $InstallerFileName
                Message = $Reason
            } #New-Object
            $EventObject | Select-Object -Property TimeStamp,ComputerName,Bios,Message
            Restart-Computer -Force
        } elseif (-not $BiosStatus) {
            $EventObject = New-Object -TypeName psobject -Property @{
                TimeStamp = [datetime] (Get-Date)
                ComputerName = $env:COMPUTERNAME
                Bios = [string] $InstallerFileName
                Message = $Reason
            } #New-Object
            $EventObject | Select-Object -Property TimeStamp,ComputerName,Bios,Message
        } #if
    } elseif ($InstallerFile.dellVersion -eq $CurrentBios) {
        Write-Verbose -Message "The current BIOS version '$CurrentBios', matches the latest version from the Dell catalog, skipping."
        $EventObject = New-Object -TypeName psobject -Property @{
            TimeStamp = [datetime] (Get-Date)
            ComputerName = $env:COMPUTERNAME
            Bios = $null
            Message = "$CurrentBios is current, skipping."
        } #New-Object
        $EventObject | Select-Object -Property TimeStamp,ComputerName,Bios,Message
    } #if
    Write-Verbose -Message 'Cleaning up'
    if (Test-Path -Path $($TempDir.FullName)) {
        $TempDir | Remove-Item -Recurse -Force
    } #if
} #process
#script Install-DellBiosUpdate