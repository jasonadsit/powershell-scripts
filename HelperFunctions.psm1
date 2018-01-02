function Disable-FipsMode {
    [CmdletBinding()]
    param () #param
    process {
        $NewItemPropertyArgs = @{   Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy'
                                    Name = 'Enabled'
                                    PropertyType = 'DWord'
                                    Value = '0'
                                    Force = $true }
        New-ItemProperty @NewItemPropertyArgs | Out-Null
    } #process
} #function Disable-FipsMode

function Repair-PowerShellGet {
    [CmdletBinding()]
    param (
        [string]
        $FipsRegPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy'
    ) #param
    process {
        # Kill FIPS so bootstrapping NuGet doesn't fail.
        $SetItemPropertyArgs = @{   Path = $FipsRegPath
                                    Name = 'Enabled'
                                    Value = '0'
                                    Force = $true }
        Set-ItemProperty @SetItemPropertyArgs | Out-Null
        # Build the ScriptBlock
        $ScriptBlock = [System.Management.Automation.ScriptBlock]::Create({
        Install-PackageProvider -Name NuGet -ForceBootstrap -Force | Out-Null
        Set-PackageSource -Name PSGallery -Trusted -ForceBootstrap -Force | Out-Null
        Install-Module -Name PowerShellGet -Scope AllUsers -Force | Out-Null
        })
        # Run it
        $StartProcArgs = @{ FilePath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
                            ArgumentList = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $ScriptBlock"
                            Wait = $true
                            NoNewWindow = $true }
        Start-Process @StartProcArgs | Out-Null
    } #process
} #function Repair-PowerShellGet

function Invoke-WuaucltDetectNow {
    [CmdletBinding()]
    param (
    [string]
    $WuaucltPath = 'C:\Windows\System32\wuauclt.exe',
    [string]
    $WuaucltArgs = '/detectnow'
    ) #param
    process {
        $SplatArgs = @{ FilePath = $WuaucltPath
                        ArgumentList = $WuaucltArgs
                        NoNewWindow = $true
                        Wait = $true }
        Start-Process @SplatArgs
    } #process
} #function Invoke-WuaucltDetectNow

function Invoke-WuaucltReportNow {
    [CmdletBinding()]
    param (
    [string]
    $WuaucltPath = 'C:\Windows\System32\wuauclt.exe',
    [string]
    $WuaucltArgs = '/reportnow'
    ) #param
    process {
        $SplatArgs = @{ FilePath = $WuaucltPath
                        ArgumentList = $WuaucltArgs
                        NoNewWindow = $true
                        Wait = $true }
        Start-Process @SplatArgs
    } #process
} #function Invoke-WuaucltReportNow

function Invoke-GracefulRestart {
    [CmdletBinding()]
    param (
        [string[]]
        $ComputerName,
        [int]
        $Minutes,
        [int]
        $Seconds = $($Minutes * 60),
        [string]
        $ShutdownArgs = "/r /f /t $Seconds /c 'Your computer will restart in $Minutes minute(s) for required maintenance. Please save your work.' /d P:0:0",
        [string]
        $ShutDownPath = 'C:\Windows\System32\shutdown.exe'
    ) #param
    process {
        if (-not $ComputerName) {
            $ComputerName = $env:COMPUTERNAME
        } #if
        $ComputerName | ForEach-Object {
            if ($_ -match $env:COMPUTERNAME) {
                $SplatArgs = @{ FilePath = "$ShutDownPath"
                                ArgumentList = "$ShutdownArgs"
                                NoNewWindow = $true
                                Wait = $true }
                Start-Process @SplatArgs
            } else {
                Invoke-Command -ComputerName $_ -ScriptBlock {
                    param (
                        $ShutDownPath,
                        $ShutdownArgs
                    ) #param
                    $SplatArgs = @{ FilePath = "$ShutDownPath"
                                    ArgumentList = "$ShutdownArgs"
                                    NoNewWindow = $true
                                    Wait = $true }
                    Start-Process @SplatArgs
                } -ArgumentList $ShutDownPath,$ShutdownArgs #Invoke-Command
            } #if
        } #ForEach
    } #process
} #function Invoke-GracefulRestart

function Invoke-SilentRestart {
    [CmdletBinding()]
    param (
        [string]
        $ShutdownArgs = '/r /f /t 1 /d P:0:0',
        [string]
        $ShutDownPath = 'C:\Windows\System32\shutdown.exe'
    ) #param
    process {
        $SplatArgs = @{ FilePath = "$ShutDownPath"
                        ArgumentList = "$ShutdownArgs"
                        NoNewWindow = $true
                        Wait = $true }
        Start-Process @SplatArgs
    } #process
} #function Invoke-SilentRestart

function Get-LoggedOnUser {
    [CmdletBinding()]
    param (
        [string]
        $QueryPath = 'C:\Windows\System32\query.exe',
        [string]
        $QueryArgs = 'session console',
        [string]
        $TempFile = (Join-Path -Path $env:TEMP -ChildPath $(([System.Guid]::NewGuid().Guid) + '.txt'))
    ) #param
    process {
        $SplatArgs = @{ FilePath = $QueryPath
                        ArgumentList = $QueryArgs
                        NoNewWindow = $true
                        Wait = $true
                        RedirectStandardOutput = $TempFile }
        Start-Process @SplatArgs
        $QueryOutput = Get-Content -Path $TempFile
        Remove-Item -Path $TempFile | Out-Null
        $QueryOutput -replace '(^\s+|\s+$)','' -replace '>','' -replace '\s+',',' |
        ConvertFrom-Csv |
        Select-Object -Property SESSIONNAME,USERNAME
    } #process
} #function Get-LoggedOnUser

function ConvertFrom-Sid {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('objectSid')]
        [string[]]
        $SID
    ) #param
    process {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $SID | ForEach-Object {
            $EachSid = $_
            if ($EachSid -eq $null) {
                $objSID = ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).Sid
            } else {
                $objSID = New-Object System.Security.Principal.SecurityIdentifier("$EachSid")
            }
            $objNtAccount = $objSID.Translate([System.Security.Principal.NTAccount])
            $UserName = Split-Path -Path $($objNtAccount.Value) -Leaf
            Write-Output -InputObject $UserName
        } #ForEach
    } #process
} #function ConvertFrom-Sid

function ConvertTo-Sid {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('sAMAccountName')]
        [string[]]
        $UserName
    ) #param
    process {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $UserName | ForEach-Object {
            $EachUserName = $_
            if ($EachUserName -eq $null) {
                $strSam = ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).SamAccountName
                $objAct = New-Object System.Security.Principal.NTAccount("$strSam")
            } else {
                $objAct = New-Object System.Security.Principal.NTAccount("$EachUserName")
            }
            $objSID = $objAct.Translate([System.Security.Principal.SecurityIdentifier])
            Write-Output -InputObject $($objSID.Value)
        } #ForEach
    } #process
} #function ConvertTo-Sid

function Remove-UserProfile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [int]
        $Days
    ) #param
    process {
        Get-CimInstance -ClassName Win32_UserProfile |
        Where-Object { $_.LastUseTime -lt $(Get-Date).Date.AddDays(-$Days) } |
        Remove-CimInstance
    } #process
} #function Remove-UserProfile

function Clear-AllRecycleBin {
    [CmdletBinding()]
    param () #param
    process {
        Get-WmiObject -Class Win32_Volume |
        Select-Object -ExpandProperty DriveLetter |
        ForEach-Object {
            Clear-RecycleBin -DriveLetter $_ -Force
            Remove-Item "$_\`$RECYCLE.BIN\*" -Force -Recurse
            Remove-Item "$_\`$RECYCLE.BIN" -Force -Recurse
        } #ForEach
    } #process
} #function Clear-AllRecycleBin

# Or, if that doesn't work.
#
# Start-Process -FilePath C:\Windows\System32\cmd.exe -ArgumentList '/C "RMDIR /S /Q C:\$Recycle.Bin"' -NoNewWindow -Wait
function New-DesktopShortcut {
    [CmdletBinding()]
    param (
        [parameter(mandatory=$true)]
        [string]
        $ShortcutTarget,
        [parameter(mandatory=$true)]
        [string]
        $ShortcutName,
        [parameter(mandatory=$true)]
        [string]
        $IconLocation
    ) #param
    process {
        $ShortcutFile = "$env:PUBLIC\Desktop\$ShortcutName.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell 
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.IconLocation = $IconLocation
        $Shortcut.TargetPath = $ShortcutTarget
        $Shortcut.Save()
    } #process
} #function New-DesktopShortcut