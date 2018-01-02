# Kill FIPS so bootstrapping NuGet doesn't fail.
$SetItemPropertyArgs = @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy'
    Name = 'Enabled'
    Value = '0'
    Force = $true
}

Set-ItemProperty @SetItemPropertyArgs | Out-Null

# Build the ScriptBlock
$ScriptBlock = [scriptblock]::Create({
Install-PackageProvider -Name NuGet -ForceBootstrap -Force | Out-Null
Set-PackageSource -Name PSGallery -Trusted -ForceBootstrap -Force | Out-Null
Install-Module -Name PowerShellGet,PSReadline,PSWindowsUpdate -Scope AllUsers -SkipPublisherCheck -Force | Out-Null
})

# Run it
$StartProcArgs = @{
    FilePath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    ArgumentList = "-NoProfile -ExecutionPolicy Bypass -Command $ScriptBlock"
    Wait = $true
    NoNewWindow = $true
}

$ProgPrefBak = $ProgressPreference
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

Start-Process @StartProcArgs | Out-Null

$ProgressPreference = $ProgPrefBak