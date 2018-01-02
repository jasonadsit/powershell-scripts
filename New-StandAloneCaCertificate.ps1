# New-StandAloneCaCertificate
#
# https://gist.github.com/jasonadsit/be97a6a3331dbee61cd74a1ebd4adf11
# https://pastebin.com/HxrjJ10z
#
[CmdletBinding()]
param (
    [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the FQDN of the server hosting the standalone CA.")]
    [string]
    $CaServerName,
    [parameter(Mandatory=$true,Position=1,HelpMessage="Enter the name of the standalone CA.")]
    [string]
    $CaName,
    [string]
    $CaString = "$CaServerName\$CaName",
    [string]
    $WorkingDirectory = $env:TEMP,
    [string]
    $StandAloneCaCertRequestInf = $(Join-Path -Path $WorkingDirectory -ChildPath 'StandAloneCaCertRequest.inf'),
    [string]
    $StandAloneCaCertRequestReq = $(Join-Path -Path $WorkingDirectory -ChildPath 'StandAloneCaCertRequest.req'),
    [string]
    $StandAloneCaCertResponseCert = $(Join-Path -Path $WorkingDirectory -ChildPath 'StandAloneCaCertResponse.cer'),
    [string]
    $StandAloneCaCertResponse = $(Join-Path -Path $WorkingDirectory -ChildPath 'StandAloneCaCertResponse.rsp'),
    [string[]]
    $TempFiles = @($StandAloneCaCertRequestInf,$StandAloneCaCertRequestReq,$StandAloneCaCertResponseCert,$StandAloneCaCertResponse),
    [string]
    $CertReqExe = $(Join-Path -Path ([System.Environment]::SystemDirectory) -ChildPath 'certreq.exe'),
    [string]
    $ComputerName = $env:COMPUTERNAME,
    [string]
    $Fqdn = "$ComputerName.$env:USERDNSDOMAIN",
    [string]
    $ComputerNameUpper = $($ComputerName.ToUpper()),
    [string]
    $FqdnUpper = $($Fqdn.ToUpper()),
    [string]
    $ComputerNameLower = $($ComputerName.ToLower()),
    [string]
    $FqdnLower = $($Fqdn.ToLower())
) #param
process {
    Write-Verbose -Message "Getting the list of machine certificates."
    $StandAloneCaCert = Get-ChildItem -Path Cert:\LocalMachine\My |
        Where-Object {
            $_.FriendlyName -Match $ComputerName -and
            $_.Issuer -EQ "CN=$CaName, DC=example, DC=com"
        }
    Write-Verbose -Message "Testing to see if any are issued by our CA."
    if (-not ($StandAloneCaCert)) {
        Write-Verbose -Message "No certificate issued by our CA found."
        Write-Verbose -Message "Testing if CertReq TempFiles already exist."
        $TempFilesToDelete = $TempFiles | Where-Object { Test-Path -Path $_ }
        if ($TempFilesToDelete) {
            Write-Verbose -Message "Deleting CertReq TempFiles."
            $TempFilesToDelete | Remove-Item | Out-Null
        } #if
        Write-Verbose -Message "Building $StandAloneCaCertRequestInf"
        [string] $RequestInf = @"
;----------------- request.inf -----------------

[Version]
Signature="`$Windows NT$"

[NewRequest]
FriendlyName = "$ComputerNameLower"
Subject = "CN=$FqdnLower, OU=Some Department, C=US, O=Some Company"
KeySpec = 1
KeyAlgorithm = ECDSA_P256
KeyLength = 256
HashAlgorithm = SHA256
Exportable = TRUE
MachineKeySet = TRUE
SMIME = False
PrivateKeyArchive = FALSE
EncryptionAlgorithm = AES
EncryptionLength = 256
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft Software Key Storage Provider"
RequestType = PKCS10
KeyUsage = 0xf0
SuppressDefaults = TRUE

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.1 ; Server Authentication
OID=1.3.6.1.5.5.7.3.2 ; Client Authentication
OID=1.3.6.1.4.1.311.54.1.2 ; Remote Desktop Authentication

[Extensions]
2.5.29.17 = "{text}"
_continue_ = "dns=$FqdnUpper&"
_continue_ = "dns=$ComputerNameUpper&"
_continue_ = "dns=$FqdnLower&"
_continue_ = "dns=$ComputerNameLower"
;-----------------------------------------------
"@
        Write-Verbose -Message "Backing-Up FIPS Policy"
        $CurrentFipsPolicy = (Get-Item -Path HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy).GetValue('Enabled')
        Write-Verbose -Message "Killing FIPS Policy"
        Set-ItemProperty HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy -Name Enabled -Value 0 -Force | Out-Null
        Write-Verbose -Message "Writing $StandAloneCaCertRequestInf"
        if (Test-Path -Path $StandAloneCaCertRequestInf) {
            Remove-Item -Path $StandAloneCaCertRequestInf -Force | Out-Null
            New-Item -Path $StandAloneCaCertRequestInf -ItemType File | Out-Null
        } elseif (-not (Test-Path -Path $StandAloneCaCertRequestInf)) {
            New-Item -Path $StandAloneCaCertRequestInf -ItemType File | Out-Null
        } #if
        Add-Content -Path $StandAloneCaCertRequestInf -Value $RequestInf | Out-Null
        Write-Verbose -Message "Generating certificate request."
        $CertReqArgs = "-New -machine `"$StandAloneCaCertRequestInf`" `"$StandAloneCaCertRequestReq`""
        $SplatArgs = @{ FilePath = $CertReqExe
                        ArgumentList = $CertReqArgs
                        NoNewWindow = $true
                        Wait = $true }
        $CertReqProc = Start-Process @SplatArgs
        Write-Verbose -Message "Submitting certificate request."
        $CertReqArgs = "-Submit -rpc -AdminForceMachine -config `"$CaString`" `"$StandAloneCaCertRequestReq`" `"$StandAloneCaCertResponseCert`""
        $SplatArgs = @{ FilePath = $CertReqExe
                        ArgumentList = $CertReqArgs
                        NoNewWindow = $true
                        Wait = $true }
        $CertReqProc = Start-Process @SplatArgs
        Write-Verbose -Message "Accepting certificate request."
        $CertReqArgs = "-Accept -machine `"$StandAloneCaCertResponseCert`""
        $SplatArgs = @{ FilePath = $CertReqExe
                        ArgumentList = $CertReqArgs
                        NoNewWindow = $true
                        Wait = $true }
        $CertReqProc = Start-Process @SplatArgs
        Write-Verbose -Message "Deleting CertReq TempFiles."
        $TempFilesToDelete = $TempFiles | Where-Object { Test-Path -Path $_ }
        if ($TempFilesToDelete) {
            $TempFilesToDelete | Remove-Item | Out-Null
        } #if
        $StandAloneCaCert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object {
            $_.FriendlyName -Match $ComputerName -and
            $_.Issuer -EQ "CN=$CaName, DC=example, DC=com"
        }
        Write-Verbose -Message "Reset FIPS policy to its original value."
        Set-ItemProperty HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy -Name Enabled -Value $CurrentFipsPolicy -Force | Out-Null
    } elseif ($StandAloneCaCert) {
        Write-Verbose -Message "There is already a certificate issued by our CA installed. Skipping."
    } #if
    Write-Output -InputObject $StandAloneCaCert
} #process
# New-StandAloneCaCertificate