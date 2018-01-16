# Set-AdDescription

[CmdletBinding()]
[OutputType([psobject])]

param (

    [string]
    $ProfileFilter = "Special=False AND SID LIKE 'S-1-5-21-%' AND NOT SID LIKE '%-500'",

    [string]
    $SitesRaw = @'
"SiteName","SiteCode","Range"
"Oregon","OR","0 1 2 3 4"
"Washington","WA","5 6 7 8 9"
"Idaho","ID","10 11 12 13 14"
"New York","NY","15 16 17 18 19"
"Texas","TX","20 21 22 23 24"
'@

) #param

begin {

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    $Ipv4Validation = '^(?:(?:1\d\d|2[0-5][0-5]|2[0-4]\d|0?[1-9]\d|0?0?\d)\.){3}(?:1\d\d|2[0-5][0-5]|2[0-4]\d|0?[1-9]\d|0?0?\d)$'

    function ConvertFrom-Sid {
        [CmdletBinding()]
        [OutputType([string])]
        param (
            [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
            [ValidateNotNullOrEmpty()]
            [Alias('objectSid')]
            [string[]]
            $SID
        ) #param
        begin {
            Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        } #begin
        process {
            if ($SID) {
                $SID | ForEach-Object {
                    $private:objSID = New-Object System.Security.Principal.SecurityIdentifier($_)
                    $private:objNtAccount = $private:objSID.Translate([System.Security.Principal.NTAccount])
                    $private:UserName = Split-Path -Path $($private:objNtAccount.Value) -Leaf
                    Write-Output -InputObject $private:UserName
                    Clear-Variable -Name objSID -Scope private
                    Clear-Variable -Name objNtAccount -Scope private
                    Clear-Variable -Name UserName -Scope private
                } #ForEach
            } #if $SID
        } #process
    } #function ConvertFrom-Sid

} #begin

process {

    $ErrorActionPreferenceBak = $ErrorActionPreference
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

    try {

        $OnDomain = [System.Net.Dns]::GetHostByName($env:USERDNSDOMAIN)

    } catch {

        $OnDomain = $false

    } #trycatch

    $ErrorActionPreference = $ErrorActionPreferenceBak

    if ($OnDomain) {

        $SplatArgs = @{ ClassName = 'Win32_IP4RouteTable'
                        Property = ('Destination','InterfaceIndex')
                        Filter = "Destination='0.0.0.0'" }

        $IfIndex = (Get-CimInstance @SplatArgs | Select-Object -First 1).InterfaceIndex

        $SplatArgs = @{ ClassName = 'Win32_NetworkAdapterConfiguration'
                        Property = ('InterfaceIndex','IPAddress')
                        Filter = "InterfaceIndex=$IfIndex" }

        $IpAddressList = (Get-CimInstance @SplatArgs).IPAddress

        $IpAddress = $IpAddressList | Where-Object { $_ -match $Ipv4Validation }

        $SplatArgs = @{ ClassName = 'Win32_ComputerSystemProduct'
                        Property = ('IdentifyingNumber','Name','Vendor') }

        $ComputerSystemProduct = Get-CimInstance @SplatArgs

        $SerialNumber = $ComputerSystemProduct.IdentifyingNumber.Trim()
        $Vendor = ($ComputerSystemProduct.Vendor.Split(' ') | Select-Object -First 1).Trim()
        $Model = $ComputerSystemProduct.Name.Trim()

        $SplatArgs = @{ ClassName = 'Win32_UserProfile'
                        Property = ('LastUseTime','LocalPath','SID','Special')
                        Filter = $ProfileFilter }

        $UserProfiles = Get-CimInstance @SplatArgs

        $LastLoggedOnUser = $UserProfiles | Sort-Object -Property LastUseTime -Descending | Select-Object -First 1

        $ErrorActionPreferenceBak = $ErrorActionPreference
        $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

        try {

            $UserName = $LastLoggedOnUser.SID | ConvertFrom-Sid

        } catch {

            $UserName = Split-Path -Path ($LastLoggedOnUser.LocalPath) -Leaf

        } #trycatch

        $ErrorActionPreference = $ErrorActionPreferenceBak

        $Sites = $SitesRaw | ConvertFrom-Csv | ForEach-Object {
            New-Object -TypeName psobject -Property @{  SiteName = $_.SiteName
                                                        SiteCode = $_.SiteCode
                                                        SubNets = $_.Range.Split(' ') }
        } # $Sites

        $ThirdOctet = $IpAddress.Split('.')[2]
        $Site = $Sites | Where-Object { $ThirdOctet -in $_.SubNets }
        $SiteCode = $Site.SiteCode

        $Description = "$UserName,$SiteCode,$SerialNumber,$IpAddress,$Vendor,$Model"

        $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectCategory=computer)(name=$env:COMPUTERNAME))"
        $SearchRoot = "DC=$($env:USERDNSDOMAIN.replace('.',',DC='))"
        $Searcher.SearchRoot = "LDAP://$SearchRoot"
        $ComputerObject = $Searcher.FindAll()
        $ComputerEntry = $ComputerObject.GetDirectoryEntry()

        $ExistingDescription = $ComputerEntry | Select-Object -ExpandProperty description

        $MatchExisting = [regex]::IsMatch($Description,[regex]::Escape($ExistingDescription))

        if (-not $MatchExisting) {

            $ComputerEntry.description = $Description
            $ComputerEntry.SetInfo()

        } #if (-not $MatchExisting)

        New-Object -TypeName psobject -Property @{  ExistingDescription = $ExistingDescription
                                                    NewDescription = $Description
                                                    MatchExisting = $MatchExisting }

    } #if $OnDomain

} #process