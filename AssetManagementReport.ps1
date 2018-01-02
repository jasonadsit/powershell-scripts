[CmdletBinding()]
param () #param
process {
    $IfIndex = (Get-CimInstance -ClassName Win32_IP4RouteTable -Filter "Destination='0.0.0.0'").InterfaceIndex[0]
    $IpAddress = (Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "InterfaceIndex=$IfIndex").IPAddress[0]
    $WmiComputerSystemProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct
    $SerialNumber = $WmiComputerSystemProduct.IdentifyingNumber.Trim()
    $Vendor = $WmiComputerSystemProduct.Vendor.Split(' ')[0].Trim().Trim(',')
    $Model = $WmiComputerSystemProduct.Name.Trim()
    $MacAddress = (Get-CimInstance -ClassName Win32_NetworkAdapter | Where-Object { $_.InterfaceIndex -eq $IfIndex }).MACAddress
    New-Object -TypeName psobject -Property @{
        ComputerName = [string] $env:COMPUTERNAME
        IpAddress = [string] $IpAddress
        MacAddress = [string] $MacAddress
        SerialNumber = [string] $SerialNumber
        Vendor = [string] $Vendor
        Model = [string] $Model
    } | Select-Object -Property ComputerName,IpAddress,MacAddress,SerialNumber,Vendor,Model
} #process