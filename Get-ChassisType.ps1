function Get-ChassisType {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [System.Object[]]
        $ComputerName
    ) #param
    begin {
        if (-not ($PSBoundParameters.ContainsKey('ComputerName'))) {
            $Computer = $false
        } elseif ($PSBoundParameters.ContainsKey('ComputerName')) {
            $Computer = $true
        } #if
        if (-not $Computer) {
            $ComputerName = $env:COMPUTERNAME
        } #if
    } #begin
    process {
        $ComputerName | ForEach-Object {
            $EachComputer = $_
            if ($EachComputer -match $env:COMPUTERNAME) {
                $SplatArgs = @{ ClassName = 'CIM_Chassis'
                                Property = 'ChassisTypes'
                                ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue }
            } elseif (-not ($EachComputer -match $env:COMPUTERNAME)) {
                $SplatArgs = @{ ComputerName = $EachComputer
                                ClassName = 'CIM_Chassis'
                                Property = 'ChassisTypes'
                                ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue }
            } #if
            Get-CimInstance @SplatArgs | Select-Object -Property ChassisTypes | ForEach-Object {
                New-Object -TypeName psobject -Property @{
                    ComputerName = $EachComputer
                    ChassisType = $_ | Select-Object -ExpandProperty ChassisTypes
                }
            } | ForEach-Object {
                $Chassis = switch -Exact ($_.ChassisType) {
                    0 {'Other'}
                    1 {'Unknown'}
                    3 {'Desktop'}
                    4 {'Low Profile Desktop'}
                    5 {'Pizza Box'}
                    6 {'Mini Tower'}
                    7 {'Tower'}
                    8 {'Portable'}
                    9 {'Laptop'}
                    10 {'Notebook'}
                    11 {'Hand-held'}
                    12 {'Docking Station'}
                    13 {'All-in-one'}
                    14 {'Sub notebook'}
                    15 {'Space-saving'}
                    16 {'Lunch Box'}
                    17 {'Main System Chassis'}
                    18 {'Expansion chassis'}
                    19 {'Sub chassis'}
                    20 {'Bus Expansion Chassis'}
                    21 {'Peripheral Chassis'}
                    22 {'Storage chassis'}
                    23 {'Rack mount chassis'}
                    24 {'Sealed-case PC'}
                }
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Chassis -Value $Chassis
                $_ | Select-Object -Property ComputerName,ChassisType,Chassis
            } #ForEach ChassisType
        } #ForEach ComputerName
    } #process
    end {} #end
} #function Get-ChassisType