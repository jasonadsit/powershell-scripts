function Get-ObjectAccessAudit {

    [CmdletBinding()]

    param (

        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [string[]]
        $ComputerName,

        [string]
        $UserName,

        [string]
        $ObjectName,

        [parameter()]
        [ValidateSet(
            'ACCESS_SYS_SEC','AppendData','DELETE','DeleteChild','Execute',
            'READ_CONTROL','ReadAttributes','ReadData','ReadEA','SYNCHRONIZE',
            'WRITE_DAC','WRITE_OWNER','WriteAttributes','WriteData','WriteEA'
        )]
        [string]
        $AccessType

    ) #param

    begin {

        Write-Verbose -Message 'Determining filter based on parameter binding'
        if ($PSBoundParameters.ContainsKey('ComputerName')) {
            Write-Debug -Message '$ComputerName parameter was bound at runtime'
        } elseif (-not ($PSBoundParameters.ContainsKey('ComputerName'))) {
            Write-Debug -Message '$ComputerName parameter was not bound at runtime. Setting it to $env:COMPUTERNAME'
            $ComputerName = $env:COMPUTERNAME
        } #if $PSBoundParameters.ContainsKey('ComputerName')
        if ($PSBoundParameters.ContainsKey('UserName')) {
            Write-Debug -Message '$UserName parameter was bound at runtime'
            $User = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('UserName'))) {
            Write-Debug -Message '$UserName parameter was not bound at runtime'
            $User = $false
        } #if $PSBoundParameters.ContainsKey('UserName')
        if ($PSBoundParameters.ContainsKey('ObjectName')) {
            $Object = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('ObjectName'))) {
            $Object = $false
        } #if $PSBoundParameters.ContainsKey('ObjectName')
        if ($PSBoundParameters.ContainsKey('AccessType')) {
            Write-Debug -Message '$AccessType parameter was bound at runtime'
            $Access = $true
            Write-Debug -Message 'Translating AccessType into AccessMask'
            $AccessMask = switch -Exact ($AccessType) {
                'ReadData' {'0x1'}
                'WriteData' {'0x2'}
                'AppendData' {'0x4'}
                'ReadEA' {'0x8'}
                'WriteEA' {'0x10'}
                'Execute' {'0x20'}
                'DeleteChild' {'0x40'}
                'ReadAttributes' {'0x80'}
                'WriteAttributes' {'0x100'}
                'DELETE' {'0x10000'}
                'READ_CONTROL' {'0x20000'}
                'WRITE_DAC' {'0x40000'}
                'WRITE_OWNER' {'0x80000'}
                'SYNCHRONIZE' {'0x100000'}
                'ACCESS_SYS_SEC' {'0x1000000'}
            } #$AccessMask
        } elseif (-not ($PSBoundParameters.ContainsKey('AccessType'))) {
            Write-Debug -Message '$AccessType parameter was not bound at runtime'
            $Access = $false
        } #if $PSBoundParameters.ContainsKey('AccessType')
        if ($User -and $Object -and $Access) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='SubjectUserName']='$UserName']]
      and
      *[EventData[Data[@Name='ObjectName']="$ObjectName"]]
      and
      *[EventData[Data[@Name='AccessMask']='$AccessMask']]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ($User -and $Object -and (-not $Access)) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='SubjectUserName']='$UserName']]
      and
      *[EventData[Data[@Name='ObjectName']="$ObjectName"]]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ($User -and (-not $Object) -and $Access) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='SubjectUserName']='$UserName']]
      and
      *[EventData[Data[@Name='AccessMask']='$AccessMask']]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ((-not $User) -and $Object -and $Access) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='ObjectName']="$ObjectName"]]
      and
      *[EventData[Data[@Name='AccessMask']='$AccessMask']]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ($User -and (-not $Object) -and (-not $Access)) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='SubjectUserName']='$UserName']]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ((-not $User) -and $Object -and (-not $Access)) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='ObjectName']="$ObjectName"]]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ((-not $User) -and (-not $Object) -and $Access) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
      and
      *[EventData[Data[@Name='AccessMask']='$AccessMask']]
    </Select>
  </Query>
</QueryList>
"@
        } elseif ((-not $User) -and (-not $Object) -and (-not $Access)) {
            $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4663)]]
    </Select>
  </Query>
</QueryList>
"@
        } #if ($User -and $Object -and $Access)

        $QueryXml = [xml]$QueryString

    } #begin

    process {

        $ComputerName | ForEach-Object {

            $EachComputer = $_

            if ($EachComputer -match $env:COMPUTERNAME) {

                Get-WinEvent -FilterXml $QueryXml | ForEach-Object {
                    $EventXml = [xml]$_.ToXml()
                    $XmlData = $null
                    if ($XmlData = @($EventXml.Event.EventData.Data)) {
                        for ($i=0; $i -lt $XmlData.Count; $i++) {
                            $SplatArgs = @{
                                InputObject = $_
                                MemberType = "NoteProperty"
                                Name = "$($XmlData[$i].name)"
                                Value = "$($XmlData[$i].'#text')"
                                Force = $true
                                Passthru = $true
                            }
                            $_ = Add-Member @SplatArgs
                        } #for
                    } #if
                    $_
                } #ForEach Event

            } elseif ($EachComputer -notmatch $env:COMPUTERNAME) {

                Get-WinEvent -ComputerName $EachComputer -FilterXml $QueryXml | ForEach-Object {
                    $EventXml = [xml]$_.ToXml()
                    $XmlData = $null
                    if ($XmlData = @($EventXml.Event.EventData.Data)) {
                        for ($i=0; $i -lt $XmlData.Count; $i++) {
                            $SplatArgs = @{
                                InputObject = $_
                                MemberType = "NoteProperty"
                                Name = "$($XmlData[$i].name)"
                                Value = "$($XmlData[$i].'#text')"
                                Force = $true
                                Passthru = $true
                            }
                            $_ = Add-Member @SplatArgs
                        } #for
                    } #if
                    $_
                } #ForEach Event

            } #if $EachComputer -match $env:COMPUTERNAME

        } #ForEach ComputerName

    } #process
    end {} #end
} #function Get-ObjectAccessAudit