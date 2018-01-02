function Get-TargetedWinEvent {

    <#
    .SYNOPSIS
        Searches Windows logs for events related to specific Event IDs or EventData.Data values
    .DESCRIPTION
        Searches Windows logs for events related to specific Event IDs or EventData.Data values
        Supports searching offline/exported evt/evtx files as well as online machines
    .PARAMETER SearchTerm
        EventData.Data property value to search for
    .PARAMETER EventId
        Windows Event ID to search for
    .PARAMETER ComputerName
        Remote computer to search logs on
        Defaults to $env:COMPUTERNAME
    .PARAMETER Credential
        Credential used to connect to remote computers if default credentials won't work
    .PARAMETER Offline
        Switch to search offline/exported logs either locally or remotely
    .PARAMETER Path
        Path to offline/exported log files
        Defaults to "$env:SystemRoot\System32\winevt\Logs"
    .PARAMETER Days
        Number of days to go back in the search for offline/exported logs
    .PARAMETER Flatten
        Switch to flatten the EventData.Data values and append them as properties to the parent object
    .EXAMPLE
        Get-TargetedWinEvent -SearchTerm user.name
        Finds local events related to user.name
    .EXAMPLE
        'host1','host2' | Get-TargetedWinEvent -SearchTerm user.name -EventId 4624
        Finds successful logon events for user.name from both hosts via the pipeline
    .EXAMPLE
        Get-TargetedWinEvent -EventId 4624 -SearchTerm user.name -Offline -Path \\server\share\LogArchive
        Finds successful logon events for user.name in all archived log files under \\server\share\LogArchive
    .NOTES
        #######################################################################################
        Author:     @oregon-national-guard/systems-administration
        Version:    1.0
        #######################################################################################
        License:    https://github.com/oregon-national-guard/powershell/blob/master/LICENCE
        #######################################################################################
    .LINK
        https://github.com/oregon-national-guard
    .LINK
        https://creativecommons.org/publicdomain/zero/1.0/
    #>

    [CmdletBinding(DefaultParameterSetName='Online')]

    param (

        [Parameter(ParameterSetName='Online')]
        [Parameter(ParameterSetName='Offline')]
        [string]
        $SearchTerm,

        [Parameter(ParameterSetName='Online')]
        [Parameter(ParameterSetName='Offline')]
        [string]
        $EventId,

        [Parameter(ParameterSetName='Online',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [System.Object[]]
        $ComputerName,

        [Parameter(ParameterSetName='Online')]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory=$true,ParameterSetName='Offline')]
        [switch]
        $Offline,

        [Parameter(ParameterSetName='Offline',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('PSPath','FullName','LogFilePath')]
        [System.Object[]]
        $Path,

        [Parameter(ParameterSetName='Offline')]
        [int]
        $Days,

        [Parameter(ParameterSetName='Online')]
        [Parameter(ParameterSetName='Offline')]
        [switch]
        $Flatten

    ) #param

    begin {

        Write-Debug -Message 'start of private function definitions'

        function FlattenWinEvent {
            [CmdletBinding()]
            param (
                [Parameter(ValueFromPipeline=$true)]
                [System.Diagnostics.Eventing.Reader.EventLogRecord[]]
                $EventLogRecord
            ) #param
            process {
                $EventLogRecord | ForEach-Object {
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
                } #ForEach EventLogRecord
            } #process
        } #function FlattenWinEvent

        Write-Debug -Message 'end of private function definitions'

        Write-Debug -Message 'start detecting which parameters were bound at runtime'

        if ($PSBoundParameters.ContainsKey('EventId')) {
            $Id = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('EventId'))) {
            $Id = $false
        } #if

        if ($PSBoundParameters.ContainsKey('SearchTerm')) {
            $Search = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('SearchTerm'))) {
            $Search = $false
        } #if

        if ($PSBoundParameters.ContainsKey('ComputerName')) {
            $Computer = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('ComputerName'))) {
            $Computer = $false
        } #if

        if ($PSBoundParameters.ContainsKey('Credential')) {
            $Creds = $true
        } elseif (-not ($PSBoundParameters.ContainsKey('Credential'))) {
            $Creds = $false
        } #if

        Write-Debug -Message 'end detecting which parameters were bound at runtime'

        Write-Debug -Message 'start setting default parameter values'

        if (-not $Computer) {

            $ComputerName = $env:COMPUTERNAME

        } #if

        if (-not ($PSBoundParameters.ContainsKey('Days'))) {
            $Days = 1
        } #if

        if (($PSCmdlet.ParameterSetName -match 'Offline') -and (-not ($PSBoundParameters.ContainsKey('Path')))) {

            Write-Debug -Message 'No $Path supplied. Using "$env:SystemRoot\System32\winevt\Logs"'

            $Path = "$env:SystemRoot\System32\winevt\Logs"

        } #if

        Write-Debug -Message 'end setting default parameter values'

    } #begin

    process {

        if ($PSCmdlet.ParameterSetName -match 'Online') {

            Write-Verbose -Message 'Running in "Online" mode'

            Write-Debug -Message 'start building "Here-Strings" for query XML'

            if ($Id -and $Search) {

                $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=$EventId)]]
      and
      *[EventData[Data and (Data='$SearchTerm')]]
    </Select>
  </Query>
</QueryList>
"@

            } elseif ($Id -and (-not $Search)) {

                $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=$EventId)]]
    </Select>
  </Query>
</QueryList>
"@

            } elseif ((-not $Id) -and $Search) {

                $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing']]]
      and
      *[EventData[Data and (Data='$SearchTerm')]]
    </Select>
  </Query>
</QueryList>
"@

            } elseif ((-not $Id) -and (-not $Search)) {

                $QueryString = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing']]]
    </Select>
  </Query>
</QueryList>
"@

            } #if ($Id -and $Search)

            Write-Debug -Message 'end building "Here-Strings" for query XML'

            $QueryXml = [xml]$QueryString

            $ComputerName | ForEach-Object {

                $EachComputer = $_

                if (-not $Computer) {

                    $SplatArgs = @{
                        FilterXml = $QueryXml
                        ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue
                    }

                    if (-not $Flatten) {

                        Get-WinEvent @SplatArgs

                    } elseif ($Flatten) {

                        Get-WinEvent @SplatArgs | FlattenWinEvent

                    } #if

                } elseif ($Computer -and (-not $Creds)) {

                    $SplatArgs = @{
                        ComputerName = $EachComputer
                        FilterXml = $QueryXml
                        ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue
                    }

                    if (-not $Flatten) {

                        Get-WinEvent @SplatArgs

                    } elseif ($Flatten) {

                        Get-WinEvent @SplatArgs | FlattenWinEvent

                    } #if

                } elseif ($Computer -and $Creds) {

                    $SplatArgs = @{
                        ComputerName = $EachComputer
                        FilterXml = $QueryXml
                        Credential = $Credential
                        ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue
                    }

                    if (-not $Flatten) {

                        Get-WinEvent @SplatArgs

                    } elseif ($Flatten) {

                        Get-WinEvent @SplatArgs | FlattenWinEvent

                    } #if

                } #if

            } #ForEach ComputerName

        } elseif ($PSCmdlet.ParameterSetName -match 'Offline') {

            Write-Verbose -Message 'Running in "Offline" mode'

            $Path | ForEach-Object {

                Write-Verbose -Message "looking for evt or evtx files in $_"

                $EachPath = $_ | Get-Item

                if (-not $EachPath.PSIsContainer) {

                    $LogFiles = $EachPath | Where-Object { $_.Name -match '\.(evt|evtx)$' }

                } elseif ($EachPath.PSIsContainer) {

                    $LogFiles = $EachPath | Get-ChildItem | Where-Object {
                        $_.Name -match '\.(evt|evtx)$' -and
                        $_.LastWriteTime -ge (Get-Date).AddDays(-$Days)
                    }

                } #if

                $LogFiles | ForEach-Object {

                    Write-Verbose -Message "now processing $($_.Name)"

                    Write-Debug -Message 'start building "Here-Strings" for query XML'

                    if ($Id -and $Search) {

                        $QueryString = @"
<QueryList>
  <Query Id="0" Path="file://$($_.FullName)">
    <Select Path="file://$($_.FullName)">
      *[System[(EventID=$EventId)]]
      and
      *[EventData[Data and (Data='$SearchTerm')]]
    </Select>
  </Query>
</QueryList>
"@

                    } elseif ($Id -and (-not $Search)) {

                        $QueryString = @"
<QueryList>
  <Query Id="0" Path="file://$($_.FullName)">
    <Select Path="file://$($_.FullName)">
      *[System[(EventID=$EventId)]]
    </Select>
  </Query>
</QueryList>
"@

                    } elseif ((-not $Id) -and $Search) {

                        $QueryString = @"
<QueryList>
  <Query Id="0" Path="file://$($_.FullName)">
    <Select Path="file://$($_.FullName)">
      *[EventData[Data and (Data='$SearchTerm')]]
    </Select>
  </Query>
</QueryList>
"@

                    } elseif ((-not $Id) -and (-not $Search)) {

                        $QueryString = @"
<QueryList>
  <Query Id="0" Path="file://$($_.FullName)">
    <Select Path="file://$($_.FullName)">
    </Select>
  </Query>
</QueryList>
"@

                    } #if ($Id -and $Search)

                    Write-Debug -Message 'end building "Here-Strings" for query XML'

                    $QueryXml = [xml]$QueryString

                    $SplatArgs = @{
                        FilterXml = $QueryXml
                        Oldest = $true
                        ErrorAction = [System.Management.Automation.ActionPreference]::SilentlyContinue
                    }

                    if (-not $Flatten) {

                        Get-WinEvent @SplatArgs

                    } elseif ($Flatten) {

                        Get-WinEvent @SplatArgs | FlattenWinEvent

                    } #if

                } #ForEach LogFile

            } #ForEach Path

        } #if ($PSCmdlet.ParameterSetName -match 'Online')

    } #process

    end {} #end

} #function Get-TargetedWinEvent