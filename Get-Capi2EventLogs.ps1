function Get-Capi2EventLogs {

    [CmdletBinding()]

    param (

        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [string[]]
        $ComputerName = $env:COMPUTERNAME

    ) #param

    begin {

        if (-not ($PSBoundParameters.ContainsKey('ComputerName'))) {

            $ComputerName = $env:COMPUTERNAME

        } #if

        $Capi2EventFilter = [xml] @"
<QueryList>
  <Query Id='0' Path='Microsoft-Windows-CAPI2/Operational'>
    <Select Path='Microsoft-Windows-CAPI2/Operational'>
      *[System[(Level=1  or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) &lt;= 86400000]]]
    </Select>
  </Query>
</QueryList>
"@

    } #begin

    process {

        $ComputerName | ForEach-Object {

            $EachComputer = $_

            if ($EachComputer -match $env:COMPUTERNAME) {

                $SplatArgs = @{ FilterXml = $Capi2EventFilter }

            } elseif (-not ($EachComputer -match $env:COMPUTERNAME)) {

                $SplatArgs = @{ ComputerName = $ComputerName
                                FilterXml = $Capi2EventFilter }

            } #if

            Get-WinEvent @SplatArgs | ForEach-Object {

                $EventXml = [xml]$_.ToXml()
                $EventXml = $EventXml.Event.UserData.CertVerifyCertificateChainPolicy

                $ServerName = $EventXml.SSLAdditionalPolicyInfo.serverName
                $ResultText = $EventXml.Result.'#text'
                $ProcessName = $EventXml.EventAuxInfo.ProcessName
                $Certificate = $EventXml.Certificate.fileRef
                $SubjectName = $EventXml.Certificate.subjectName
                $CorrelationTaskId = $EventXml.CorrelationAuxInfo.TaskId

                New-Object -TypeName psobject -Property @{  'TimeCreated' = $_.TimeCreated
                                                            'Id' = $_.Id
                                                            'TaskDisplayName' = $_.TaskDisplayName
                                                            'LevelDisplayName' = $_.LevelDisplayName
                                                            'MachineName' = $_.MachineName
                                                            'ProcessId' = $_.ProcessId
                                                            'ThreadId' = $_.ThreadId
                                                            'ProcessName' = $ProcessName
                                                            'ServerName' = $ServerName
                                                            'SubjectName' = $SubjectName
                                                            'Certificate' = $Certificate
                                                            'ResultText' = $ResultText
                                                            'CorrelationTaskId' = $CorrelationTaskId }

            } | Select-Object -Property TimeCreated,Id,TaskDisplayName,LevelDisplayName,
                                        MachineName,ProcessId,ThreadId,ProcessName,ServerName,
                                        SubjectName,Certificate,ResultText,CorrelationTaskId

        } #ForEach ComputerName

    } #process

    end {} #end

} #function Get-Capi2EventLogs