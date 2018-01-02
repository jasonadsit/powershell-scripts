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