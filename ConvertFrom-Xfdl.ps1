function ConvertFrom-Xfdl {
    <#
    .SYNOPSIS
        Extracts data from XFDL forms
    .DESCRIPTION
        Extracts data from XFDL forms
    .PARAMETER Path
        Path to XFDL forms to be extracted
    .PARAMETER ToXml
        Switch to output the data as a [xml]
    .PARAMETER ToText
        Switch to output the data as a [string]
    .EXAMPLE
        ConvertFrom-Xfdl -Path \\path\to\xfdl
        Finds all XFDL files in \\path\to\xfdl and extracts the data from them as [string]
    .EXAMPLE
        Get-ChildItem -Path \\path\to\xfdl | ConvertFrom-Xfdl -ToXml
        Extracts the data from XFDL files from the pipeline and outputs them as [xml]
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
    .LINK
        https://www.w3.org/TR/NOTE-XFDL
    .LINK
        https://www.ibm.com/support/knowledgecenter/SSS28S_3.5.0/com.ibm.form.designer.xfdl.doc/XFDL.pdf
    #>
    # Q2lwaGVyU2NydXBsZXM=
    [CmdletBinding(DefaultParameterSetName='Text')]
    [OutputType([string],ParameterSetName='Text')]
    [OutputType([xml],ParameterSetName='Xml')]
    param (
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath','FullName')]
        [System.Object[]]
        $Path,
        [Parameter(Mandatory=$true,ParameterSetName='Xml')]
        [switch]
        $ToXml,
        [Parameter(ParameterSetName='Text')]
        [switch]
        $ToText
    ) #param
    begin {
        function Expand-GzByteArray {
            [CmdletBinding()]
            [OutputType([byte[]])]
            param (
                [Parameter(Mandatory=$true)]
                [ValidateNotNullOrEmpty()]
                [byte[]]
                $InputByteArray
            ) #param
            process {
                $InputStream = New-Object -TypeName System.IO.MemoryStream -ArgumentList ( ,$InputByteArray)
                $OutputStream = New-Object -TypeName System.IO.MemoryStream
                $SplatArgs = @{ TypeName = 'System.IO.Compression.GzipStream'
                                ArgumentList = ($InputStream,[IO.Compression.CompressionMode]::Decompress) }
                $GzipStream = New-Object @SplatArgs
                $GzipStream.CopyTo( $OutputStream )
                $GzipStream.Close()
                $InputStream.Close()
                [byte[]] $OutputByteArray = $OutputStream.ToArray()
                Write-Output -InputObject $OutputByteArray
            } #process
        } #function Expand-GzByteArray
    } #begin
    process {
        $Path | ForEach-Object {
            $EachPath = $_ | Get-Item
            if (-not $EachPath.PSIsContainer) {
                $XfdlFiles = $EachPath | Where-Object { $_.Name -match '\.xfdl$' }
            } elseif ($EachPath.PSIsContainer) {
                $XfdlFiles = $EachPath | Get-ChildItem -Recurse | Where-Object { $_.Name -match '\.xfdl$' }
            } #if
            $XfdlFiles | ForEach-Object {
                $XfdlContent = Get-Content -Path $_.PSPath
                $RawXfdlContent = Get-Content -Path $_.PSPath -Raw
                if ($XfdlContent[0] -match '^<\?xml') {
                    $Xml = [xml] $RawXfdlContent
                } elseif ($XfdlContent[0] -match '^application/vnd\.xfdl;\ content-encoding="base64-gzip"') {
                    $Base64GzippedBytes = $XfdlContent | Select-Object -Skip 1
                    [byte[]] $GzippedBytes = ([System.Convert]::FromBase64String($Base64GzippedBytes))
                    $Bytes = (Expand-GzByteArray -InputByteArray $GzippedBytes)
                    $InnerText = ([System.Text.Encoding]::ASCII.GetString($Bytes))
                    $Xml = [xml] $InnerText
                } #if
                if ($PSCmdlet.ParameterSetName -match 'Text') {
                    $ExtractedText = $Xml.XFDL.page.label.value
                    Write-Output -InputObject $ExtractedText
                } elseif ($PSCmdlet.ParameterSetName -match 'Xml') {
                    Write-Output -InputObject $Xml
                } #if
            } #ForEach
        } #ForEach
    } #process
} #function ConvertFrom-Xfdl
