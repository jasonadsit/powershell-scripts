#script Get-DellServerCatalogXml
[CmdletBinding()]
param (
    [string] $Uri = 'https://downloads.dell.com/catalog/Catalog.xml.gz',
    [int] $CharsToSkip = 42
) #param
begin {
    function Get-DecompressedByteArray {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
            [byte[]] $byteArray = $(Throw("-byteArray is required"))
        ) #param
        process {
            Write-Verbose "Get-DecompressedByteArray"
            $input = New-Object System.IO.MemoryStream( , $byteArray )
            $output = New-Object System.IO.MemoryStream
            $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
            $gzipStream.CopyTo( $output )
            $gzipStream.Close()
            $input.Close()
            [byte[]] $byteOutArray = $output.ToArray()
            Write-Output $byteOutArray
        } #process
    } #function Get-DecompressedByteArray
} #begin
process {
    $CatalogCompressedBytes = (New-Object System.Net.WebClient).DownloadData("$Uri")
    $CatalogBytes = Get-DecompressedByteArray -byteArray $CatalogCompressedBytes
    $CatalogString = [System.Text.Encoding]::Unicode.GetString($CatalogBytes)
    $Catalog = $CatalogString.Substring($CharsToSkip)
    $CatalogXml = $Catalog -as [xml]
    Write-Output -InputObject $CatalogXml
} #process
#script Get-DellServerCatalogXml