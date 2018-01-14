#script Get-DellServerCatalogXml
[CmdletBinding()]
param (
    [string] $Uri = 'https://downloads.dell.com/catalog/Catalog.xml.gz',
    [int] $CharsToSkip = 42
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
    $CatalogCompressedBytes = (New-Object System.Net.WebClient).DownloadData("$Uri")
    $CatalogBytes = Expand-GzByteArray -InputByteArray $CatalogCompressedBytes
    $CatalogString = [System.Text.Encoding]::Unicode.GetString($CatalogBytes)
    $Catalog = $CatalogString.Substring($CharsToSkip)
    $CatalogXml = $Catalog -as [xml]
    Write-Output -InputObject $CatalogXml
} #process
#script Get-DellServerCatalogXml