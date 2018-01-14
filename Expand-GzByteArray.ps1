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