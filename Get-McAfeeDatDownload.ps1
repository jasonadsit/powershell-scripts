#Get-McAfeeDatDownload
[CmdletBinding()]
[OutputType([System.IO.FileInfo])]
param (
    $BaseUrl = 'http://download.nai.com/products/licensed/superdat/english/intel'
) #param
process {
    $DatFile = ((Invoke-WebRequest -UseBasicParsing -Uri $BaseUrl).Links.href |
    Where-Object { $_ -match '\.exe$' } |
    Sort-Object -Descending)[0].Split('/')[-1]
    $DatVersion = [int] $($DatFile.Split('x')[0])
    $SourceDat = "$BaseUrl/$DatFile"
    $LocalDat = "$env:ProgramFiles\$DatFile"
    (New-Object System.Net.WebClient).DownloadFile("$SourceDat","$LocalDat")
    while (-not (Test-Path -Path $LocalDat)) { Start-Sleep -Seconds 5 }
    Start-Sleep -Seconds 3
    Unblock-File -Path $LocalDat
    $Hash = (Get-FileHash -Path $LocalDat -Algorithm SHA1).Hash
    $RenamedDatFile = "$($DatFile.Split('.')[0])_$Hash`.exe"
    $RenamedLocalDat = "$env:ProgramFiles\$RenamedDatFile"
    Rename-Item -Path $LocalDat -NewName $RenamedLocalDat | Out-Null
    Get-Item -Path $RenamedLocalDat
} #process
#Get-McAfeeDatDownload