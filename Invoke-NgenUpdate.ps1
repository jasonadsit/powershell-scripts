function Invoke-NgenUpdate {
    Resolve-Path -Path "$env:SystemDrive\Windows\Microsoft.NET\Framework*\*\ngen.exe" | ForEach-Object {
            &$_.Path update /force /nologo /silent 2>&1>$null
    } #ForEach
} #function Invoke-NgenUpdate
