function Get-FileShareAudit {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true,Position=0,HelpMessage="Enter the UNC path of the share you want to audit permissions for.")]
        [string]
        $SharePath
    ) #param
    process {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $Context = [System.DirectoryServices.AccountManagement.ContextType]::Domain
        Get-Acl -Path $SharePath |
        Select-Object -ExpandProperty Access |
        Where-Object { $_.IdentityReference -match "^$env:USERDOMAIN\\" } | ForEach-Object {
            if ($_.FileSystemRights -match 'FullControl') {
                New-Object -TypeName psobject -Property @{
                    GroupName = $_.IdentityReference.Value.Split('\')[1]
                    Access = 'FullControl'
                }
            } elseif ($_.FileSystemRights -match 'Modify') {
                New-Object -TypeName psobject -Property @{
                    GroupName = $_.IdentityReference.Value.Split('\')[1]
                    Access = 'Modify'
                }
            } elseif ($_.FileSystemRights -match 'Read') {
                New-Object -TypeName psobject -Property @{
                    GroupName = $_.IdentityReference.Value.Split('\')[1]
                    Access = 'Read'
                }
            } #if
        } | ForEach-Object {
            $GroupName = $_.GroupName
            $GroupMembers = @()
            [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context,$GroupName).GetMembers() | ForEach-Object {
                if ($_.StructuralObjectClass -eq 'group') {
                    $_.GetMembers() | ForEach-Object {
                        if ($_.StructuralObjectClass -eq 'group') {
                            $_.GetMembers() | ForEach-Object {
                                if ($_.StructuralObjectClass -eq 'group') {
                                    $_.GetMembers() | ForEach-Object {
                                        if ($_.StructuralObjectClass -eq 'group') {
                                            $_.GetMembers() | ForEach-Object {
                                                if ($_.StructuralObjectClass -eq 'group') {
                                                    $_.GetMembers() | ForEach-Object {
                                                        $GroupMembers += $_.SamAccountName
                                                    } #ForEach 6th level
                                                } else {
                                                    $GroupMembers += $_.SamAccountName
                                                } #if
                                            } #ForEach 5th level
                                        } else {
                                            $GroupMembers += $_.SamAccountName
                                        } #if
                                    } #ForEach 4th level
                                } else {
                                    $GroupMembers += $_.SamAccountName
                                } #if
                            } #ForEach 3rd level
                        } else {
                            $GroupMembers += $_.SamAccountName
                        } #if
                    } #ForEach 2nd level
                } else {
                    $GroupMembers += $_.SamAccountName
                } #if
            } #ForEach 1st level
            $GroupMembers = $GroupMembers | Sort-Object -Unique
            Add-Member -InputObject $_ -MemberType NoteProperty -Name GroupMembers -Value $GroupMembers -PassThru
        } | ForEach-Object {
            $Access = $_.Access
            $_.GroupMembers | ForEach-Object {
                New-Object -TypeName psobject -Property @{
                    UserName = $_
                    Access = $Access
                    SharePath = $SharePath
                }
            }
        } | Group-Object -Property UserName | ForEach-Object {
            $UserName = $_.Name
            $Group = $_.Group
            $SharePath = $Group[0].SharePath
            $Access = ''
            $Group | ForEach-Object {
                $Access = "$Access$($_.Access),"
            } #ForEach
            $Access = $Access.Trim(',')
            New-Object -TypeName psobject -Property @{
                UserName = $UserName
                Access = $Access
                SharePath = $SharePath
            }
        } | Select-Object -Property UserName,Access,SharePath | Sort-Object -Property UserName
    } #process
} #function Get-FileShareAudit