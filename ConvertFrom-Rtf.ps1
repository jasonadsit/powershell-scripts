#Requires -Version 3.0
function ConvertFrom-Rtf {

    <#
    .SYNOPSIS
        Converts richtext (rtf) documents to plaintext
    .DESCRIPTION
        Converts richtext (rtf) documents to plaintext and outputs
        text strings or psobjects based on a switch.
    .EXAMPLE
        Get-ChildItem 'C:\Windows'  -Recurse `
                                    -File `
                                    -Filter *.rtf `
                                    -ErrorAction SilentlyContinue |
                                    Select-Object -First 5 |
                                    ConvertFrom-Rtf
        Converts first five rtf files in C:\Windows to plaintext
    .INPUTS
        System.IO.FileInfo
    .OUTPUTS
        System.Management.Automation.PSCustomObject
    .PARAMETER PathToFile
        System.IO.FileInfo objects to operate on
    .PARAMETER AsObject
        Switch to output psobjects instead of strings
    .PARAMETER Hash
        Switch to hash files or not
    .PARAMETER epoch
        DateTime of epoch (01/01/1970)
    .NOTES
        #######################################################################################
        Author:     @oregon-national-guard/systems-administration
        Version:    1.0
        #######################################################################################
        License:    https://github.com/oregon-national-guard/HunterGatherer/blob/master/LICENCE
        #######################################################################################
    .LINK
        https://github.com/oregon-national-guard
    .LINK
        https://creativecommons.org/publicdomain/zero/1.0/
    .LINK
        http://blog.kmsigma.com/2014/10/01/converting-rtf-to-txt-via-powershell/
    #>

[CmdletBinding()]

param (

    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [Alias('PSPath','FullName')]
    [System.IO.FileInfo[]]
    $PathToFile = @((Get-Item 'C:\Windows\Help\en-US\credits.rtf')),

    [switch]
    $AsObject,

    [switch]
    $Hash,

    [datetime]
    $epoch = '01/01/1970'

) #param

begin {

    $StartTime = (Get-Date)

    $NameOfFunction = 'ConvertFrom-Rtf'

    $CountFiles = ($PathToFile | Measure-Object).Count

    $ErrorActionPreference = 'Stop'

    Add-Type -AssemblyName System.Windows.Forms

    #
    # Thank @kmsigma for the rtf conversion part.
    # I just adapted it for the pipeline and put some
    # window dressing on it to fit my use case.
    #

} #begin

process {

    $PathToFile | ForEach-Object {

        $RichTextBox = New-Object -TypeName System.Windows.Forms.RichTextBox

        try {

            $RichTextBox.Rtf = [System.IO.File]::ReadAllText($_.FullName)

        } catch {

            Remove-Variable RichTextBox -ErrorAction SilentlyContinue

            $ExceptionName = $_.Exception.GetType().FullName

            $RichTextBox = New-Object -TypeName psobject -Property @{
                "Text" = "#!#!# Failed to convert with a $ExceptionName exception #!#!#"
            }

        } finally {

            if ($AsObject) {

                if ($Hash) {

                    New-Object -TypeName psobject -Property @{

                        "Name" = $_.Name
                        "FullName" = $_.FullName
                        "Text" = $RichTextBox.Text
                        "LastWriteTime" = $(($_.LastWriteTime) -as [datetime])
                        "EpochTime" = $(
                            (New-TimeSpan -Start $epoch -End (
                                [system.timezoneinfo]::ConvertTime(
                                    ($_.LastWriteTime),([system.timezoneinfo]::UTC)
                                ))
                            ).TotalSeconds
                        )
                        "TimeStamp" = $((($_.LastWriteTime).GetDateTimeFormats('s')) -as [string])
                        "Length" = $_.Length
                        "SHA256" = $((Get-FileHash -Path $_.FullName -Algorithm SHA256).Hash)

                    } |

                    Select-Object -Property LastWriteTime,
                                            TimeStamp,
                                            EpochTime,
                                            Name,
                                            FullName,
                                            Length,
                                            SHA256,
                                            Text

                    Remove-Variable RichTextBox -ErrorAction SilentlyContinue

                } elseif (!($Hash)) {

                    New-Object -TypeName psobject -Property @{

                        "Name" = $_.Name
                        "FullName" = $_.FullName
                        "Text" = $RichTextBox.Text
                        "LastWriteTime" = $(($_.LastWriteTime) -as [datetime])
                        "EpochTime" = $(
                            (New-TimeSpan -Start $epoch -End (
                                [system.timezoneinfo]::ConvertTime(
                                    ($_.LastWriteTime),([system.timezoneinfo]::UTC)
                                ))
                            ).TotalSeconds
                        )
                        "TimeStamp" = $((($_.LastWriteTime).GetDateTimeFormats('s')) -as [string])
                        "Length" = $_.Length

                    } |

                    Select-Object -Property LastWriteTime,
                                            TimeStamp,
                                            EpochTime,
                                            Name,
                                            FullName,
                                            Length,
                                            Text

                    Remove-Variable RichTextBox -ErrorAction SilentlyContinue

                } #if ($Hash)

            } elseif (!($AsObject)) {

                $RichTextBox.Text

                Remove-Variable RichTextBox -ErrorAction SilentlyContinue

            } # if else (`$AsObject)

        } #try catch finally

    } #ForEach

} #process

end {

    $EndTime = (Get-Date)

    Write-Verbose "Finished running $NameOfFunction on $CountFiles files"

    Write-Verbose "Elapsed Time: $(($EndTime - $StartTime).TotalSeconds) seconds"

} #end

} #function ConvertFrom-Rtf