function Get-ExampleFunction {

    # ALWAYS write comment-based help (see example below)
    # https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Core/about_Comment_Based_Help

    <#
    .SYNOPSIS
        Describe function.
    .DESCRIPTION
        Describe function in more detail.
    .EXAMPLE
        Give an example of how to use it.
    .EXAMPLE
        Give another example of how to use it.
    .PARAMETER ComputerName
        Describe each parameter.
    .PARAMETER AnotherParameter
        Describe each parameter.
    .INPUTS
        System.Object
    .OUTPUTS
        System.Object
    .NOTES
        #######################################################################################
        Author:     State of Oregon, OSCIO, ESO, Cybersecurity Assessment Team
        Version:    1.0
        #######################################################################################
        License:    https://unlicense.org/UNLICENSE
        #######################################################################################
    .LINK
        https://github.com/oregon-eso-cyber-assessments
    .FUNCTIONALITY
        Explain the intended use of the function.
    #>

    [CmdletBinding()]

    param (

        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [string[]]
        $ComputerName,

        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]
        $AnotherParameter = 'DefaultValueOfAnotherParameter'

    ) #param

    #
    # Terminate parenthesis, brackets, and braces with a meaningful comment.
    # It makes your code more readable and helps editors with code folding
    #

    begin {

        #
        # Set up to do the stuff
        # Maybe set-up some static global variables
        #

        if (-not $PSBoundParameters.ContainsKey('ComputerName')) {

            $ComputerName = $env:COMPUTERNAME

        } #if

        # Define local helper functions
        function Get-EpochTimeStamp {
            $Epoch = Get-Date -Date '01/01/1970'
            $TimeNow = Get-Date
            $TimeNowUtc = [system.timezoneinfo]::ConvertTime($TimeNow,[system.timezoneinfo]::UTC)
            $EpochSpan = New-TimeSpan -Start $Epoch -End $TimeNowUtc
            $EpochTimeStamp = [math]::Round($EpochSpan.TotalSeconds).ToString()
            $EpochTimeStamp
        } #function Get-EpochTimeStamp

        # Package local functions to pass to remote sessions

        $GetEpochTimeStampFunction = "function Get-EpochTimeStamp { ${function:Get-EpochTimeStamp} }"

        $StartTime = Get-EpochTimeStamp

        Write-Verbose -Message "$StartTime"

    } #begin
    
    process {

        #
        # Do the stuff
        #

        $ComputerName | ForEach-Object {

            $EachComputer = $_

            Write-Verbose -Message "$EachComputer"

            if ($EachComputer -match $env:COMPUTERNAME) {

                # Running locally

                Get-EpochTimeStamp

            } else {

                # Running remotely

                Invoke-Command -ComputerName $EachComputer -ScriptBlock {

                    param (

                        $NewAudioNotificationFunction,
                        $AnotherParameter

                    ) #param

                    # Defining the function in the remote session that was
                    # passed in by the $GetEpochTimeStampFunction parameter.

                    . ([ScriptBlock]::Create($GetEpochTimeStampFunction))

                    Get-EpochTimeStamp

                    Write-Verbose -Message "$AnotherParameter"

                } -ArgumentList $GetEpochTimeStampFunction,$AnotherParameter

            } #if

        } #ForEach-Object

    } #process

    end {

        #
        # Clean up from doing stuff
        #

        $EndTime = Get-EpochTimeStamp

        Write-Verbose -Message "$EndTime"

    } #end

} #function Get-ExampleFunction