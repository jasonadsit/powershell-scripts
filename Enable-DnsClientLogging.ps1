function Enable-DnsClientLogging {

	<#
	.SYNOPSIS
		Enables DNS client logging on endpoints.
	.DESCRIPTION
		Enables DNS client logging.
		Sets the max log size to 128 MB.
		Gives a cut & paste example of how to retrieve filtered events.
	.PARAMETER ComputerName
		The computer name to act on. Defaults to localhost (actually... `$env:COMPUTERNAME)
		Accepts values from the pipeline.
	.EXAMPLE
		Enable-DnsClientLogging -ComputerName host1
		Enables DNS client logging on host1.
	.EXAMPLE
		'host1','host2','host3' | Enable-DnsClientLogging
		Enables DNS client logging on the three host input via the pipeline.
	.NOTES
		###################################################################################
		Author:     @oregon-national-guard/systems-administration
		Version:    1.0
		###################################################################################
		License:    https://github.com/oregon-national-guard/powershell/blob/master/LICENCE
		###################################################################################
	.LINK
		https://github.com/oregon-national-guard
	.LINK
		https://creativecommons.org/publicdomain/zero/1.0/
	#>

	[CmdletBinding()]

	param (

		[parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
		[Alias('PSComputerName','DNSHostName','CN','Hostname')]
		[string[]]
		$ComputerName = @($env:COMPUTERNAME)

	) #param

	begin {} #begin

	process {

		$AllTheComputers = @()

		$ComputerName | ForEach-Object {

			$EachComputer = $_

			Write-Verbose "Begin building ScriptBlock to enable the Microsoft-Windows-DNS-Client/Operational event log"

			$EnableDnsClientEventLogScriptBlock = {

				$logName = 'Microsoft-Windows-DNS-Client/Operational'
				$log = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $logName
				$log.IsEnabled=$true
				$log.MaximumSizeInBytes=131072000
				$log.SaveChanges()

			}

			Write-Verbose "End building scriptblock to enable the Microsoft-Windows-DNS-Client/Operational event log"
			Write-Verbose "Begin execution of scriptblock on remote machine"

			Invoke-Command -ComputerName $_ -ScriptBlock $EnableDnsClientEventLogScriptBlock

			$AllTheComputers += $EachComputer

			Write-Verbose "End execution of scriptblock on remote machine"

		} # End of ForEach

	} #process

	end {

		Write-Host ''
		Write-Host ''
		Write-Host "DNS client event logging has been enabled on the following computers:" -ForegroundColor "Green"
		Write-Host ''
		Write-Host "$AllTheComputers" -ForegroundColor "Cyan"
		Write-Host ''
		Write-Host 'You can now retrieve DNS client events like' -ForegroundColor "Green"
		Write-Host 'in the following example (copy & paste):' -ForegroundColor "Green"
		Write-Host ''
		Write-Host '$DnsClientEventFilter = [xml]@"'
		Write-Host '<QueryList>'
		Write-Host "  <Query Id='0' Path='Microsoft-Windows-DNS-Client/Operational'>"
		Write-Host "    <Select Path='Microsoft-Windows-DNS-Client/Operational'>*</Select>"
		Write-Host '  </Query>'
		Write-Host '</QueryList>'
		Write-Host '"@'
		Write-Host ''
		Write-Host "Get-WinEvent -ComputerName $($AllTheComputers[0]) -FilterXml `$DnsClientEventFilter"
		Write-Host ''
		Write-Host 'Happy hunting!' -ForegroundColor "Green"
		Write-Host ''
		Write-Host ''

	} #end

} #function Enable-DnsClientLogging