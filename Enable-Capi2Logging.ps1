function Enable-Capi2Logging {

	<#
	.SYNOPSIS
		Enables the Capi2 event log.
	.DESCRIPTION
		Enables the Capi2 event log.
	.PARAMETER ComputerName
		The ComputerName to enable the Capi2 log on.
	.EXAMPLE
		Enable-Capi2Logging -ComputerName host1
		Enable Capi2 logging on host1.
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
	#>

	param (

		[parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
		[Alias('PSComputerName','DNSHostName','CN','Hostname')]
		[string[]]
		$ComputerName = @($env:COMPUTERNAME)

	) #param

	process {

		$AllTheComputers = @()

		$ComputerName | ForEach-Object {

			$EachComputer = $_

			Write-Verbose "Begin building ScriptBlock to enable the Microsoft-Windows-CAPI2/Operational event log"

			$EnableCapi2EventLogScriptBlock = {

				$logName = 'Microsoft-Windows-CAPI2/Operational'
				$log = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $logName
				$log.IsEnabled=$true
				$log.MaximumSizeInBytes=67108864
				$log.SaveChanges()

			} #$EnableCapi2EventLogScriptBlock

			Write-Verbose "End building scriptblock to enable the Microsoft-Windows-CAPI2/Operational event log"
			Write-Verbose "Begin execution of scriptblock on remote machine"

			Invoke-Command -ComputerName $_ -ScriptBlock $EnableCapi2EventLogScriptBlock

			$AllTheComputers += $EachComputer

			Write-Verbose "End execution of scriptblock on remote machine"

		} #ForEach

	} #process

	end {

		Write-Host ''
		Write-Host ''
		Write-Host "CAPI2 event logging has been enabled on the following computers:" -ForegroundColor "Green"
		Write-Host ''
		Write-Host "$AllTheComputers" -ForegroundColor "Cyan"
		Write-Host ''
		Write-Host 'You can now retrieve CAPI2 events like' -ForegroundColor "Green"
		Write-Host 'in the following example (copy & paste):' -ForegroundColor "Green"
		Write-Host ''
		Write-Host '$Capi2EventFilter = [xml]@"'
		Write-Host '<QueryList>'
		Write-Host "  <Query Id='0' Path='Microsoft-Windows-CAPI2/Operational'>"
		Write-Host "    <Select Path='Microsoft-Windows-CAPI2/Operational'>"
		Write-Host "      *[System[(Level=1  or Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) &lt;= 86400000]]]"
		Write-Host '    </Select>'
		Write-Host '  </Query>'
		Write-Host '</QueryList>'
		Write-Host '"@'
		Write-Host ''
		Write-Host "Get-WinEvent -ComputerName $($AllTheComputers[0]) -FilterXml `$Capi2EventFilter"
		Write-Host ''
		Write-Host 'This would get you Crypto API errors and warnings from the last 24 hours' -ForegroundColor "Green"
		Write-Host ''
		Write-Host 'Happy hunting!' -ForegroundColor "Green"
		Write-Host ''
		Write-Host ''

	} #end

} #function Enable-Capi2Logging