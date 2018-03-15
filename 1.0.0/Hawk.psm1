#############################################################################################
# DISCLAIMER:																				#
#																							#
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT					#
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY				#
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT		#
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR		#
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS		#
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR			#
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE	#
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS	#
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)	#
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,		#
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES						#
#############################################################################################



#Region Utility Functions
# ============== Utility Functions ==============


# Build an OauthToken
# https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/AzureAppEnumerationViaGraph.ps1
Function New-OauthToken 
{
	
	# Make sure we have a connection to msol since we needed it for this
	$null = Test-MSOLConnection

	[string]$TenantName = (Get-MsolCompanyInformation).initialdomain
	
	# See if we have the needed authentication package
	$TestModule = Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue


	# If we don't then we need to ask the user to install it
	if ($null -ne $TestModule) 
	{
		
		# Get ADAL path
		[array]$AzureADModule = (Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending)

		# Get the adal auth module from the azuread module
		$adal = ($AzureADModule[0].filelist | where { $_ -like "*Microsoft.IdentityModel.Clients.ActiveDirectory.dll"})
		$adalplatform = ($AzureADModule[0].filelist | where {$_ -like "*Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"})
				
		# Load the azuread auth module
		[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
		[System.Reflection.Assembly]::LoadFrom($adalplatform) | Out-Null
		
		# Build our values for connecting
		[string]$resourceAppIdURI = "https://graph.windows.net"
		[string]$clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
		[uri]$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
		$authority = "https://login.windows.net/$TenantName"

		# Build the auth context and our always prompt option
		$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
		$authParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList 1
		
		# Crease the async task
		$authtask = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId,$redirectUri,$authParameters)
		
		# Wait on our task and return our result
		$authtask.wait()
		return $authtask.result

	}
	else 
	{
		Out-LogFile "Please install the needed authentication module from an Elevated PowerShell Prompt: "
		Out-LogFile 'Install-Package -Name "Microsoft.IdentityModel.Clients.ActiveDirectory"'
		break		
	}
		

}

# Get the Location of an IP using the freegeoip.net rest API
Function Get-IPGeolocation
{

    Param
	(
	[Parameter(Mandatory=$true)]
	$IPAddress
	)

	# Check the global IP cache and see if we already have the IP there
	if ($IPLocationCache.ip -contains $IPAddress)
	{
		return ($IPLocationCache | Where-Object {$_.ip -eq $IPAddress } )
	}
	# If not then we need to look it up and populate it into the cache
	else 
	{
		
		# URI to pull the data from
		$resource = "http://freegeoip.net/xml/$IPAddress"

		# Return Data from web
		$Error.Clear()
		$geoip = Invoke-RestMethod -Method Get -URI $resource -ErrorAction SilentlyContinue
	
		if ($Error.Count -gt 0)
		{
			Out-LogFile ("Failed to retreive location for IP " + $IPAddress)
			$hash = @{
			IP = $IPAddress
			CountryName = "Failed to Resolve"
			RegionCode = "Unknown"
			RegionName = "Unknown"
			City = "Unknown"
			ZipCode = "Unknown"
			KnownMicrosoftIP = "Unknown"
			}
		}
		else 
		{
			# Sleep 1 second to be a good citizen this is a free resource
			Start-Sleep 1
			
			# Determine if this IP is known to be owned by Microsoft
			[string]$isMSFTIP = Test-MicrosoftIP -IP ($connection.clientip)
			
			# Push return into a response object
			$hash = @{
				IP = $geoip.Response.IP
				CountryName = $geoip.Response.CountryName
				RegionCode = $geoip.Response.RegionCode
				RegionName = $geoip.Response.RegionName
				City = $geoip.Response.City
				ZipCode = $geoip.Response.ZipCode
				KnownMicrosoftIP = $isMSFTIP
				}
			$result = New-Object PSObject -Property $hash
		}

		# Push the result to the global IPLocationCache
		[array]$Global:IPlocationCache += $result		

		# Return the result to the user
		return $result
	}
}

# Convert output from search-adminauditlog to be more human readable
Function Get-SimpleAdminAuditLog
{
	Param (
		[Parameter(
			Position=0,
			Mandatory=$true,
			ValueFromPipeline=$true,
			ValueFromPipelineByPropertyName=$true)
		]
		$SearchResults
	)

	# Setup to process incomming results
	Begin {
		
		# Make sure the array is null
		[array]$ResultSet = $null
		
	}

	# Process thru what ever is comming into the script
	Process {
		
		# Deal with each object in the input
		$searchresults | ForEach-Object {
			
			# Reset the result object
			$Result = New-Object PSObject
			
			# Get the alias of the User that ran the command
			[string]$user = $_.caller
			if ([string]::IsNullOrEmpty($user)){$user = "***"}
			else {$user = ($_.caller.split("/"))[-1]}
			
			# Build the command that was run
			$switches = $_.cmdletparameters
			[string]$FullCommand = $_.cmdletname
			
			# Get all of the switchs and add them in "human" form to the output
			foreach ($parameter in $switches){
								
				# Format our values depending on what they are so that they are as close
				# a match as possible for what would have been entered
				switch -regex ($parameter.value)
				{
				
					# If we have a multi value array put in then we need to break it out and add quotes as needed
					'[;]'	{ 
				
						# Reset the formatted value string
						$FormattedValue = $null
						
						# Split it into an array
						$valuearray = $switch.current.split(";")
						
						# For each entry in the array add quotes if needed and add it to the formatted value string
						$valuearray | ForEach-Object {
							if ($_ -match "[ \t]"){$FormattedValue = $FormattedValue + "`"" + $_ + "`";"}
							else {$FormattedValue = $FormattedValue + $_ + ";"}
						}
						
						# Clean up the trailing ;
						$FormattedValue = $FormattedValue.trimend(";")
						
						# Add our switch + cleaned up value to the command string
						$FullCommand = $FullCommand + " -" + $parameter.name + " " + $FormattedValue
					}
					
					# If we have a value with spaces add quotes
					'[ \t]'				{$FullCommand = $FullCommand + " -" + $parameter.name + " `"" + $switch.current + "`""}
					
					# If we have a true or false format them with :$ in front ( -allow:$true )
					'^True$|^False$'	{$FullCommand = $FullCommand + " -" + $parameter.name + ":`$" + $switch.current}
					
					# Otherwise just put the switch and the value
					default				{$FullCommand = $FullCommand + " -" + $parameter.name + " " + $switch.current}
				
				}
			}
			
			# Format our modified object
			if ([string]::IsNullOrEmpty($_.objectModified)){$ObjModified = ""}
			else { 
				$ObjModified = ($_.objectmodified.split("/"))[-1]
				$ObjModified = ($ObjModified.split("\"))[-1]
			}
			
			# Get just the name of the cmdlet that was run
			[string]$cmdlet = $_.CmdletName
			
			# Build the result object to return our values
			$Result | Add-Member -MemberType NoteProperty -Value $user -Name Caller
			$Result | Add-Member -MemberType NoteProperty -Value $cmdlet -Name Cmdlet
			$Result | Add-Member -MemberType NoteProperty -Value $FullCommand -Name FullCommand
			$Result | Add-Member -MemberType NoteProperty -Value $_.rundate -Name RunDate
			$Result | Add-Member -MemberType NoteProperty -Value $ObjModified -Name ObjectModified
			
			# Add the object to the array to be returned
			$ResultSet = $ResultSet + $Result
			
		}
	}

	# Final steps
	End {
		# Return the array set
		Return $ResultSet
	}
}

# Make sure we get back all of the unified audit log results for the search we are doing
Function Get-AllUnifiedAuditLogEntry
{
	param 
	(
	[Parameter(Mandatory=$true)]
	[string]$UnifiedSearch
	)
	
	# Validate the incoming search command
	if (($UnifiedSearch -match "-StartDate") -or ($UnifiedSearch -match "-EndDate") -or ($UnifiedSearch -match "-SessionCommand") -or ($UnifiedSearch -match "-ResultSize") -or ($UnifiedSearch -match "-SessionId"))
	{
		Out-LogFile "Do not include any of the following in the Search Command"
		Out-LogFile "-StartDate, -EndDate, -SessionCommand, -ResultSize, -SessionID"
		Write-Error -Message "Unable to process search command, switch in UnifiedSearch that is handled by this cmdlet specified" -ErrorAction Stop
	}
		
	# Make sure key variables are null
	[string]$cmd = $null
	
	# build our search command to execute
	$cmd = $UnifiedSearch + " -StartDate " + $Hawk.StartDate + " -EndDate " + $Hawk.EndDate + " -SessionCommand ReturnLargeSet -resultsize 1000 -sessionid " + (Get-Date -UFormat %H%M%S)
	Out-LogFile ("Running Unified Audit Log Search")
	Out-Logfile $cmd

	# Run the initial command
	[array]$Output=$null
	[array]$Output = (Invoke-Expression $cmd)

	# Make sure we got something back ... if not then we need to return here and abort
	if ($null -eq $output)
	{
		Out-LogFile ("[WARNING] - Unified Audit log returned no results for the search")
		Return $null
	}

	# Check to see if we have more results than returned
	If ($output[-1].Resultindex -lt $Output[-1].ResultCount)
	{
		# Change our command string to return the next page and not redo the search
		# $cmd = $cmd.Replace("-SessionCommand ReturnLargeSet","-SessionCommand ReturnNextPreviewPage")
		Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5," ") + " Total: " + $Output[-1].ResultCount)
	
		# Since we have more than 1k results we need to keep returning results until we have them all
		while ($Output[-1].Resultindex -lt $Output[-1].ResultCount)
		{
			# Out-LogFile $cmd
			[array]$Output += (Invoke-Expression $cmd)
			$Output = $Output | Sort-Object -Property ResultIndex
			Out-LogFile ("Retrieved:" + $Output[-1].ResultIndex.tostring().PadRight(5," ") + " Total: " + $Output[-1].ResultCount)
		}		
	}
		
	# Return our whole array
	return $Output
}

# Writes output to a log file with a time date stamp
Function Out-LogFile
{
	Param 
	( 
		[string]$string,
		[switch]$action,
		[switch]$notice,
		[switch]$silentnotice
	)
	
	# Make sure we have the Hawk Global Object
	Initialize-HawkGlobalObject
	$LogFile = Join-path $Hawk.FilePath "Hawk.log"
	$ScreenOutput = $true
	$LogOutput = $true
	
	# Get the current date
	[string]$date = Get-Date -Format G
		
	# Deal with each switch and what log string it should put out and if any special output

	# Action indicates that we are starting to do something
	if ($action)
	{
		[string]$logstring = ( "[" + $date + "] - [ACTION] - " + $string)

	}
	# If notice is true the we should write this to intersting.txt as well
	elseif ($notice)
	{
		[string]$logstring = ( "[" + $date + "] - ## INVESTIGATE ## - " + $string)

		# Build the file name for Investigate stuff log
		[string]$InvestigateFile = Join-Path (Split-Path $LogFile -Parent) "_Investigate.txt"
		$logstring | Out-File -FilePath $InvestigateFile -Append
	}
	# For silent we need to supress the screen output
	elseif ($silentnotice)
	{
		[string]$logstring = ( "Addtional Information: " + $string)
		# Build the file name for Investigate stuff log
		[string]$InvestigateFile = Join-Path (Split-Path $LogFile -Parent) "_Investigate.txt"
		$logstring | Out-File -FilePath $InvestigateFile -Append
		
		# Supress screen and normal log output
		$ScreenOutput = $false
		$LogOutput = $false

	}
	# Normal output
	else 
	{
		[string]$logstring = ( "[" + $date + "] - " + $string)
	}

	# Write everything to our log file
	if ($LogOutput)
	{
		$logstring | Out-File -FilePath $LogFile -Append
	}
	
	# Output to the screen
	if ($ScreenOutput)
	{
		Write-Information -MessageData $logstring -InformationAction Continue
	}

}

# Sends the output of a cmdlet to a txt file and a clixml file
Function Out-MultipleFileType
{
	param 
	(
		[Parameter (ValueFromPipeLine=$true)]
		$Object,
		[Parameter (Mandatory=$true)]
		[string]$FilePrefix,
		[string]$User,
		[switch]$Append=$false,
		[switch]$xml=$false,
		[Switch]$csv=$false,
		[Switch]$txt=$false,
		[Switch]$Notice

	)
	
	begin 
	{
		
		# If no file types were specified then we need to error out here
		if (($xml -eq $false) -and ($csv -eq $false) -and ($txt -eq $false))
		{
			Out-LogFile "[ERROR] - No output type specified on object"
			Write-Error -Message "No output type specified on object" -ErrorAction Stop
		}
		
		# Null out our array
		[array]$AllObject = $null
		
		# Set the output path
		if ($null -eq $user)
		{
			$Path = $Hawk.FilePath
		}
		else 
		{
			$path = join-path $Hawk.filepath $user
			# Test the path if it is there do nothing otherwise create it
			if (test-path $path){}
			else 
			{
				Out-LogFile ("Making output directory for user " + $Path)
				$Null = New-Item $Path -ItemType Directory
			}
		}
		
	}
	
	process 
	{
		# Collect up all of the incoming data into a single object for processing and output
		[array]$AllObject = $AllObject + $Object
		
	}
	
	end 
	{		
		if ($null -eq $AllObject)
		{
			Out-LogFile "No Data Found"
		}
		else 
		{
			
			# Determine what file type or types we need to write this object into and output it
			# Output XML File
			if ($xml -eq $true)
			{
				# lets put the xml files in a seperate directory to not clutter things up
				$xmlpath = Join-path $Path XML
				if (Test-path $xmlPath){}
				else 
				{
					Out-LogFile ("Making output directory for xml files " + $xmlPath)
					$null = New-Item $xmlPath -ItemType Directory
				}

				# Build the file name and write it out
				$filename = Join-Path $xmlPath ($FilePrefix + ".xml")
				Out-LogFile ("Writing Data to " + $filename)

				# Output our objects to clixml
				$AllObject | Export-Clixml $filename

				# If notice is set we need to write the file name to _Investigate.txt
				if ($Notice){Out-LogFile -string ($filename) -silentnotice}
			}
			
			# Output CSV file
			if ($csv -eq $true)
			{
				# Build the file name
				$filename = Join-Path $Path ($FilePrefix + ".csv")
				
				# If we have -append then append the data
				if ($append)
				{

					Out-LogFile ("Appending Data to " + $filename)
					
					# Write it out to csv making sture to append
					$AllObject | Export-Csv $filename -NoTypeInformation -Append
				}
				
				# Otherwise overwrite
				else 
				{
					Out-LogFile ("Writing Data to " + $filename)
					$AllObject | Export-Csv $filename -NoTypeInformation
				}

				# If notice is set we need to write the file name to _Investigate.txt
				if ($Notice){Out-LogFile -string ($filename) -silentnotice}
			}
			
			# Output Text files
			if ($txt -eq $true)
			{
				# Build the file name
				$filename = Join-Path $Path ($FilePrefix + ".txt")
				
				# If we have -append then append the data
				if ($Append)
				{
					Out-LogFile ("Appending Data to " + $filename)
					$AllObject | Format-List * | Out-File $filename -Append	
				}
				
				# Otherwise overwrite
				else 
				{
					Out-LogFile ("Writing Data to " + $filename)
					$AllObject | Format-List * | Out-File $filename
				}

				# If notice is set we need to write the file name to _Investigate.txt
				if ($Notice){Out-LogFile -string ($filename) -silentnotice}	
			}
		}
	}

}

# Returns a collection of unique objects filtered by a single property
Function Select-UniqueObject
{
	param
	(
		[Parameter(Mandatory=$true)]
		[array]$ObjectArray,
		[Parameter(Mandatory=$true)]
		[string]$Property
	)
	
	# Null out our output array
	[array]$Output = $null
	
	# Get the ID of the unique objects based ont he sort property
	[array]$UniqueObjectID = $ObjectArray | Select-Object -Unique -ExpandProperty $Property
	
	# Select the whole object based on the unique names found
	foreach ($Name in $UniqueObjectID)
	{
		[array]$Output = $Output + ($ObjectArray | Where-Object {$_.($Property) -eq $Name} | Select-Object -First 1)
	}
	
	return $Output

}

# Test if we are connected to the compliance center online and connect if now
Function Test-CCOConnection
{
	Write-Output "Not yet implemented"
}

# Test if we are connected to Exchange Online and connect if not
Function Test-EXOConnection
{
	try { $null = Get-OrganizationConfig -erroraction stop }
	catch [System.Management.Automation.CommandNotFoundException]
	{
		Out-LogFile "[ERROR] - Not Connected to Exchange Online"
		Write-Output "`nPlease connect to Exchange Online Prior to running"
		Write-Output "`nStandard connection method"
		Write-Output "https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx"
		Write-Output "`nFor Accounts protected by MFA"
		Write-Output "https://technet.microsoft.com/en-us/library/mt775114(v=exchg.160).aspx `n"
		break
	}
}

# Test if we are connected to MSOL and connect if we are not
Function Test-MSOLConnection
{
	
	try { Get-MsolCompanyInformation -ErrorAction Stop}
	catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException]
	{
		Out-LogFile "Please connect to MSOL prior to running this cmdlet"
		Out-LogFile "https://docs.microsoft.com/en-us/powershell/module/msonline/?view=azureadps-1.0#msonline `n"
		break
	}
}

# Test if we have a connection with the AzureAD Cmdlets
Function Test-AzureADConnection 
{
	$TestModule = Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue
	$MinimumVersion = New-Object -TypeName Version -ArgumentList "2.0.0.131"

	if ($null -eq $TestModule)
	{
		Out-LogFile "Please Install the AzureAD Module with the following command:"
		Out-LogFile "Install-Module AzureAD"
		break
	}
	# Since we are not null pull the highest version
	else 
	{
		$TestModuleVersion = ($TestModule | Sort-Object -Property Version -Descending)[0].version
	}
	
	# Test the version we need at least 2.0.0.131
	if ($TestModuleVersion -lt $MinimumVersion)
	{
		Out-LogFile ("AzureAD Module Installed Version: " + $TestModuleVersion)
		Out-LogFile ("Miniumum Required Version: " + $MinimumVersion)
		Out-LogFile "Please update the module with: Update-Module AzureAD"
		break
	}
	# Do nothing
	else {}

	try { $Null = Get-AzureADTenantDetail -ErrorAction Stop}
	catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
	{
		Out-LogFile "Please connect to AzureAD prior to running this cmdlet"
		Out-LogFile "Connect-AzureAD"
		break
	}
}

# Check to see if a recipient object was created since our start date
Function Test-RecipientAge
{
	Param([string]$RecipientID)
	
	$recipient = Get-Recipient -Identity $RecipientID -erroraction SilentlyContinue
	# Verify that we got something back
	if ($null -eq $recipient)
	{
		Return 2
	}
	# If the date created is newer than our StartDate return non zero (1)
	elseif ($recipient.whencreated -gt $Hawk.StartDate)
	{
		Return 1
	}
	# If it is older than the start date return 0
	else
	{
		Return 0
	}
	
}

# Determine if an IP listed in on the O365 XML list
Function Test-MicrosoftIP
{
	param
	(
	[Parameter(Mandatory=$true)]
	[string]$IPToTest
	)

	# Check if we have imported all of our IP Addresses
	if ($null -eq $MSFTIPList)
	{
		Out-Logfile "Building MSFTIPList"
		
		# Load our networking dll pulled from https://github.com/lduchosal/ipnetwork
		$dll = join-path (Split-path ((get-module Hawk).path) -Parent) "System.Net.IPNetwork.dll"

		
		$Error.Clear()
		Out-LogFile ("Loading Networking functions from " + $dll)
		[Reflection.Assembly]::LoadFile($dll)

		if ($Error.Count -gt 0)
		{
			Out-Logfile "[WARNING] - DLL Failed to load can't process IPs"
			Return "Unknown"
		}

		$Error.clear()
		# Read in the XML file from the internet
		Out-LogFile ("Reading XML for MSFT IP Addresses https://support.content.office.net/en-us/static/O365IPAddresses.xml")
		[xml]$msftxml = (Invoke-webRequest -Uri https://support.content.office.net/en-us/static/O365IPAddresses.xml).content

		if ($Error.Count -gt 0)
		{
			Out-Logfile "[WARNING] - Unable to retrieve XML file"
			Return "Unknown"
		}

		# Make sure our arrays are null
		[array]$ipv6 = $Null
		[array]$ipv4 = $Null

		# Go thru each product in the XML
		foreach ($Product in $msftxml.products.product)
		{
			
			# For each product look thru the list of ip addresses
			foreach ($addresslist in $Product.addresslist)
			{
				# If IPv6 add to that list
				if ($addresslist.type -eq "Ipv6")
				{
					$ipv6 += $addresslist.address

				}
				# if IPv4 add to that list
				elseif ($addresslist.type -eq "IPv4")
				{
					$ipv4 += $addresslist.address
				}
				# if anything else ignore
				else {}
			}
		}

		# Now we need to filter out the duplicate addresses in the lists
		$ipv6 = $ipv6 | select-object -Unique
		$ipv4 = $ipv4 | Select-Object -Unique

		Out-LogFile ("Found " + $ipv6.Count + " unique MSFT IPv6 address ranges")
		Out-LogFile ("Found " + $ipv4.count + " unique MSFT IPv4 address ranges")
		# New up using our networking dll we need to pull these all in as network objects
		foreach ($ip in $ipv6)
		{
			[array]$ipv6objects +=  [System.Net.IPNetwork]::Parse($ip)
		}
		foreach ($ip in $ipv4)
		{
			[array]$ipv4objects +=  [System.Net.IPNetwork]::Parse($ip)
		}

		# Now create our output object
		$output = $Null
		$output = New-Object -TypeName PSObject
		$output | Add-Member -MemberType NoteProperty -Value $ipv6objects -Name IPv6Objects
		$output | Add-Member -MemberType NoteProperty -Value $ipv4objects -Name IPv4Objects

		# Create a global variable to hold our IP list so we can keep using it
		Out-LogFile "Creating global variable `$MSFTIPList"
		New-Variable -Name MSFTIPList -Value $output -Scope global
	}
	
	# Determine if we have an ipv6 or ipv4 address
	if ($IPToTest -like "*:*")
	{

		# Compare to the IPv6 list
		[int]$i = 0
		[int]$count = $MSFTIPList.ipv6objects.count - 1
		# Compare each IP to the ip networks to see if it is in that network
		# If we get back a True or we are beyond the end of the list then stop
		do 
		{
			# Test the IP
			$parsedip = [System.Net.IPAddress]::Parse($IPToTest)
			$test = [System.Net.IPNetwork]::Contains($MSFTIPList.ipv6objects[$i],$parsedip)
			$i++
		}	
		until(($test -eq $true) -or ($i -gt $count))
		
		# Return the value of test true = in MSFT network
		Return $test
	}
	else 
	{
		# Compare to the IPv4 list
		[int]$i = 0
		[int]$count = $MSFTIPList.ipv4objects.count - 1
		
		# Compare each IP to the ip networks to see if it is in that network
		# If we get back a True or we are beyond the end of the list then stop
		do 
		{
			# Test the IP
			$parsedip = [System.Net.IPAddress]::Parse($IPToTest)
			$test = [System.Net.IPNetwork]::Contains($MSFTIPList.ipv4objects[$i],$parsedip)
			$i++
		}	
		until(($test -eq $true) -or ($i -gt $count))
				
		# Return the value of test true = in MSFT network
		Return $test
	}
}

# Determine if we have an array with UPNs or just a single UPN / UPN array unlabeled
Function Test-UserObject 
{
	param ([array]$ToTest)

	# See if we can get the UserPrincipalName property off of the input object
	# If we can't then we need to see if this is a UPN and convert it into an object for acceptable input
	if ($null -eq $ToTest[0].UserPrincipalName)
	{
		# Very basic check to see if this is a UPN
		if ($ToTest[0] -match '@')
		{
			[array]$Output = $ToTest | Select-Object -Property @{Name="UserPrincipalName";Expression={$_}}
			Return $Output
		}
		else 
		{
			Write-Log "[ERROR] - Unable to determine if input is a UserPrincipalName"
			Write-Log "Please provide a UPN or array of objects with propertly UserPrincipalName populated"
			Write-Error "Unable to determine if input is a User Principal Name" -ErrorAction Stop
		}
	}
	# If we can pull the value of UserPrincipalName then just return the same object back
	else
	{
		Return $ToTest
	}


}

# Hawk upgrade check
Function Update-HawkModule
{
	param 
	(
		[switch]$ElevatedUpdate
	)

	# If ElevatedUpdate is true then we are running from a forced elevation and we just need to run without prompting
	if ($ElevatedUpdate)
	{
		# Set upgrade to true
		$Upgrade = $true
	}
	else 
	{

		# See if we can do an upgrade check
		if ($null -eq (Get-Command Find-Module)){}
		
		# If we can then look for an updated version of the module
		else 
		{
			Write-Output "Checking for latest version online"
			$onlineversion = Find-Module -name Hawk -erroraction silentlycontinue
			$Localversion = (Get-Module Hawk | Sort-Object -Property Version -Descending)[0]
			
			if ($onlineversion.version -gt $localversion.version)
			{
				Write-Output "New version of Hawk module found online"
				Write-Output ("Local Version: " + $localversion.version + " Online Version: " + $onlineversion.version)
				
				# Prompt the user to upgrade or not
				$title = "Upgrade version"
				$message = "A Newer version of the Hawk Module has been found Online. `nUpgrade to latest version?"
				$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Stops the function and provides directions for upgrading."
				$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Continues running current function"
				$options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
				$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

				# Check to see what the user choose
				switch ($result)
				{
					0 {$Upgrade=$true}
					1 {$Upgrade=$false}
				}
			}
			# If the versions match then we don't need to upgrade
			else 
			{ 
				Write-Output "Latest Version Installed"
			}
		}
	}

	# If we determined that we want to do an upgrade make the needed checks and do it
	if ($Upgrade)
	{
		# Determine if we have an elevated powershell prompt
		If (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
		{
			# Update the module
			Write-Output "Downloading Updated Hawk Module"
			Update-Module Hawk -Force
			Write-Output "Update Finished"
			Sleep 3

			# If Elevated update then this prompt was created by the Update-HawkModule function and we can close it out otherwise leave it up
			if ($ElevatedUpdate){exit}
			
			# If we didn't elevate then we are running in the admin prompt and we need to import the new hawk module
			else 
			{
				Write-Output "Starting new PowerShell Window with the updated Hawk Module loaded"
				
				# We can't load a new copy of the same module from inside the module so we have to start a new window
				Start-Process powershell.exe -ArgumentList "-noexit -Command Import-Module Hawk -force" -Verb RunAs
				Write-Warning "Updated Hawk Module loaded in New PowerShell Window. `nPlease Close this Window."
				break		
			}

		}
		# If we are not running as admin we need to start an admin prompt
		else 
		{
			# Relaunch as an elevated process:
			Write-Output "Starting Elevated Prompt"
			Start-Process powershell.exe -ArgumentList "-noexit -Command Import-Module Hawk;Update-HawkModule -ElevatedUpdate" -Verb RunAs -Wait
						
			Write-Output "Starting new PowerShell Window with the updated Hawk Module loaded"
			
			# We can't load a new copy of the same module from inside the module so we have to start a new window
			Start-Process powershell.exe -ArgumentList "-noexit -Command Import-Module Hawk -force"
			Write-Warning "Updated Hawk Module loaded in New PowerShell Window. `nPlease Close this Window."
			break
		}
	}
	# Since upgrade is false we log and continue
	else 
	{
		Write-Output "Skipping Upgrade"
	}
}					

#endregion

# ============== Global Functions ==============
# Region Global Function

# Shows a basic "help" document on how to use Hawk
Function Show-HawkHelp
{
	Out-LogFile "Creating Hawk Help File"

	$help = "
BASIC USAGE INFORMATION FOR THE HAWK MODULE
===========================================
Hawk is in constant development.  We will be adding addtional data gathering and information analysis.


DISCLAIMER:
===========================================
THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT
PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY
OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT
LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS
AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR
ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE
FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS
PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)
ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,
EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES

PURPOSE:
===========================================
The Hawk module has been designed to ease the burden on O365 administrators who are performing 
a forensic analysis in their organization.

It does NOT take the place of a human reviewing the data generated and is simply here to make
data gathering easier.

HOW TO USE:
===========================================
Hawk is divided into two primary forms of cmdlets; user based Cmdlets and Tenant based cmdlets.

User based cmdlets take the form Verb-HawkUser<action>.  They all expect a -user switch and 
will retrieve information specific to the user that is specified.  Tenant based cmdlets take
the form Verb-HawkTenant<Action>.  They don't need any switches and will return information
about the whole tenant.

A good starting place is the Start-HawkTenantInvestigation this will run all the tenant based
cmdlets and provide a collection of data to start with.  Once this data has been reviewed
if there are specific user(s) that more information should be gathered on 
Start-HawkUserInvestigation will gather all the User specific information for a single user.

All Hawk cmdlets include help that provides an overview of the data they gather and a listing
of all possible output files.  Run Get-Help <cmdlet> -full to see the full help output for a 
given Hawk cmdlet.

Some of the Hawk cmdlets will flag results that should be further reviewed.  These will appear
in _Investigate files.  These are NOT indicative of unwanted activity but are simply things 
that should reviewed.

REVIEW HAWK CODE:
===========================================
The Hawk module is written in PowerShell and only uses cmdlets and function that are availble
to all O365 customers.  Since it is written in PowerShell anyone who has downloaded it can
and is encouraged to review the code so that they have a clear understanding of what it is doing
and are comfortable with it prior to running it in their environment.

To view the code in notepad run the following command in powershell:

	notepad (join-path ((get-module hawk -ListAvailable)[0]).modulebase 'Hawk.psm1')

To get the path for the module for use in other application run:
	((Get-module Hawk -listavailable)[0]).modulebase

	"

	$help | Out-MultipleFileType -FilePrefix "Hawk_Help" -txt

	Notepad (Join-Path $hawk.filepath "Hawk_Help.txt")

	<#
 
	.SYNOPSIS
	Creates the Hawk_Help.txt file

	.DESCRIPTION
	Create the Hawk_Help.txt file
	Opens the file in Notepad

	.OUTPUTS
	
	Hawk_Help.txt file

	.EXAMPLE
	Show-HawkHelp
	
	Creates the Hawk_Help.txt file and opens it in notepad
	
	#>


}

# Create the hawk global object for use by other cmdlets in the hawk module
Function Initialize-HawkGlobalObject 
{
	param 
	(
		[switch]$Force
	)

	# True if Doesn't exits; -force is true; variable is null
	if (($null -eq (Get-Variable -Name Hawk -ErrorAction SilentlyContinue)) -or ($Force -eq $true) -or ($null -eq $Hawk))
	{

		# Check to see if there is an Update for Hawk
		Update-HawkModule

		# If the global variable Hawk doesn't exist or we have -force then set the variable up
		Write-Output "Setting Up initial Hawk environment variable"
		
		# Check to see if the user has accepted the EULA
		# If they haven't prompt and ask to accept
		if ([string]::IsNullOrEmpty($Hawk.EULA))
		{
			Write-Output @(" 
			
	DISCLAIMER:

	THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT
	PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY
	OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT
	LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR
	PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS
	AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR
	ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE
	FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS
	PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)
	ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,
	EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES
			")

			# Prompt the user to agree with EULA
			$title = "Disclaimer"
			$message = "Do you agree with the above disclaimer?"
			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Logs agreement and continues use of the Hawk Functions."
			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Stops execution of Hawk Functions"
			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
			$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
			# If yes log and continue
			# If no log error and exit
			switch ($result)
			{
			    0 { 
					Write-Output "`n" 
					$Eula = ("Agreed " + (get-date))
				}
			    1 { 
					Write-Output "Aborting Cmdlet"					
					Write-Error -Message "Failure to agree with EULA" -ErrorAction Stop
					break
				}
			}
		}
		else {$Eula = $Hawk.EULA}

		# Null our object then create it
		$Output = $null
		$Output = New-Object -TypeName PSObject
		
		$ValidPath = $false
		While ($ValidPath -eq $false)
		{
		
			[string]$OutputPath = Read-Host "Please provide an output directory"			
			# Need to validate that the outputpath is a folder
			# Check if the path provided contains a file name
			if ((Split-Path $OutputPath -Leaf) -like "*.*")
			{
				Write-Output "Please provide the path to an existing directory and Not to a specific file name"
				continue
			}
			# Test if the path exists
			if (Test-Path $OutputPath)
			{
				# Verify that what we found is a container and not just a file with no extension
				if ((Get-Item $OutputPath).PSIsContainer -eq $true)
				{
					# Create our date_time subfolder
					[string]$FolderID = (get-date -UFormat %Y%m%d_%H%M).tostring()
					
					$FullOutputPath = Join-path $OutputPath $FolderID
					# Just in case we run this twice in a min lets not throw an error
					if (Test-Path $FullOutputPath)
					{
						Write-Output "Path Exists"
						$ValidPath = $true
					}
					# If it is not there make it
					else 
					{
						Write-Output ("Creating subfolder with name " + $FullOutputPath)
										
						$null = New-Item $FullOutputPath -ItemType Directory
					
						# Set validpath to true so we stop the loop
						$ValidPath = $true
					}
				}
				
				# If it exists but isn't a directory then throw an error
				else
				{
					Write-Output "Please provide a path to a directory"
				}
			}
			# If we can't find the path at all then the directory does exist
			else 
			{
				Write-Output "Please provide a path to a diretory that exists"
			}
		}
		
		# Add the path to our setting object
		$Output | Add-Member -MemberType NoteProperty -Name FilePath -Value $FullOutputPath
		
		# Get the number of days to look back 
		Do 
		{
			$Days = Read-Host "How far back in the past should we search? (1-90 Default 90)"
			
			# If nothing is entered default to 90
			if ([string]::IsNullOrEmpty($Days)){$Days="90"}
		}
		while
		(
			#Validate that we have a number between 1 and 90
			$Days -notmatch '\d{1,2}' -or (1..365) -notcontains $Days	
		)
		
		# Add the Days to look back and the calculated Start and End Dates based on that
		$Output | Add-Member -MemberType NoteProperty -Name DaysToLookBack -Value $Days
		$Output | Add-Member -MemberType NoteProperty -Name StartDate -Value (Get-date ((Get-Date).adddays(-([int]$Days))) -UFormat %m/%d/%Y)
		$Output | Add-Member -MemberType NoteProperty -Name EndDate -Value (Get-date ((Get-Date).adddays(1)) -UFormat %m/%d/%Y)
		$Output | Add-Member -MemberType NoteProperty -Name WhenCreated -Value (Get-Date -Format g)
		$Output | Add-Member -MemberType NoteProperty -Name EULA -Value $Eula
		
		# Create the global hawk variable
		Write-Output "Setting up Global Hawk environment variable`n"
		New-Variable -Name Hawk -Scope Global -value $Output -Force
		Out-LogFile "Global Variable Configured"
		Out-LogFile ("Version " + (Get-Module Hawk).version)
		Out-LogFile $Hawk
		
	}
	
	
	<#
 
	.SYNOPSIS
	Create global variable $Hawk for use by all Hawk cmdlets.

	.DESCRIPTION
	Creates the global variable $Hawk and populates it with information needed by the other Hawk cmdlets.
	
	* Checks for latest version of the Hawk module
	* Creates path for output files
	* Records target start and end dates for searches
	
	.PARAMETER Force
	Switch to force the function to run and allow the variable to be recreated

	.OUTPUTS
	Creates the $Hawk global variable and populates it with a custom PS object with the following properties

	Property Name	Contents
	==========		==========
	FilePath		Path to output files
	DaysToLookBack	Number of day back in time we are searching
	StartDate		Calculated start date for searches based on DaysToLookBack
	EndDate			One day in the future
	WhenCreated		Date and time that the variable was created
	EULA			If you have agreed to the EULA or not

	.EXAMPLE
	Initialize-HawkGlobalObject -Force
	
	This Command will force the creation of a new $Hawk variable even if one already exists.
	
	#>
	
}

# Compress all hawk data for upload
Function Compress-HawkData
{
	Out-LogFile ("Compressing all data in " + $Hawk.FilePath + " for Upload")
	# Make sure we don't already have a zip file
	if ($null -eq (Get-ChildItem *.zip -Path $Hawk.filepath)){}
	else 
	{
		Out-LogFile ("Removing existing zip file(s) from " + $Hawk.filepath)
		$allfiles = Get-ChildItem *.zip -Path $Hawk.FilePath
		# Remove the existing zip files
		foreach ($file in $allfiles)
		{
			$Error.Clear()
			Remove-Item $File.FullName -Confirm:$false -ErrorAction SilentlyContinue
			# Make sure we didn't throw an error when we tried to remove them
			if ($Error.Count -gt 0)
			{
				Out-LogFile "Unable to remove existing zip files from " + $Hawk.filepath + " please remove them manually"
				Write-Error -Message "Unable to remove existing zip files from " + $Hawk.filepath + " please remove them manually" -ErrorAction Stop
			}
			else {}
		}
	}


	
	# Get all of the files in the output directory
	#[array]$allfiles = Get-ChildItem -Path $Hawk.filepath -Recurse
	#Out-LogFile ("Found " + $allfiles.count + " files to add to zip")
	
	# create the zip file name
	[string]$zipname = "Hawk_" + (Split-path $Hawk.filepath -Leaf) + ".zip"
	[string]$zipfullpath = Join-Path $env:TEMP $zipname

	Out-LogFile ("Creating temporary zip file " + $zipfullpath)
	
	# Load the zip assembly
	Add-Type -Assembly System.IO.Compression.FileSystem

	# Create the zip file from the current hawk file directory
	[System.IO.Compression.ZipFile]::CreateFromDirectory($Hawk.filepath,$zipfullpath)
	
	# Move the item from the temp directory to the full filepath
	Out-LogFile ("Moving file to the " + $hawk.filepath + " directory")
	Move-Item $zipfullpath (Join-Path $Hawk.filepath $zipname)
	
	<#
 
	.SYNOPSIS
	Compresses all files located in the $Hawk.FilePath folder
	
	.DESCRIPTION
	Compresses all files located in the $Hawk.FilePath folder
	
	* Removes any zip files from the existing folder
	* Creates a zip file with name of Hawk_<folder name>
	* Adds all contents of the folder to the new zip file
	* Opens file explorer to the file path $Hawk.FilePath
	
	.OUTPUTS
	Zip file with all contents from $Hawk.FilePath

	.EXAMPLE
	Compress-HawkData
	
	Compressess all files and open explorer to the specified file path
	
	#>
	
}

#endregion

# ============== Tenant Centric Functions ==============
# Region Tenant Function

# Gathers basic tenant information and generates output
## TODO: Put in some analysis ... flag some key things that we know we should
# Auditing Off
# Dig thru transport rules and look for ones forwarding or turfing mail
Function Get-HawkTenantConfiguration 
{
	
	Test-EXOConnection
	
	#Check Audit Log Config Setting and make sure it is enabled
	Out-LogFile "Gathering Tenant Configuration Information" -action
	
	Out-LogFile "Admin Audit Log"
	Get-AdminAuditLogConfig | Out-MultipleFileType -FilePrefix "AdminAuditLogConfig" -txt -xml
	
	Out-LogFile "Organization Configuration"
	Get-OrganizationConfig| Out-MultipleFileType -FilePrefix "OrgConfig" -xml -txt
	
	Out-LogFile "Remote Domains"
	Get-RemoteDomain | Out-MultipleFileType -FilePrefix "RemoteDomain" -xml -csv
	
	Out-LogFile "Transport Rules"
	Get-TransportRule | Out-MultipleFileType -FilePrefix "TransportRules" -xml -csv
	
	Out-LogFile "Transport Configuration"
	Get-TransportConfig | Out-MultipleFileType -FilePrefix "TransportConfig" -xml -csv	
	
	<#
 
	.SYNOPSIS
	Gathers basic tenant information.

	.DESCRIPTION
	Gathers information about tenant wide settings
	* Admin Audit Log Configuration
	* Organization Configuration
	* Remote domains
	* Transport Rules
	* Transport Configuration
	
	.OUTPUTS
	File: AdminAuditLogConfig.txt
	Path: \
	Description: Output of Get-AdminAuditlogConfig

	File: AdminAuditLogConfig.xml
	Path: \XML
	Description: Output of Get-AdminAuditlogConfig as CLI XML

	File: OrgConfig.txt
	Path: \
	Description: Output of Get-OrganizationConfig

	File: OrgConfig.xml
	Path: \XML
	Description: Output of Get-OrganizationConfig as CLI XML

	File: RemoteDomain.txt
	Path: \
	Description: Output of Get-RemoteDomain

	File: RemoteDomain.xml
	Path: \XML
	Description: Output of Get-RemoteDomain as CLI XML

	File: TransportRules.txt
	Path: \
	Description: Output of Get-TransportRule

	File: TransportRules.xml
	Path: \XML
	Description: Output of Get-TransportRule as CLI XML

	File: TransportConfig.txt
	Path: \
	Description: Output of Get-TransportConfig

	File: TransportConfig.xml
	Path: \XML
	Description: Output of Get-TransportConfig as CLI XML
	
	#>
	
}

# Find any roles that have access to key edisocovery cmdlets and output the folks who have those rights
Function Get-HawkTenantEDiscoveryConfiguration 
{

	Test-EXOConnection

	Out-LogFile "Gathering Tenant information about E-Discovery Configuration" -action
	
	# Nulling our our role arrays
	[array]$Roles = $null
	[array]$RoleAssignements = $null
	
	# Look for E-Discovery Roles and who they might be assigned to
	$EDiscoveryCmdlets = "New-MailboxSearch","Search-Mailbox"
	
	# Find any roles that have these critical ediscovery cmdlets in them
	# Bad actors with sufficient rights could have created new roles so we search for them
	Foreach ($cmdlet in $EDiscoveryCmdlets)
	{
		[array]$Roles = $Roles + (Get-ManagementRoleEntry ("*\" + $cmdlet))
	}
	
	# Select just the unique entries based on role name
	$UniqueRoles = Select-UniqueObject -ObjectArray $Roles -Property Role
	
	Out-LogFile ("Found " + $UniqueRoles.count + " Roles with E-Discovery Rights")
	$UniqueRoles | Out-MultipleFileType -FilePrefix "EDiscoveryRoles" -csv -xml
	
	# Get everyone who is assigned one of these roles
	Foreach ($Role in $UniqueRoles)
	{
		[array]$RoleAssignements = $RoleAssignements + (Get-ManagementRoleAssignment -Role $Role.role -Delegating $false)
	}
	
	Out-LogFile ("Found " + $RoleAssignements.count + " Role Assignements for these Roles")
	$RoleAssignements | Out-MultipleFileType -FilePreFix "EDiscoveryRoleAssignments" -csv -xml

	<#
 
	.SYNOPSIS
	Looks for users that have e-discovery rights.

	.DESCRIPTION
	Searches for all roles that have e-discovery cmdlets.
	Searches for all users / groups that have access to those roles.	
		
	.OUTPUTS

	File: EDiscoveryRoles.csv
	Path: \
	Description: All roles that have access to the New-MailboxSearch and Search-Mailbox cmdlets

	File: EDiscoveryRoles.xml
	Path: \XML
	Description: All roles that have access to the New-MailboxSearch and Search-Mailbox cmdlets as CLI XML

	File: EDiscoveryRoleAssignments.csv
	Path: \
	Description: All users that are assigned one of the discovered roles

	File: EDiscoveryRoleAssignments.xml
	Path: \XML
	Description: All users that are assigned one of the discovered roles as CLI XML

	.EXAMPLE
	Get-HawkTenantEDiscoveryConfiguration 

	Runs the cmdlet against the current logged in tenant and outputs ediscovery information
	
	#>
	
}

# Search for any changes made to RBAC in the search window and report them
Function Get-HawkTenantRBACChanges
{

	Test-EXOConnection

	Out-LogFile "Gathering any changes to RBAC configuration" -action

	# Search EXO audit logs for any RBAC changes
	[array]$RBACChanges = Search-AdminAuditLog -Cmdlets New-ManagementRole,New-ManagementRoleAssignment,New-ManagementScope,Remove-ManagementRole,Remove-ManagementRoleAssignment,Set-MangementRoleAssignment,Remove-ManagementScope,Set-ManagementScope -StartDate $Hawk.StartDate -EndDate $Hawk.EndDate

	# If there are any results push them to an output file 
	if ($RBACChanges.Count -gt 0)
	{
		Out-LogFile ("Found " + $RBACChanges.Count + " Changes made to Roles Based Access Control")
		$RBACChanges | Get-SimpleAdminAuditLog | Out-MultipleFileType -FilePrefix "Simple_RBAC_Changes" -csv
		$RBACChanges | Out-MultipleFileType -FilePrefix "RBAC_Changes" -csv -xml
	}
	# Otherwise report no results found
	else
	{
		Out-LogFile "No RBAC Changes found."
	}
	


	<#
 
	.SYNOPSIS
	Looks for any changes made to Roles Based Access Control

	.DESCRIPTION
	Searches the EXO Audit logs for the following commands being run.
	New-ManagementRole
	Remove-ManagementRole
	New-ManagementRoleAssignment
	Remove-ManagementRoleAssignment
	Set-MangementRoleAssignment
	New-ManagementScope
	Remove-ManagementScope
	Set-ManagementScope	
		
	.OUTPUTS

	File: Simple_RBAC_Changes.csv
	Path: \
	Description: All RBAC cmdlets that were run in an easy to read format

	File: RBAC_Changes.csv
	Path: \
	Description: All RBAC changes in Raw format

	File: RBAC_Changes.xml
	Path: \XML
	Description: All RBAC changes as a CLI XML

	.EXAMPLE
	Get-HawkTenantRBACChanges

	Looks for all RBAC changes in the tenant within the search window
	
	#>



}

# RBAC Changes
# Changes to impersonation
Function Search-HawkTenantEXOAuditLog
{

	Test-EXOConnection

	#Region New-InboxRules
	Out-LogFile "Searching EXO Audit Logs" -Action 
	Out-LogFile ("Searching Entire Admin Audit Log for Specific cmdlets")
	Out-LogFile "Hunting for Inbox Rules Created in the Shell" -action
	[array]$TenantInboxRules = Search-AdminAuditLog -Cmdlets New-InboxRule -StartDate $Hawk.StartDate -EndDate $Hawk.EndDate
	
	# If we found anything report it and log it
	if ($TenantInboxRules.count -gt 0)
	{
	
		Out-LogFile ("Found " + $TenantInboxRules.count + " Inbox Rule(s) created from PowerShell")
		$TenantInboxRules | Get-SimpleAdminAuditLog | Out-MultipleFileType -fileprefix "Simple_New_InboxRule" -csv
		$TenantInboxRules | Out-MultipleFileType -fileprefix "New_InboxRules" -xml
	}
	
	# Running the search again instead of processing existing output in $tenantinboxrules, want the service to return this
	Out-LogFile "Hunting for Inbox Rules Created in the Shell" -action
	[array]$InvestigateInboxRules = Search-AdminAuditLog -StartDate $Hawk.StartDate -EndDate $Hawk.EndDate -cmdlets New-InboxRule -Parameters ForwardTo,ForwardAsAttachmentTo,RedirectTo,DeleteMessage
	
	# if we found a rule report it and output it to the _Investigate files
	if ($InvestigateInboxRules.count -gt 0)
	{
		Out-LogFile ("Found " + $InvestigateInboxRules.count + " Investigate rules") -notice
		$InvestigateInboxRules | Get-SimpleAdminAuditLog | Out-MultipleFileType -fileprefix "_Investigate_Simple_New_InboxRule" -csv -Notice
		$InvestigateInboxRules | Out-MultipleFileType -fileprefix "_Investigate_New_InboxRules" -xml -txt -Notice
	}
	
	#endregion
	
	#Region User Forwarding
	
	Out-LogFile "Hunting for user Forwarding Changes" -action
	[array]$TenantForwardingChanges = Search-AdminAuditLog -Cmdlets Set-Mailbox -Parameters ForwardingAddress,ForwardingSMTPAddress
	
	if ($TenantForwardingChanges.count -gt 0)
	{
		Out-LogFile ("Found " + $TenantForwardingChanges.count + " Change(s) to user Email Forwarding") -notice
		$TenantForwardingChanges | Get-SimpleAdminAuditLog | Out-MultipleFileType -FilePrefix "Simple_Forwarding_Changes" -csv -Notice
		$TenantForwardingChanges | Out-MultipleFileType -FilePrefix "Forwarding_Changes" -xml -Notice
		
		# Make sure our output array is null
		[array]$Output = $null
		
		# Checking if addresses were added or removed
		# If added compile a list
		Foreach ($Change in $TenantForwardingChanges)
		{

			# Get the user object modified
			$user = ($Change.CmdletParameters | Where-Object ($_.name -eq "Identity")).value

			# Check the ForwardingSMTPAddresses first
			if ([string]::IsNullOrEmpty(($Change.CmdletParameters | Where-Object {$_.name -eq "ForwardingSMTPAddress"}).value)){}
			# If not null then push the email address into $output
			else 
			{
				[array]$Output = $Output + ($Change.CmdletParameters | Where-Object {$_.name -eq "ForwardingSMTPAddress"}) | Select-Object -Property @{Name="UserModified";Expression={$user}},@{Name="TargetSMTPAddress";Expression={$_.value.split(":")[1]}}
			}
			
			# Check ForwardingAddress
			if ([string]::IsNullOrEmpty(($Change.CmdletParameters | Where-Object {$_.name -eq "ForwardingAddress"}).value)){}
			else 
			{
				# Here we get back a recipient object in EXO not an SMTP address
				# So we need to go track down the recipient object
				$recipient = Get-Recipient (($Change.CmdletParameters | Where-Object {$_.name -eq "ForwardingAddress"}).value) -ErrorAction SilentlyContinue
				
				# If we can't resolve the recipient we need to log that
				if ($null -eq $recipient)
				{
					Out-LogFile ("Unable to resolve forwarding Target Recipient " + ($Change.CmdletParameters | Where-Object {$_.name -eq "ForwardingAddress"})) -notice
				}
				# If we can resolve it then we need to push the address the mail was being set to into $output
				else 
				{
					# Determine the type of recipient and handle as needed to get out the SMTP address
					Switch ($recipient.RecipientType)
					{
						# For mailcontact we needed the external email address
						MailContact {[array]$Output += $recipient | Select-Object -Property @{Name="UserModified";Expression={$user}};@{Name="TargetSMTPAddress";Expression={$_.ExternalEmailAddress.split(":")[1] }}}
						# For all others I believe primary will work
						Default {[array]$Output += $recipient| Select-Object -Property @{Name="UserModified";Expression={$user}};@{Name="TargetSMTPAddress";Expression={$_.PrimarySmtpAddress}}}
					}
				}
			}					
		}
		
		# Output our email address user modified pairs
		Out-logfile ("Found " + $Output.count + " email addresses set to be forwarded mail") -notice
		$Output | Out-MultipleFileType -FilePrefix "Forwarding_Recipients" -csv -Notice

	}
	
	#endregion
	
	#Region Permission Changes
	Out-LogFile "Hunting for Mailbox Permissions Changes" -Action
	[array]$TenantMailboxPermissionChanges = Search-AdminAuditLog -StartDate $Hawk.StartDate -EndDate $Hawk.EndDate -cmdlets Add-MailboxPermission
	
	if ($TenantMailboxPermissionChanges.count -gt 0)
	{
		Out-LogFile ("Found " + $TenantMailboxPermissionChanges.count + " changes to mailbox permissions")
		$TenantMailboxPermissionChanges | Get-SimpleAdminAuditLog | Out-MultipleFileType -fileprefix "Simple_Mailbox_Permissions" -csv
		$TenantMailboxPermissionChanges | Out-MultipleFileType -fileprefix "Mailbox_Permissions" -xml

		## TODO: Possibly check who was added with permissions and see how old their accounts are		
	}
	
	## TODO: Hunt for mailbox folder permission changes
	## No sign of this being used / done so pushing this for now
	#endregion
	
	#Region Impersonation Access
	Out-LogFile "Hunting Impersonation Access" -action
	[array]$TenantImpersonatingRoles = Get-ManagementRoleEntry "*\Impersonate-ExchangeUser"
	if ($TenantImpersonatingRoles.count -gt 1)
	{
		Out-LogFile ("Found " + $TenantImpersonatingRoles.count + " Impersonation Roles.  Default is 1") -notice
		$TenantImpersonatingRoles | Out-MultipleFileType -fileprefix "_Investigate_Impersonation_Roles" -csv -xml -Notice
	}
	elseif ($TenantImpersonatingRoles.count -eq 0){}
	else 
	{
		$TenantImpersonatingRoles | Out-MultipleFileType -fileprefix "Impersonation_Roles" -csv -xml
	}
	
	$Output = $null
	# Search all impersonation roles for users that have access
	foreach ($Role in $TenantImpersonatingRoles)
	{
		[array]$Output += Get-ManagementRoleAssignment -Role $Role.role -GetEffectiveUsers -Delegating:$false}
	
	if ($Output.count -gt 1)
	{
		Out-LogFile ("Found " + $Output.cout + " Users/Groups with Impersonation rights.  Default is 1") -notice
		$Output | Out-MultipleFileType -fileprefix "_Investigate_Impersonation_Rights" -csv -xml -Notice
	}
	elseif ($Output.count -eq 1)
	{
		Out-LogFile ("Found default number of Impersonation users")
		$Output | Out-MultipleFileType -fileprefix "Impersonation_Rights" -csv -xml
	}
	else {}
		
	#endregion

		<#
 
	.SYNOPSIS
	Searches the admin audit logs for possible bad actor activities

	.DESCRIPTION
	Searches the Exchange admin audkit logs for a number of possible bad actor activies.
	
	* New inbox rules
	* Changes to user forwarding configurations
	* Changes to user mailbox permissions
	* Granting of impersonation rights
			
	.OUTPUTS

	File: Simple_New_InboxRule.csv
	Path: \
	Description: cmdlets to create any new inbox rules in a simple to read format
	
	File: New_InboxRules.xml
	Path: \XML
	Description: Search results for any new inbox rules in CLI XML format

	File: _Investigate_Simple_New_InboxRule.csv
	Path: \
	Description: cmdlets to create inbox rules that forward or delete email in a simple format

	File: _Investigate_New_InboxRules.xml
	Path: \XML
	Description: Search results for newly created inbox rules that forward or delete email in CLI XML
	
	File: _Investigate_New_InboxRules.txt
	Path: \
	Description: Search results of newly created inbox rules that forward or delete email

	File: Simple_Forwarding_Changes.csv
	Path: \
	Description: cmdlets that change forwarding settings in a simple to read format

	File: Forwarding_Changes.xml
	Path: \XML
	Description: Search results for cmdlets that change forwarding settings in CLI XML
	
	File: Forwarding_Recipients.csv
	Path: \
	Description: List of unique Email addresses that were setup to recieve email via forwarding

	File: Simple_Mailbox_Permissions.csv
	Path: \
	Description: Cmdlets that add permissions to users in a simple to read format

	File: Mailbox_Permissions.xml
	Path: \XML
	Description: Search results for cmdlets that change permissions in CLI XML

	File: _Investigate_Impersonation_Roles.csv
	Path: \
	Description: List all users with impersonation rights if we find more than the default of one

	File: _Investigate_Impersonation_Roles.csv
	Path: \XML
	Description: List all users with impersonation rights if we find more than the default of one as CLI XML

	File: Impersonation_Rights.csv
	Path: \
	Description: List all users with impersonation rights if we only find the default one

	File: Impersonation_Rights.csv
	Path: \XML
	Description: List all users with impersonation rights if we only find the default one as CLI XML
	
	.EXAMPLE
	Search-HawkTenantEXOAuditLog 

	Searches the tenant audit logs looking for changes that could have been made in the tenant.
	
	#>
	
}

# Executes the series of Hawk cmdets that search the whole tenant
Function Start-HawkTenantInvestigation
{

	Out-LogFile "Starting Tenant Sweep"
	
	Get-HawkTenantConfiguration
	Get-HawkTenantEDiscoveryConfiguration
	Search-HawkTenantEXOAuditLog
	Get-HawkTenantRBACChanges

	<#
 
	.SYNOPSIS
	Gathers common data about a tenant.

	.DESCRIPTION
	Runs all Hawk tenant related cmdlets and gathers the data.

	Cmdlet									Information Gathered
	-------------------------				-------------------------
	Get-HawkTenantConfigurationn			Basic Tenant information
	Get-HawkTenantEDiscoveryConfiguration	Looks for changes to ediscovery configuration
	Search-HawkTenantEXOAuditLog			Searches the EXO audit log for activity
	Get-HawkTenantRBACChanges				Looks for changes to Roles Based Access Control
	
	.OUTPUTS
	See help from individual cmdlets for output list.
	All outputs are placed in the $Hawk.FilePath directory

	.EXAMPLE
	Start-HawkTenantInvestigation

	Runs all of the tenant investigation cmdlets.
	
	#>
}

# Searches the unified audit log for logon activity by IP address
Function Search-HawkTenantActivityByIP
{
	param
	(
		[parameter(Mandatory=$true)]
		[string]$IpAddress
	)

	Test-EXOConnection

	# Replace an : in the IP address with . since : isn't allowed in a directory name
	$DirectoryName = $IpAddress.replace(":",".")

	# Make sure we got only a single IP address
	if ($IpAddress -like "*,*")
	{
		Out-LogFile "Please provide a single IP address to search."
		Write-Error -Message "Please provide a single IP address to search." -ErrorAction Stop
	}	

	Out-LogFile ("Searching for events related to " + $IpAddress) -action

	# Gather all of the events related to these IP addresses
	[array]$ipevents = Get-AllUnifiedAuditLogEntry -UnifiedSearch ("Search-UnifiedAuditLog -IPAddresses " + $IPAddress )

	# If we didn't get anything back log it
	if ($null -eq $ipevents)
	{
		Out-LogFile ("No IP logon events found for IP "	+ $IpAddress)
	}

	# If we did then process it
	else 
	{

		# Expand out the Data and convert from JSON
		[array]$ipeventsexpanded = $ipevents | Select-object -ExpandProperty AuditData | ConvertFrom-Json
		Out-LogFile ("Found " + $ipeventsexpanded.count + " related to provided IP" )
		$ipeventsexpanded | Out-MultipleFileType -FilePrefix "All_Events" -csv -xml -User $DirectoryName

		# Get the logon events that were a success
		[array]$successipevents = $ipeventsexpanded | where {$_.ResultStatus -eq "success"}
		Out-LogFile ("Found " + $successipevents.Count + " Successful logons related to provided IP")
		$successipevents | Out-MultipleFileType -FilePrefix "Success_Events" -csv -User $DirectoryName

		# Select all unique users accessed by this IP
		[array]$uniqueuserlogons = Select-UniqueObject -ObjectArray $ipeventsexpanded -Property "UserID"
		Out-LogFile ("IP " + $ipaddress + " has tried to access " + $uniqueuserlogons.count + " users") -notice
		$uniqueuserlogons | Out-MultipleFileType -FilePrefix "Unique_Users_Attempted" -csv -User $DirectoryName -Notice

		[array]$uniqueuserlogonssuccess = Select-UniqueObject -ObjectArray $successipevents -Property "UserID"
		Out-LogFile ("IP " + $IpAddress + " SUCCESSFULLY accessed " + $uniqueuserlogonssuccess.count + " users") -notice
		$uniqueuserlogonssuccess | Out-MultipleFileType -FilePrefix "Unique_Users_Success" -csv -xml -User $DirectoryName -Notice
	
	}	

	<#
 
	.SYNOPSIS
	Gathers logon activity based on a submitted IP Address.

	.DESCRIPTION
	Pulls logon activity from the Unified Audit log based on a provided IP address.
	Processes the data to highlight successful logons and the number of users accessed by a given IP address.

	.OUTPUTS
	
	File: All_Events.csv
	Path: \<IP>
	Description: All logon events

	File: All_Events.xml
	Path: \<IP>\xml
	Description: Client XML of all logon events

	File: Success_Events.csv
	Path: \<IP>
	Description: All logon events that were successful

	File: Unique_Users_Attempted.csv
	Path: \<IP>
	Description: List of Unique users that this IP tried to log into

	File: Unique_Users_Success.csv
	Path: \<IP>
	Description: Unique Users that this IP succesfully logged into

	File: Unique_Users_Success.xml
	Path: \<IP>\XML
	Description: Client XML of unique users the IP logged into

	
	.EXAMPLE

	Search-HawkTenantActivityByIP -IPAddress 10.234.20.12

	Searches for all Logon activity from IP 10.234.20.12.
	
	#>

}

# Uses start-robust cloud command to pull specific data from each user in the tenant
Function Get-HawkTenantInboxRules
{
	param ([string]$CSVPath)


	Test-EXOConnection

	# Prompt the user that this is going to take a long time to run
	$title = "Long Running Command"
	$message = "Running this search can take a very long time to complete (~1min per user). `nDo you wish to continue?"
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Continue operation"
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Exit Cmdlet"
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
	# If yes log and continue
	# If no log error and exit
	switch ($result)
	{
		0 { Out-LogFile "Starting full Tenant Search"}
		1 { Write-Error -Message "User Stopped Cmdlet" -ErrorAction Stop }
	}

	# Get the exo PS session
	$exopssession = get-pssession | where {($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.State -eq 'Opened')}

	# Gather all of the mailboxes
	Out-LogFile "Getting all Mailboxes"
	
	# If we don't have a value for csvpath then gather all users in the tenant
	if ([string]::IsNullOrEmpty($CSVPath))
	{
		$AllMailboxes = Invoke-Command -Session $exopssession -ScriptBlock {Get-Recipient -RecipientTypeDetails UserMailbox -ResultSize Unlimited |Select-Object -Property DisplayName,PrimarySMTPAddress,UserPrincipalName}
		$Allmailboxes | Out-MultipleFileType -FilePrefix "All_Mailboxes" -csv
	}
	# If we do read that in
	else 
	{
		# Import the csv with error checking
		$error.clear()
		$AllMailboxes = Import-Csv $CSVPath
		if ($error.Count -gt 0)
		{
			Write-Error "Problem importing csv file aborting" -ErrorAction Stop
		}
	}
	
	# Report how many mailboxes we are going to operate on
	Out-LogFile ("Found " + $AllMailboxes.count + " Mailboxes")

	# Get the path to start-robustcloudcommand
	$scriptpath = Join-Path (Split-path ((get-module Hawk).path) -Parent) "Start-RobustCloudCommand.ps1"
	
	# get EXO Credentials
	Out-LogFile "Gathering EXO Admin Credentials"
	$cred = Get-Credential -Message "EXO Credentials"

	# Path for robust log file
	$RobustLog = Join-path $Hawk.FilePath "Robust.log"

	# Build the command we are going to need to run with start-robustcloudcommand
	$cmd = $scriptpath + " -Agree -Credential `$cred -logfile `$RobustLog -recipients `$AllMailboxes -scriptblock {Get-HawkUserInboxRule -User `$input.primarySMTPAddress.tostring();Get-HawkUserEmailForwarding -user `$input.primarySMTPAddress.tostring()}"
	
	# Invoke our Start-Robust command to get all of the inbox rules
	Out-LogFile "===== Starting Robust Cloud Command to Gather User Specific information from all tenant users ====="
	Out-LogFile $cmd
	Invoke-Expression $cmd

	Out-LogFile "Process Complete"	

	<#
 
	.SYNOPSIS
	Gets inbox rules and forwarding directly from all mailboxes in the org.

	.DESCRIPTION
	Uses start-robustcloudcommand.ps1 to gather data from each mailbox in the org.
	Gathers inbox rules with Get-HawkUserInboxRule
	Gathers forwarding with Get-HawkUserEmailForwarding

	.PARAMETER CSVPath
	Path to a CSV file with a list of users to run against.
	CSV header should have DisplayName,PrimarySMTPAddress at minimum

	.OUTPUTS
	
	See Help for Get-HawkUserInboxRule for inbox rule output
	See Help for Get-HawkUserEmailForwarding for email forwarding output

	File: Robust.log
	Path: \
	Description: Logfile for Start-RobustCloudCommand.ps1
		
	.EXAMPLE
	Start-HawkTenantIndividualUserSearch
	
	Runs Get-HawkUserInboxRule and Get-HawkUserEmailForwarding against all mailboxes in the org

	.EXAMPLE
	Start-HawkTenantIndividualUserSearch -csvpath c:\temp\myusers.csv

	Runs Get-HawkUserInboxRule and Get-HawkUserEmailForwarding against all mailboxes listed in myusers.csv

	.LINK
	https://gallery.technet.microsoft.com/office/Start-RobustCloudCommand-69fb349e

	
	#>
}

# Retrives a list of all applciations that have the ability to access user data
# There are Azure AD Cmdlets for these
# https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/AzureAppEnumerationViaGraph.ps1
Function Get-HawkTenantOauthConsentGrants 
{
	Out-LogFile "Gathering Oauth Consent Grants"

	Test-AzureADConnection 

	# Next up gather the consent grants using the azureadcommand
	[array]$Grant = Get-AzureADOauth2PermissionGrant -all:$true

	# Check if we have a return
	if ($null -eq $Grant)
	{
		Out-LogFile "No Grants Found."
	}
	# If we do then we need to pull some addtional information then output
	else 
	{
		Out-LogFile ("Found " + $Grant.count + " OAuth Grants")
		Out-LogFile "Processing Grants"

		# Add in the display name information
		$FullGrantInfo = $Grant | Select-Object -Property *,@{Name="DisplayName";Expression={(Get-AzureADServicePrincipal -ObjectId $_.clientid).displayname}}

		# Push our data out to a file
		Out-MultipleFileType -Object $FullGrantInfo -FilePrefix AzureADOauthGrants -csv -txt

	}

	<#
 
	.SYNOPSIS
	Gathers application Oauth grants

	.DESCRIPTION
	Gathers Application Oauth grants along with their display names.  The grants listed are applications
	that have been granted access to various data inside the tenant.  The scope field outlines
	what data a given application has access to.

	.OUTPUTS
	File: AzureADOauthGrants.csv
	Path: \
	Description: Output of all grants as CSV.

	File: AzureADOauthGrants.txt
	Path: \
	Description: Output of all grants as txt
		
	.EXAMPLE
	Get-HawkTenantOauthConsentGrants
	
	Gathers all Oauth Grants

	#>

}

# Gets details about the applications that have access
# There are azure Ad cmdlets for these
# https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/AzureAppEnumerationViaGraph.ps1
Function Get-HawkTenantApplicationDetails {}

#endregion

# ============== User Centric Functions ==============
#Region User cmdlets

# Get the applications granted access to users
# https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/AzureAppEnumerationViaGraph.ps1
Function Get-HawkUserApplicationPrivileges {}

# Gets user inbox rules and looks for Investigate rules
Function Get-HawkUserInboxRule
{
	param
	(
	[Parameter(Mandatory=$true)]
	[array]$UserPrincipalName
	
	)
	
	Test-EXOConnection

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName
	
	foreach ($Object in $UserArray)
	{

		[string]$User = $Object.UserPrincipalName

		# Get Inbox rules
		Out-LogFile ("Gathering Inbox Rules: " + $User) -action
		$InboxRules = Get-InboxRule -mailbox  $User
	
		# If the rules contains one of a number of known suspecious properties flag them
		foreach ($Rule in $InboxRules)
		{
			# Set our flag to false
			$Investigate = $false
		
			# Evaluate each of the properties that we know bad actors like to use and flip the flag if needed
			if ($Rule.DeleteMessage -eq $true){ $Investigate = $true }
			if (!([string]::IsNullOrEmpty($Rule.ForwardAsAttachmentTo))){ $Investigate = $true}
			if (!([string]::IsNullOrEmpty($Rule.ForwardTo))){ $Investigate = $true}
			if (!([string]::IsNullOrEmpty($Rule.RedirectTo))){ $Investigate = $true}
		
			# If we have set the Investigate flag then report it and output it to a seperate file
			if ($Investigate -eq $true)
			{
				Out-LogFile ("Possible Investigate inbox rule found ID:" + $Rule.Identity + " Rule:" + $Rule.Name) -notice
				$Rule | Out-MultipleFileType -FilePreFix "_Investigate_InboxRules" -user $user -csv -append -Notice
			}	
		}
	
		# Output all of the inbox rules to a generic csv
		$InboxRules | Out-MultipleFileType -FilePreFix "InboxRules" -User $user -csv
	
		# Add all of the inbox rules to a generic collection file
		$InboxRules | Out-MultipleFileType -FilePrefix "All_InboxRules" -csv -Append
	}

	<#
 
	.SYNOPSIS
	Pulls inbox rules for the specified user.

	.DESCRIPTION
	Gathers inbox rules for the provided uers.
	Looks for rules that forward or delete email and flag them for follow up

	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS
	
	File: _Investigate_InboxRules.csv
	Path: \<User>
	Description: Inbox rules that delete or forward messages.

	File: InboxRules.csv
	Path: \<User>
	Description: All inbox rules that were found for the user.

	File: All_InboxRules.csv
	Path: \
	Description: All users inbox rules.
	
	.EXAMPLE

	Get-HawkUserInboxRule -UserPrincipalName user@contoso.com

	Pulls all inbox rules for user@contoso.com and looks for Investigate rules.

	.EXAMPLE

	Get-HawkUserInboxRule -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"})

	Gathers inbox rules for all users who have "C-Level" set in CustomAttribute1

	
	#>
}

# Looks to see if a single user has Email forwarding configured
Function Get-HawkUserEmailForwarding
{
	param
	(
	[Parameter(Mandatory=$true)]
	[array]$UserPrincipalName
	)

	Test-EXOConnection

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName

	foreach ($Object in $UserArray)
	{

		[string]$User = $Object.UserPrincipalName

		# Looking for email forwarding stored in AD
		Out-LogFile ("Gathering possible Forwarding changes for: " + $User) -action
		Out-LogFile "Collecting AD Forwarding Settings" -action
		$mbx = Get-Mailbox -identity $User
	
		# Check if forwarding is configured by user or admin	
		if ([string]::IsNullOrEmpty($mbx.ForwardingSMTPAddress) -and [string]::IsNullOrEmpty($mbx.ForwardingAddress))
		{
			Out-LogFile "No forwarding configuration found"
		}
		# If populated report it and add to a CSV file of positive finds
		else 
		{
			Out-LogFile ("Found Email forwarding User:" + $mbx.primarySMTPAddress + " ForwardingSMTPAddress:" + $mbx.ForwardingSMTPAddress + " ForwardingAddress:" + $mbx.ForwardingAddress) -notice
			$mbx | Select-Object DisplayName,UserPrincipalName,PrimarySMTPAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxAndForward,WhenChangedUTC | Out-MultipleFileType -FilePreFix "_Investigate_Users_WithForwarding" -append -csv -notice
		}	
	
		# Add all users searched to a generic output	
		$mbx | Select-Object DisplayName,UserPrincipalName,PrimarySMTPAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxAndForward,WhenChangedUTC | Out-MultipleFileType -FilePreFix "User_ForwardingReport" -append -csv
		# Also add to an output specific to this user
		$mbx | Select-Object DisplayName,UserPrincipalName,PrimarySMTPAddress,ForwardingSMTPAddress,ForwardingAddress,DeliverToMailboxAndForward,WhenChangedUTC | Out-MultipleFileType -FilePreFix "ForwardingReport" -user $user -csv

	}

		<#
 
	.SYNOPSIS
	Pulls mail forwarding for a specified user.

	.DESCRIPTION
	Pulls the values of ForwardingSMTPAddress and ForwardingAddress to see if the user has these configured.

	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS
	
	File: _Investigate_Users_WithForwarding.csv
	Path: \
	Description: All users that are found to have forwarding configured.

	File: User_ForwardingReport.csv
	Path: \
	Description: Mail forwarding configuration for all searched users; even if null.

	File: ForwardingReport.csv
	Path: \<user>
	Description: Forwarding confiruation of the searched user.
		
	.EXAMPLE

	Get-HawkUserEmailForwarding -UserPrincipalName user@contoso.com

	Gathers possible email forwarding configured on the user.

	.EXAMPLE

	Get-HawkUserEmailForwarding -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"})

	Gathers possible email forwarding configured for all users who have "C-Level" set in CustomAttribute1
	
	#>

}

# TODO: Filter out successful logons and report those seperate from full list
# With that possibily include a "expected region" to do more filtering?
# Maybe a seperate function for that?
Function Get-HawkUserAuthHistory
{
	param
	(
	[Parameter(Mandatory=$true)]
	[array]$UserPrincipalName,
	[switch]$ResolveIPLocations
	)
	
	Test-EXOConnection

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName

	foreach ($Object in $UserArray)
	{
		[string]$User = $Object.UserPrincipalName
	
		# Make sure our array is null
		[array]$UserLogonLogs = $null
	
		Out-LogFile ("Retrieving Logon History for " + $User) -action
	
		# Get back the account logon logs for the user
		$UserLogonLogs = Get-AllUnifiedAuditLogEntry -UnifiedSearch ("Search-UnifiedAuditLog -ObjectIds " + $User + " -RecordType AzureActiveDirectoryAccountLogon")
	
		# Expand out the AuditData and convert from JSON
		$ExpandedUserLogonLogs = $UserLogonLogs | Select-object -ExpandProperty AuditData | ConvertFrom-Json
	
		# Get only the unique IP addresses and report them
		[array]$LogonIPs = $ExpandedUserLogonLogs | Select-Object -Unique -Property ClientIP
		Out-LogFile ("Found " + $LogonIPs.count + " Unique IPs connecting to this user")
		$LogonIPs | Out-MultipleFileType -fileprefix "Logon_IPAddresses" -User $user -txt -csv

		# Make sure we have some logons before we process them
		if ($null -eq $LogonIPs)
		{
			Out-LogFile ("No logons found")
		}
		# If we do then process the logon objects further
		else 
		{
	
			# Set our Output array to null
			[array]$Output = $Null
	
			# if we have the resolve ip locations switch then we need to resolve the ip address to the location
			if ($ResolveIPLocations)
			{
	
				# Make sure our arrays are null
				[array]$IPLocations = $null
				$i = 0
				$EstimatedLookupTime = [int]($LogonIPs.Count * 1.5)
		
				# Loop thru each connection and get the location
				Foreach ($connection in $ExpandedUserLogonLogs)
				{
		
					Write-Progress -Activity "Looking Up Ip Address Locations" -CurrentOperation $connection.ClientIP -PercentComplete (($i/$ExpandedUserLogonLogs.count)*100) -Status ("Approximate Max Run time " + $EstimatedLookupTime + " seconds")
			
					# Get the location information for this IP address
					$Location = Get-IPGeolocation -ipaddress $connection.clientip

					# Add all of the locations for this user to an array of locations
					[array]$IPLocations += $Location
			
					# Combine the connection object and the location object so that we have a single output ready
					$Output += $connection | Select-Object -Property *,@{Name="CountryName";Expression={$Location.CountryName}},@{Name="RegionCode";Expression={$Location.RegionCode}},@{Name="RegionName";Expression={$Location.RegionName}},@{Name="City";Expression={$Location.City}},@{Name="ZipCode";Expression={$Location.ZipCode}},@{Name="KnownMicrosoftIP";Expression={$Location.KnownMicrosoftIP}}
			
					# increment our counter for the progress bar
					$i++

				}

				Write-Progress -Completed -Activity "Looking Up Ip Address Locations" -Status " "
		
				Out-LogFile "Writing Logon sessions with IP Locations"
				$Output | Out-MultipleFileType -fileprefix "Logon_Events_With_Locations" -User $User -csv -xml
				$Output | Where-Object {$_.LoginStatus -eq '0'} | Out-MultipleFileType -FilePrefix "Successful_Logon_Events_With_Locations" -User $User -csv -xml

				Out-LogFile "Writing List of unique logon locations"
				Select-UniqueObject -ObjectArray $IPLocations -Property ip | Out-MultipleFileType -fileprefix "Logon_Locations" -user $user -csv -txt
				$Global:IPlocationCache | Out-MultipleFileType -FilePrefix "All_Logon_Locations" -csv -txt
			}
	
			# if we don't have the lookup ip address switch then ouput just = our existing data
			else 
			{
				$Output = $ExpandedUserLogonLogs
				Out-LogFile "Writing Logon Session"
				$Output | Out-MultipleFileType -fileprefix "Logon_Events" -User $user -csv -xml
				$Output | Where-Object {$_.LoginStatus -eq '0'} | Out-MultipleFileType -FilePrefix "Successful_Logon_Events" -User $User -csv -xml
			}	
		}
	}
	
	<#
 
	.SYNOPSIS
	Gathers ip addresses that logged onto the user account

	.DESCRIPTION
	Pulls AzureActiveDirectoryAccountLogon events from the unified audit log for the provided user.
	
	If used with -ResolveIPLocations:
	Attempts to resolve the IP location using freegeoip.net
	Will flag ip addresses that are known to be owned by Microsoft using the XML from:
	https://support.office.com/en-us/article/URLs-and-IP-address-ranges-for-Office-365-operated-by-21Vianet-5C47C07D-F9B6-4B78-A329-BFDC1B6DA7A0

	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS
	
	File: Logon_IPAddresses.csv
	Path: \<User>
	Description: All unique logon IP addresses for this user.

	File: Logon_IPAddresses.txt
	Path: \<User>
	Description: All unique logon IP addresses for this user.

	File: Logon_IPAddresses.txt
	Path: \<User>
	Description: All unique logon IP addresses for this user.
	
	==== If -ResolveIPLocations is specified. ====
	
	File: Logon_Events_With_Locations.csv
	Path: \<User>
	Description: List of all logon events with the location discovered for the IP and if it is a Microsoft IP.

	File: Logon_Events_With_Locations.xml
	Path: \<User>\XML
	Description: List of all logon events with the location discovered for the IP and if it is a Microsoft IP in CLI XML.

	File: Successful_Logon_Events_With_Locations.csv
	Path: \<User>
	Description: List of all logon events that had LoginStatus = 0. Includes the location discovered for the IP and if it is a Microsoft IP.

	File: Successful_Logon_Events_With_Locations.xml
	Path: \<User>\XML
	Description: List of all logon events that had LoginStatus = 0. Includes the location discovered for the IP and if it is a Microsoft IP in CLI XML.

	File: All_Logon_Locations.csv
	Path: \
	Description: All ip addresses and their resolved locations for all users investigated.

	File: All_Logon_Locations.txt
	Path: \
	Description: All ip addresses and their resolved locations for all users investigated.

	==== If -ResolveIPLocations is NOT specified. ====

	File: Logon_Events.csv
	Path: \<User>
	Description:  All logon events that were found.

	File: Logon_Events.xml
	Path: \<User>\XML
	Description: All logon events that were found in CLI XML.

	File: Successful_Logon_Events.csv
	Path: \<User>
	Description:  All logon events that had LoginStatus = 0.

	File: Successful_Logon_Events.xml
	Path: \<User>\XML
	Description: All logon events that had LoginStatus = 0 in CLI XML.

	
	.EXAMPLE

	Get-HawkUserAuthHistory -UserPrincipalName user@contoso.com -ResolveIPLocations

	Gathers authenication information for user@contoso.com.
	Attempts to resolve the IP locations for all authetnication IPs found.

	.EXAMPLE

	Get-HawkUserAuthHistory -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"}) -ResolveIPLocations

	Gathers authenication information for all users that have "C-Level" set in CustomAttribute1
	Attempts to resolve the IP locations for all authetnication IPs found.
	
	#>	
}

# Get any unified audit logs related to mailbox auditing if enabled
function Get-HawkUserMailboxAuditing
{
	param
	(
	[Parameter(Mandatory=$true)]
	[array]$UserPrincipalName
	)

	Test-EXOConnection

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName

	foreach ($Object in $UserArray)
	{
		[string]$User = $Object.UserPrincipalName

		Out-LogFile ("Attempting to Gather Mailbox Audit logs " + $User) -action

		# Test if mailbox auditing is enabled
		$mbx = Get-Mailbox -identity $User
		if ($mbx.AuditEnabled -eq $true)
		{
			# if enabled pull the mailbox auditing from the unified audit logs
			Out-LogFile "Mailbox Auditing is enabled."
			Out-LogFile "Searching for Exchange related Audit Logs"
			$UserLogonLogs = Get-AllUnifiedAuditLogEntry -UnifiedSearch ("Search-UnifiedAuditLog -UserIDs " + $User + " -RecordType ExchangeItem")
		
			Out-LogFile ("Found " + $UserLogonLogs.Count + " Exchange audit records.")

			# Output the data we found
			$UserLogonLogs | Out-MultipleFileType -FilePrefix "Exchange_Audit" -User $User -xml -csv
			
		}
		# If auditing is not enabled log it and move on
		else
		{
			Out-LogFile ("Auditing not enabled for " + $User)
		}
	}

	<#
 
	.SYNOPSIS
	Gathers Mailbox Audit data if enabled for the user.

	.DESCRIPTION
	Check if mailbox auditing is enabled for the user.
	If it is pulls the mailbox audit logs fromt he time period specified for the investigation.

	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS
	
	File: Exchange_Audit.csv
	Path: \<User>
	Description: All exchange related audit events found.

	File: Exchange_Audit.xml
	Path: \<User>\xml
	Description: Client XML of all Exchange related audit events (Large file).
	
	.EXAMPLE

	Get-HawkUserMailboxAuditing -UserPrincipalName user@contoso.com

	Search for all Mailbox Audit logs from user@contoso.com

	.EXAMPLE

	Get-HawkUserMailboxAuditing -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"})

	Search for all Mailbox Audit logs for all users who have "C-Level" set in CustomAttribute1
	
	#>

}

# Gather basic information about a user for investigation
## TODO: Anything to flag here?  Folder stats ... folders that we don't normally see data in?
Function Get-HawkUserConfiguration
{
	param
	(
	[Parameter(Mandatory=$true)]
	[array]$UserPrincipalName
	)

	Test-EXOConnection

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName

	foreach ($Object in $UserArray)
	{
		[string]$User = $Object.UserPrincipalName

		Out-LogFile ("Gathering information about " + $User) -action

		#Gather mailbox information
		Out-LogFile "Gathering Mailbox Information"
		Get-Mailbox -identity $user | Out-MultipleFileType -FilePrefix "Mailbox_Info" -User $User -txt -xml
		Get-MailboxStatistics -identity $user | Out-MultipleFileType -FilePrefix "Mailbox_Statistics" -User $User -txt -xml
		Get-MailboxFolderStatistics -identity $user | Out-MultipleFileType -FilePrefix "Mailbox_Folder_Statistics" -User $User -txt -xml

		# Gather cas mailbox sessions
		Out-LogFile "Gathering CAS Mailbox Information"
		Get-CasMailbox -identity $user | Out-MultipleFileType -FilePrefix "CAS_Mailbox_Info" -User $User -txt -xml
	}

	<#
 
	.SYNOPSIS
	Gathers basic information about the provided user.

	.DESCRIPTION
	Gathers and records basic information about the provided user.
	
	* Get-Mailbox
	* Get-MailboxStatistics
	* Get-MailboxFolderStatistics
	* Get-CASMailbox
	
	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS

	File: Mailbox_Info.txt
	Path: \<User>
	Description: Output of Get-Mailbox for the user

	File: Mailbox_Info.xml
	Path: \<User>\XML
	Description: Client XML of Get-Mailbox cmdlet

	File: Mailbox_Statistics.txt
	Path : \<User>
	Description: Output of Get-MailboxStatistics for the user

	File: Mailbox_Statistics.xml
	Path : \<User>\XML
	Description: Client XML of Get-MailboxStatistics for the user

	File: Mailbox_Folder_Statistics.txt
	Path : \<User>
	Description: Output of Get-MailboxFolderStatistics for the user

	File: Mailbox_Folder_Statistics.xml
	Path : \<User>\XML
	Description: Client XML of Get-MailboxFolderStatistics for the user

	File: CAS_Mailbox_Info.txt
	Path : \<User>
	Description: Output of Get-CasMailbox for the user

	File: CAS_Mailbox_Info.xml
	Path : \<User>\XML
	Description: Client XML of Get-CasMailbox for the user

	.EXAMPLE
	Get-HawkUserConfiguration -user bsmith@contoso.com

	Gathers the user configuration for bsmith@contoso.com

	.EXAMPLE

	Get-HawkUserConfiguration -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"})

	Gathers the user configuration for all users who have "C-Level" set in CustomAttribute1

	
	#>
	
}

# String together the hawk user functions to pull data for a single user
Function Start-HawkUserInvestigation
{
	param
	(
		[Parameter(Mandatory=$true)]
		[array]$UserPrincipalName
	)

	Get-HawkTenantConfiguration

	# Verify our UPN input
	[array]$UserArray = Test-UserObject -ToTest $UserPrincipalName

	foreach ($Object in $UserArray)
	{
		[string]$User = $Object.UserPrincipalName
		
		Get-HawkUserConfiguration -User $User
		Get-HawkUserInboxRule -User $User
		Get-HawkUserEmailForwarding -User $User
		Get-HawkUserAuthHistory -User $user -ResolveIPLocations
		Get-HawkUserMailboxAuditing -User $User
	}

	<#
 
	.SYNOPSIS
	Gathers common data about a provided user.

	.DESCRIPTION
	Runs all Hawk users related cmdlets against the specified user and gathers the data.

	Cmdlet								Information Gathered
	-------------------------			-------------------------
	Get-HawkTenantConfigurationn        Basic Tenant information
	Get-HawkUserConfiguration           Basic User information
	Get-HawkUserInboxRule               Searches the user for Inbox Rules
	Get-HawkUserEmailForwarding         Looks for email forwarding configured on the user
	Get-HawkuserAuthHistory             Searches the unified audit log for users logons
	Get-HawkUserMailboxAuditing         Searches the unified audit log for mailbox auditing information			

	.PARAMETER UserPrincipalName
	Single UPN of a user, commans seperated list of UPNs, or array of objects that contain UPNs.

	.OUTPUTS
	See help from individual cmdlets for output list.
	All outputs are placed in the $Hawk.FilePath directory

	.EXAMPLE
	Start-HawkUserInvestigation -UserPrincipalName bsmith@contoso.com

	Runs all Get-HawkUser* cmdlets against the user with UPN bsmith@contoso.com

	.EXAMPLE

	Start-HawkUserInvestigation -UserPrincipalName (get-mailbox -Filter {Customattribute1 -eq "C-level"})

	Runs all Get-HawkUser* cmdlets against all users who have "C-Level" set in CustomAttribute1
	
	#>

}

#endregion


## TODO: Pull the Possible_Bad_Actors_Forwarding.csv file and do message tracking based on email addresses found

## TODO: Get All Audit logs related to a single user

## TODO: Figure out a way to determine if that bad actor has added rules via EWS/Outlook vs. cmdlets

## TODO: OWA changes to forwarding aren't logged in the audit log so I need to sweep the whole tenant to pull the forwarding information
## Get-mailbox ... should put this into a whole user data gathering

## TODO: RBAC Check against accounts ... list out unexpected roles

## TODO: Need the user inbox rule bit to spit out if no rules are found

## TODO:  Convert Get-HawkUserMailboxAuditing from search search unified audit log to -> search mailbox audit log

## TODO: Put in a cmdlet to change the date range ... should be obvious that you run this to do that

## TODO: Need Error Handling on the web lookups for ip -> location

## TODO: Needs to search up to 180 days ... but need to handle that somehow

## TODO: Add Start-HawkGUI to spawn basic gui that will launch Powershell with needed cmdlets

# TODO: Investigate MAPI Delivery Tables they should be null in default mailbox need to figure out how to pull them and make sure they are null

