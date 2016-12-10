[cmdletbinding(SupportsShouldProcess=$true, confirmimpact="none")] # Change confirmimpact to "high" if this function can break something
Param
(
	[Parameter(Mandatory=$false)]
	[string]$SectionMatchesConfigPath="C:\Temp\practigrab\sectionMatches.json",
	
	[Parameter(Mandatory=$false)]
	[string[]]$AdditionalMatchFiles=@("C:\Temp\practigrab\SheltonChampionshipSeries.csv"),
	
	[Parameter(Mandatory=$false)]
	[switch]$PassThruRaw,
	
	[Parameter(Mandatory=$false)]
	[switch]$PassThruFinal
)

function getHTML ($uri, $timeoutNum=5) {
	$count = 0
	$success = $false
	Write-Debug "Function: getHTML; Input: uri; Value: *$uri*"
	while (($count -lt $timeoutNum-1) -and (!$success)) {
		try {
			$html = Invoke-WebRequest -Uri $uri
			$success = $true
		}
		catch {
			$count++
			if ($count -eq $timeoutNum-1) {
				Write-Host "Failed to get HTML. Exiting script..."
				Exit 1
			}
			Start-Sleep 2
		}
	}
	
	# Return html object.
	return $html
}


function Get-MatchDef ([string]$matchID)
{
	$uri = "https://s3.amazonaws.com/ps-scores/production/$matchID/match_def.json"
	#Write-Host "Match Def URI: $uri"
	
	$html = getHTML $uri
	$json = $html | ConvertFrom-Json
	return $json
}

function Get-MatchResults ([string]$matchID)
{
	$uri = "https://s3.amazonaws.com/ps-scores/production/$matchID/results.json"
	#Write-Host "Match Results URI: $uri"
	
	$html = getHTML $uri
	$json = $html | ConvertFrom-Json
	return $json
}

function Get-OverallByDivisionPercent
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$false)]
		$sectionShooters,
		
		[Parameter(Mandatory=$false)]
		$matchInfo
	)
	
	$matchShooters = @()
	<#
	foreach ($shooter in $sectionShooters)
	{
		$matchShooter = $matchInfo.matchShooters | Where sh_id -eq $shooter.USPSANumber
		if ($true)
		{
			$division = $matchShooter.sh_dvp
			$class = $matchShooter.sh_grd
			$shooterUUID = $matchShooter.sh_uuid
			$divPercent = ($matchInfo.matchResults.match."$division" | Where shooter -eq $shooterUUID ).matchPercent
			
			$matchShooters += [pscustomobject]@{
				USPSANumber = $shooter.USPSANumber.ToUpper()
				FirstName = $shooter.FirstName
				LastName = $shooter.LastName
				MatchName = $matchInfo.matchName
				Club = $matchInfo.Club
				Division = $division
				Class = $class
				DivisionPercent = $divPercent
				}
		}
	}
	
	#>
	
	foreach ($shooter in $matchInfo.matchShooters)
	{
		$uspsaNumber = $shooter.sh_id
		$firstName = $shooter.sh_fn
		$lastName = $shooter.sh_ln
		$division = $shooter.sh_dvp
		$class = $shooter.sh_grd
		$shooterUUID = $shooter.sh_uuid
		$divPercent = ($matchInfo.matchResults.match."$division" | Where shooter -eq $shooterUUID ).matchPercent
		
		# Exceptions
		
		if ($lastName -eq "Hong" -and $firstName -eq "Andrew")
		{
			$uspsaNumber = "A83199"
		}
		
		if ($class -eq "GM")
		{
			$class = "G"
		}
		
		if ($uspsaNumber -eq "L2124" -and $division -eq "Limited 10")
		{
			$class = "B"
		}
		
		if ($lastName -eq "Fenlin" -and $firstName -eq "Jim")
		{
			$uspsaNumber = "TY77726"
		}
		
		if ($lastName -eq "Cook" -and $firstName -eq "Jason")
		{
			$uspsaNumber = "A85741"
		}
		
		if ($lastName -eq "Paolini" -and $firstName -eq "Austin")
		{
			$uspsaNumber = "A85741"
		}
		
		$matchShooters += [pscustomobject]@{
			USPSANumber = $uspsaNumber.ToUpper()
			FirstName = $firstName
			LastName = $lastName
			MatchName = $matchInfo.matchName
			Club = $matchInfo.Club
			ClubOrdered = $matchInfo.ClubOrdered
			Division = $division
			Class = $class
			DivisionPercent = [single]$divPercent
			}
	}
	
	return $matchShooters
}


function Get-StandingsRaw
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$sectionShooters,
		
		[Parameter(Mandatory=$true)]
		$sectionMatch,
		
		[Parameter(Mandatory=$false)]
		$excelPath
	)

	$matchDefJson = Get-MatchDef $sectionMatch.PractiScoreID
	$matchResultsJson = Get-MatchResults $sectionMatch.PractiScoreID


	$matchName = $matchDefJson.match_name
	$matchShooters = $matchDefJson.match_shooters
	$matchStages = $matchDefJson.match_stages

	$matchInfo += [pscustomobject]@{
				matchName = $matchName
				Club = $sectionMatch.Club
				ClubOrdered = "$($sectionMatch.MatchNumber) - $($sectionMatch.Club)"
				matchShooters = $matchShooters
				matchStages = $matchStages
				matchResults = $matchResultsJson
				}
	$matchOverallByDivision = Get-OverallByDivisionPercent -sectionShooters $sectionShooters -matchInfo $matchInfo
	$matchOverallByDivision | Export-CSV "C:\temp\practigrab\$($sectionMatch.Club)-ovrbydiv.csv" -NoTypeInformation
	if ($excelPath)
	{
		$matchOverallByDivision | Export-Excel -Path $excelPath -WorkSheetname $sectionMatch.Club -FreezeTopRow -AutoSize
	}


	return $matchOverallByDivision

}

function Build-MasterSheet
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$standingsRaw
	)
	
	$clubs = $standingsRaw | Select Club -Unique | Sort
	
	$sectionShooterResult = [pscustomobject]@{
			USPSANumber = ""
			FirstName = ""
			LastName = ""
			Division = ""
			Class = ""
			SectionScore = ""
			ScoresUsed = ""
			OverallAward = ""
			ClassAward = ""
			}
	
	foreach ($club in $clubs)
	{
		$sectionShooterResult | Add-Member -MemberType NoteProperty -Name $club.club -Value ""
	}
	
	return $sectionShooterResult
}


function Process-Standings
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$standingsRaw,
		
		[Parameter(Mandatory=$false)]
		[int]$BestXOf=3
	)
	
	$standingsRaw | foreach-object {$_.DivisionPercent = [single]$_.DivisionPercent}
	
	$uspsaNumbers = ($standingsRaw | Where {($_.DivisionPercent -ne 0) -and ($_.USPSAnumber -ne "") -and ($_.USPSAnumber -ne "PEN")} | Select USPSAnumber -Unique | Sort).USPSANumber
	$shooterStandingObj = Build-MasterSheet -standingsRaw $standingsRaw
	$finalStandings = @()

	foreach ($uspsaNumber in $uspsaNumbers)
	{
		# This year shooters may have enough scores for multiple divisions. Make sure we separate out divisions.
		$shooterDivs = ($standingsRaw | Where {($_.USPSANumber -eq $uspsaNumber) -and ($_.DivisionPercent -ne 0)} | Select Division -Unique).Division
		
		foreach ($division in $shooterDivs)
		{
			#Write-Host "Calculating average scores for shooter, $uspsaNumber"
			$shooterStanding = $shooterStandingObj.PsObject.Copy()
			$shooterResults = $standingsRaw | Where {($_.USPSANumber -eq $uspsaNumber) -and ($_.DivisionPercent -ne 0) -and $_.Division -eq $division}
			$bestOfResults = @()
			$bestOfResults += $shooterResults | Sort DivisionPercent -Descending | Select -First $BestXOf
			
			if ($bestOfResults.length -lt $BestXOf)
			{
				[single]$average = $null
				$shooterStanding.ScoresUsed = "Not eligible for series score. $($bestOfResults.length) out of $BestXOf required matches."
			}
			else
			{
				$averageObj = $bestOfResults.DivisionPercent | Measure-Object -Average
				[single]$average = [single]([math]::Round($averageObj.Average,2))
				$shooterStanding.ScoresUsed = $bestOfResults.Club -join ';'
			}
			#Write-Host "Average is, $average"
			
			
			foreach ($shooterResult in $shooterResults)
			{ 
				$shooterStanding."$($shooterResult.Club)" = $shooterResult.DivisionPercent
			}
			
			$shooterStanding.USPSANumber = $uspsaNumber
			$shooterStanding.FirstName = $shooterResults[0].FirstName
			$shooterStanding.LastName = $shooterResults[0].LastName
			$shooterStanding.Division = $division
			$shooterStanding.Class = $shooterResults[0].Class
			$shooterStanding.SectionScore = $average
			
			
			$finalStandings += $shooterStanding
		}
	}
	
	$finalStandings
}

function Calculate-OverallByDivisionPercent
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings
	)
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:standingByDivisionHtml
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:standingByDivisionHtml -innerHtml "Overall Results By Division"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByDivisionHtml -innerHtml "* indicates the shooter qualified for a division or class award. Refer to the awards section for details."
	#Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByDivisionHtml -innerHtml "Overall Results By Division"
	
	foreach ($division in $global:divisions)
	{
		Write-Host $division
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:standingByDivisionHtml
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:standingByDivisionHtml -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:standingByDivisionHtml
		Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByDivisionHtml
		
		$shooters = @()
		$shooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0)} | Sort SectionScore -Descending
		if ($shooters -ne $null)
		{
			$numEligibleShooters = $shooters.Length
			#Write-Host "$numEligibleShooters eligible shooters"
			Write-Host
			
			
			
			$place = 1
			foreach ($shooter in $shooters)
			{
				$firstName = $shooter.FirstName
				$lastName = $shooter.LastName
				$uspsaNumber = $shooter.USPSANumber
				$sectionScore = $shooter.SectionScore
				
				$awardNotation = ""
				if ($shooter.OverallAward -ne "")
				{
					$overallPlace = $shooter.OverallAward.Substring(0,1)
					#$awardNotation = "&nbsp&nbsp&nbsp&nbsp*O$overallPlace"
					$awardNotation = "&nbsp&nbsp*"
				}
				if ($shooter.ClassAward -ne "")
				{
					$classPlace = $shooter.ClassAward.Substring(0,1)
					#$awardNotation = "&nbsp&nbsp&nbsp&nbsp*C$classPlace"
					$awardNotation = "&nbsp&nbsp*"
				}
				
				$placeFull = Get-PlaceFull -place ([string]($place))
				$shooterOutput = "$placeFull Place - $firstName $lastName ($uspsaNumber) - $sectionScore%$($awardNotation)"
				Write-Host $shooterOutput
				Generate-Html -elementType "html" -htmlOutputPath $global:standingByDivisionHtml -innerHtml $shooterOutput
				Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtml
				$place++
			}
			
		}
		else
		{
			Write-Host "No eligible shooters."
			Generate-Html -elementType "html" -htmlOutputPath $global:standingByDivisionHtml -innerHtml "No eligible shooters."
			Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtml
		}
		
		Write-Host
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByDivisionHtml
		Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByDivisionHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByDivisionHtml
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:standingByDivisionHtml
}

function Calculate-ClassByDivisionPercent
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings
	)
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:standingByClassHtml
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:standingByClassHtml -innerHtml "Class Results By Division"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByClassHtml -innerHtml "* indicates the shooter qualified for a division or class award. Refer to the awards section for details."
	
	foreach ($division in $global:divisions)
	{
		Write-Host $division
		Write-Host
		
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:standingByClassHtml
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:standingByClassHtml -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:standingByClassHtml
		Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByClassHtml
		
		foreach ($class in $global:classes)
		{
			$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			Write-Host $fullName
			
			Generate-Html -elementType "divClass" -htmlOutputPath $global:standingByClassHtml
			Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:standingByClassHtml -innerHtml $fullName
			Generate-Html -elementType "divClassBody" -htmlOutputPath $global:standingByClassHtml
			Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByClassHtml
			
			
			$uniqueShooters = @()
			$uniqueShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $class)} | Select USPSANumber -Unique
			$numUniqueShooters = $uniqueShooters.Count
			
			$eligibleShooters = @()
			$eligibleShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.Class -eq $class)} | Sort SectionScore -Descending
			if ($eligibleShooters -ne $null)
			{
				$numEligibleShooters = $eligibleShooters.Length
				Write-Host "$numEligibleShooters eligible shooters"
				Write-Host
				
				$place = 1
				foreach ($eligibleShooter in $eligibleShooters)
				{
					$firstName = $eligibleShooter.FirstName
					$lastName = $eligibleShooter.LastName
					$uspsaNumber = $eligibleShooter.USPSANumber
					$sectionScore = $eligibleShooter.SectionScore
					
					$awardNotation = ""
					if ($eligibleShooter.OverallAward -ne "")
					{
						$overallPlace = $eligibleShooter.OverallAward.Substring(0,1)
						#$awardNotation = "&nbsp&nbsp&nbsp&nbsp*O$overallPlace"
						$awardNotation = "&nbsp&nbsp*"
					}
					if ($eligibleShooter.ClassAward -ne "")
					{
						$classPlace = $eligibleShooter.ClassAward.Substring(0,1)
						#$awardNotation = "&nbsp&nbsp&nbsp&nbsp*C$classPlace"
						$awardNotation = "&nbsp&nbsp*"
					}
					
					$placeFull = Get-PlaceFull -place ([string]($place))
					$shooterOutput = "$placeFull Place - $firstName $lastName ($uspsaNumber) - $sectionScore%$($awardNotation)"
					Write-Host $shooterOutput
					Generate-Html -elementType "html" -htmlOutputPath $global:standingByClassHtml -innerHtml $shooterOutput
					Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtml
					$place++
				}
			}
			else
			{
				Write-Host "No eligible shooters."
				Generate-Html -elementType "html" -htmlOutputPath $global:standingByClassHtml -innerHtml "No eligible shooters."
				Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtml
			}
			
			Write-Host
			Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByClassHtml
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtml
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtml
		}
		
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByClassHtml
		Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtml
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:standingByClassHtml
}

function Calculate-SectionStats
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings,
		
		[Parameter(Mandatory=$true)]
		$rawStandings
	)
	
	$sectionStats = @()
	
	foreach ($division in $global:divisions)
	{		
		#Write-Host $division
		
		$sectionShooterResult = [pscustomobject]@{
			Division = $division
			Class = "Overall"
			TotalUniqueShooters = @($rawStandings | Where {($_.Division -eq $division) -and ($_.USPSANumber -ne "") -and ($_.USPSANumber -ne "PEN")} | Select USPSANumber -Unique).Count
			TotalEligibleShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0)}).Count
		}
		$sectionStats += $sectionShooterResult
		
		foreach ($class in $global:classes)
		{
			$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			#Write-Host $fullName
			
			$uniqueShooters = @()
			$uniqueShooters += $rawStandings | Where {($_.Division -eq $division) -and ($_.USPSANumber -ne "") -and ($_.USPSANumber -ne "PEN") -and ($_.Class -eq $class)} | Select USPSANumber -Unique
			$numUniqueShooters = $uniqueShooters.Count
			
			$eligibleShooters = @()
			$eligibleShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.Class -eq $class)} | Sort SectionScore -Descending
			$numEligibleShooters = $eligibleShooters.Length
			
			$sectionShooterResult = [pscustomobject]@{
				Division = $division
				Class = $fullName
				TotalUniqueShooters = $numUniqueShooters
				TotalEligibleShooters = $numEligibleShooters
			}
			
			$sectionStats += $sectionShooterResult
		}
	}
	
	$sectionStats
	
}

function Calculate-OverallAwards
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings,
		
		[Parameter(Mandatory=$true)]
		$sectionStats
	)
	
	#Write-Host "Overall Awards Calc"
	foreach ($division in $global:divisions)
	{
		#Write-Host "Division: $division"
		$numberUniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalUniqueShooters
		$numberEligibleShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalEligibleShooters
		
		if ($numberUniqueShooters -ge $global:overallMin)
		{
			#Write-Host "The number of shooters in this division ($($numberUniqueShooters)) met the minimum required shoooters ($($global:overallMin))."
			$shooters = @()
			$shooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0)} | Sort SectionScore -Descending
			if ($shooters -ne $null)
			{
				$numShooters = $shooters.Length
				if ($numShooters -lt $global:overallPlaceLimit) { $placeLimit = $numShooters }
				else { $placeLimit = $global:overallPlaceLimit }
				
				if ($placeLimit -gt $numberEligibleShooters)
				{
					$placeLimit = $numberEligibleShooters
				}
				
				#Write-Host "numShooters: $numShooters"
				#Write-Host "PlaceLimit: $placeLimit"
				
				for ($i = 0; $i -lt $placeLimit; $i++)
				{
					#Write-Host "Working on place $($i + 1) of $placeLimit"
					$uspsaNumber = $shooters[$i].USPSANumber
					#Write-Host "placed $uspsaNumber"
					$place = Get-PlaceFull -place ([string]($i+1))
					($finalStandings | Where {($_.USPSANumber -eq $uspsaNumber) -and $_.Division -eq $division}).OverallAward = "$place Place $division Overall"
				}
			}
		}
		else
		{
			#Write-Host "The number of shooters in this division ($($numberUniqueShooters)) did not meet the minimum required shoooters ($($global:overallMin))."
		}
	}
	
}

function Calculate-ClassAwards
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings,
		
		[Parameter(Mandatory=$true)]
		$sectionStats
	)
	
	#Write-Host "Class Awards Calc"
	foreach ($division in $global:divisions)
	{
		#Write-Host "Division: $division"
		
		foreach ($class in $global:classes)
		{
			$classFullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			#Write-Host "Class: $classFullName"
			
			$numberUniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $classFullName)}).TotalUniqueShooters
			$numberEligibleShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $classFullName)}).TotalEligibleShooters

			
			if ($numberUniqueShooters -ge $global:classMinFirst)
			{
				#Write-Host "The number of shooters in this division ($($numberUniqueShooters)) met the minimum required shoooters ($($global:classMinFirst))."
				
				# Determine the number of places with the following formula. First place awarded after minimum 5 shooters.
				# 1 place for every 3 shooters after that till we reach the maximum number of places. This max is configurable.
				$placeLimit = ([Math]::floor(($numberUniqueShooters - $global:classMinFirst) / $global:ClassInterval)) + 1	
				if ($placeLimit -gt $global:classPlaceLimit)
				{
					$placeLimit = $global:classPlaceLimit
				}
				
				if ($placeLimit -gt $numberEligibleShooters)
				{
					$placeLimit = $numberEligibleShooters
				}
				
				$shooters = @()
				$shooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $classFullName) -and ($_.SectionScore -gt 0) -and ($_.OverallAward -eq "")} | Sort SectionScore -Descending
				
				if ($placeLimit -gt $shooters.Count)
				{
					$placeLimit = $shooters.Count
				}
				
				if ($shooters -ne $null)
				{
					$numShooters = $shooters.Length
					
					#Write-Host "numShooters: $numShooters"
					#Write-Host "PlaceLimit: $placeLimit"
					
					for ($i = 0; $i -lt $placeLimit; $i++)
					{
						#Write-Host "Working on place $($i + 1) of $placeLimit"
						$uspsaNumber = $shooters[$i].USPSANumber
						#Write-Host "placed $uspsaNumber"
						$place = Get-PlaceFull -place ([string]($i+1))
						($finalStandings | Where {($_.USPSANumber -eq $uspsaNumber) -and ($_.Division -eq $division) -and ($_.Class -eq $classFullName)}).ClassAward = "$place Place $division $classFullName"
					}
				}
			}
			else
			{
				#Write-Host "The number of shooters in this division and class ($($numberUniqueShooters)) did not meet the minimum required shoooters ($($global:classMinFirst))."
			}
		}
	}
	
}

function Write-OverallAwards
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings,
		
		[Parameter(Mandatory=$true)]
		$sectionStats
	)
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:awardsHtml
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:awardsHtml -innerHtml "Awards Qualification"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:awardsHtml -innerHtml $global:awardsDescription
	
	foreach ($division in $global:divisions)
	{
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:awardsHtml
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:awardsHtml -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:awardsHtml
		Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtml
			
		$overallShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.OverallAward -ne "")} | Sort OverallAward)
		
		if ($overallShooters -ne $null)
		{
			Write-Host $division
			
			
			Write-Host
			Write-Host "Overall"
			
			$uniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalUniqueShooters
			Generate-Html -elementType "divClass" -htmlOutputPath $global:awardsHtml
			Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:awardsHtml -innerHtml "Overall <span class=`"classUniqueShooters`">($uniqueShooters unique shooters)</span>"
			Generate-Html -elementType "divClassBody" -htmlOutputPath $global:awardsHtml
			Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtml
			
		
			$place = 1
			foreach ($overallShooter in $overallShooters)
			{
				$firstName = $overallShooter.FirstName
				$lastName = $overallShooter.LastName
				$uspsaNumber = $overallShooter.USPSANumber
				$sectionScore = $overallShooter.SectionScore
				$placeFull = Get-PlaceFull -place ([string]$place)
				$shooterOutput = "$placeFull - $firstName $lastName ($uspsaNumber) - $sectionScore%"
				Write-Host $shooterOutput
				Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtml -innerHtml $shooterOutput
				Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtml
				$place++
			}
			Write-Host
			Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtml
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
			
			foreach ($class in $global:classes)
			{
							
				$classShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $class) -and ($_.ClassAward -ne "")} | Sort ClassAward)
				
				
				if ($classShooters  -ne $null)
				{
					$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
					Write-Host $fullName
					
					$uniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $fullName)}).TotalUniqueShooters
					Generate-Html -elementType "divClass" -htmlOutputPath $global:awardsHtml
					Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:awardsHtml -innerHtml "$fullName <span class=`"classUniqueShooters`">($uniqueShooters unique shooters)</span>"
					Generate-Html -elementType "divClassBody" -htmlOutputPath $global:awardsHtml
					Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtml
				
					$place = 1
					foreach ($classShooter in $classShooters)
					{
						$firstName = $classShooter.FirstName
						$lastName = $classShooter.LastName
						$uspsaNumber = $classShooter.USPSANumber
						$sectionScore = $classShooter.SectionScore
						$placeFull = Get-PlaceFull -place ([string]$place)
						$shooterOutput = "$placeFull - $firstName $lastName ($uspsaNumber) - $sectionScore%"
						Write-Host $shooterOutput
						Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtml -innerHtml $shooterOutput
						Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtml
						$place++
					}
					Write-Host
					Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtml
					Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
					Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
				}
				else
				{
					#Write-Host "Not enough shooters for this class award."
				}
				
				#Write-Host
			}
		}
		else
		{
			Write-Host "Not enough shooters for division or class awards."
			Write-Host
			Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtml -innerHtml "Not enough shooters for division or class awards."
			Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtml
		}
		
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtml
		Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtml
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:awardsHtml
}

function Get-PlaceFull
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[string]$place
	)

	# Get last char of place
	[string]$last_char = $place.SubString($place.Length-1)
	$last_two_char = $null
	if ($place.Length -ge 2)
	{
		[string]$last_two_char = $place.SubString($place.Length-2)
	}
	
	switch ($last_char)
	{
		"1" {
				if ($last_two_char -eq "11")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "st"
				}
			}
		"2" {
				if ($last_two_char -eq "12")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "nd"
				}
			}
		"3" {
				if ($last_two_char -eq "13")
				{
					$place_end = "th"
				}
				else
				{
					$place_end = "rd"
				}
			}
		default { $place_end = "th" }
	}
	
	$place_full = "$($place)$($place_end)"
	$place_full
}

function Generate-Html
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[string]$elementType,
		
		[Parameter(Mandatory=$true)]
		[string]$htmlOutputPath,
		
		[Parameter(Mandatory=$false)]
		[string]$innerHtml
	)
	
	$html = ""
	
	switch ($elementType)
	{
		"bodystart" {
						$html += "<head>"
						$html += "<style>"
						$html += $global:style
						$html += "</head>"
						$html += "</style>"
						$html += "<body>"
					}
		"bodyEnd"	{
						$html += "</body>"
					}
		"divHeader"	{
						$html += "<div class=`"headerContainer`">$innerHtml</div>"
					}
		"divDescription"	{
						$html += "<div class=`"descriptionContainer`">$innerHtml</div>"
					}
		"divDivision"	{
						$html += "<div class=`"divisionContainer`">"
					}
		"divDivisionHeader"	{
						$html += "<div class=`"divisionHeaderContainer`">$innerHtml</div>"
					}
		"divDivisionBody"	{
						$html += "<div class=`"divisionBodyContainer`">"
					}
		"divClass"	{
						$html += "<div class=`"classContainer`">"
					}
		"divClassHeader"	{
						$html += "<div class=`"classHeaderContainer`">$innerHtml</div>"
					}
		"divClassBody"	{
						$html += "<div class=`"classBodyContainer`">"
					}
		"divEnd"	{
						$html += "</div>"
					}
		"innerHtmlP"	{
						$html += "<p>$innerHtml</p>"
					}
		"pStart"	{
						$html += "<p>"
					}
		"pEnd"	{
						$html += "</p>"
					}
		"html"		{
						$html += "$innerHtml"
					}
		"br"		{
						$html += "<br/>"
					}
	}
	
	$html | Out-File -FilePath $htmlOutputPath -Append
}


$date = (get-date -f yyyyMMdd-hhmmss)
$sectionShooterCSV = "C:\Temp\practigrab\sectionShooters.csv"
$sectionMatchesCSV = "C:\Temp\practigrab\sectionMatches.csv"
$sectionShooters = Import-CSV $sectionShooterCSV
$sectionMatches = Import-CSV $sectionMatchesCSV
$sectionMatchesConfigJson = Get-Content $SectionMatchesConfigPath | ConvertFrom-Json
$standingsRawOutputCSV = "C:\temp\practigrab\standings\sectionStandingsRaw-$($date).csv"
$finalStandingsCSV = "C:\temp\practigrab\standings\finalStandingsRaw-$($date).csv"
$global:standingByDivisionHtml = "C:\temp\practigrab\standings\standingByDivisionHtml-$($date).html"
$global:standingByClassHtml = "C:\temp\practigrab\standings\standingByClassHtml-$($date).html"
$global:awardsHtml = "C:\temp\practigrab\standings\awardsHtml-$($date).html"
$global:css = ".\nwsectionresults.css"
$global:style = Get-Content $global:css

$finalStandingsExcel = "C:\temp\practigrab\standings\finalStandings-$($date).xlsx"

$uspsaConfigPath = ".\uspsaconfig.json"
$global:uspsaConfigJson = Get-Content $uspsaConfigPath | ConvertFrom-Json
$global:divisions = $global:uspsaConfigJson.Divisions
$global:classes = $global:uspsaConfigJson.Classes
$global:overallPlaceLimit = $global:uspsaConfigJson.AwardParameters.OverallPlaceLimit
$global:overallMin = $global:uspsaConfigJson.AwardParameters.OverallMin
$global:classPlaceLimit = $global:uspsaConfigJson.AwardParameters.ClassPlaceLimit
$global:classMinFirst = $global:uspsaConfigJson.AwardParameters.ClassMinFirst
$global:ClassInterval = $global:uspsaConfigJson.AwardParameters.ClassInterval

if ($global:classPlaceLimit -eq -1)
{
	$classLimitText = "."
	$global:classPlaceLimit = 99999999
}
else
{
	$classLimitText = ", up to $($global:classPlaceLimit) places"
}



$global:awardsDescription = @"
Awards are calculated in the following format:</br>
</br>
<b>Division:</b></br>
Top 1-$($global:overallPlaceLimit) shooters ($($global:overallMin) shooter minimum)</br>
</br>
<b>Class:</b></br>
Top 1-n shooters ($($global:classMinFirst) shooter minimum, where n is increased by 1 for every $($global:ClassInterval) shooters past the initial $($global:classMinFirst) shooter minimum$($classLimitText))</br>
e.g.</br>
<$($global:classMinFirst - 1) shooters = No shooters awarded</br>
$($global:classMinFirst) shooters = 1 shooter awarded</br>
$($global:classMinFirst + $global:ClassInterval) shooter = 2 shooters awarded</br>
$($global:classMinFirst + ($global:ClassInterval * 2)) shooters = 3 shooters awarded</br>
"@

$standingsRaw = @()


foreach ($sectionMatch in $sectionMatchesConfigJson.Matches)
{
	Write-Host "Getting overall results by division for club, $($sectionMatch.Club)"
	if ($sectionMatch.InputType -eq "CSV")
	{
		$additionalMatchCSV = Import-CSV $sectionMatch.CSVPath
		foreach ($shooter in $additionalMatchCSV)
		{
			$shooter | Add-Member -MemberType NoteProperty -Name ClubOrdered -Value "$($sectionMatch.MatchNumber) - $($sectionMatch.Club)"
		}
		$additionalMatchCSV | Export-Excel $finalStandingsExcel -WorkSheetname $additionalMatchCSV[1].Club -FreezeTopRow -AutoSize
		#$additionalMatchCSV | export-csv C:\temp\practigrab\testme.csv
		$standingsRaw += $additionalMatchCSV
	}
	else
	{
		$standingsRaw += Get-StandingsRaw -sectionShooters $sectionShooters -sectionMatch $sectionMatch -excelPath $finalStandingsExcel
	}
	
}
<#
foreach ($sectionMatch in $sectionMatches)
{
	Write-Host "Getting overall results by division for club, $($sectionMatch.Club)"
	$standingsRaw += Get-StandingsRaw -sectionShooters $sectionShooters -sectionMatch $sectionMatch -excelPath $finalStandingsExcel
}

foreach ($AdditionalMatch in $AdditionalMatchFiles)
{
	Write-Host "Getting overall results by division for additional file, $AdditionalMatch"
	$additionalMatchCSV = Import-CSV $AdditionalMatch
	$additionalMatchCSV | Export-Excel $finalStandingsExcel -WorkSheetname $additionalMatchCSV[1].Club -FreezeTopRow -AutoSize
	#$additionalMatchCSV | export-csv C:\temp\practigrab\testme.csv
	$standingsRaw += $additionalMatchCSV
}#>


$standingsRaw | Export-CSV $standingsRawOutputCSV -NoTypeInformation
$standingsRaw | Export-Excel $finalStandingsExcel -WorkSheetname RawStandings -FreezeTopRow -AutoSize

$finalStandings = Process-Standings -standingsRaw $standingsRaw



$sectionStats = Calculate-SectionStats -finalStandings $finalStandings -rawStandings $standingsRaw

Calculate-OverallAwards -finalStandings $finalStandings -sectionStats $sectionStats
Calculate-ClassAwards -finalStandings $finalStandings -sectionStats $sectionStats
Calculate-OverallByDivisionPercent -finalStandings $finalStandings
Calculate-ClassByDivisionPercent -finalStandings $finalStandings

$finalStandings | Export-Excel $finalStandingsExcel -WorkSheetname FinalStandings -FreezeTopRow -AutoSize

$finalStandings | Export-CSV $finalStandingsCSV -NoTypeInformation

Write-OverallAwards -finalStandings $finalStandings -sectionStats $sectionStats

if ($PassThruRaw)
{
	$standingsRaw
}

if ($PassThruFinal)
{
	$finalStandings
}
$sectionStats

