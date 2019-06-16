[cmdletbinding(SupportsShouldProcess=$true, confirmimpact="none")] # Change confirmimpact to "high" if this function can break something
Param
(
	[Parameter(Mandatory=$false)]
	[string]$Year,
	
	[Parameter(Mandatory=$false)]
	[string]$SectionMatchesConfigPath=".\sectionMatches.json",
	
	[Parameter(Mandatory=$false)]
	[switch]$PassThruRaw,
	
	[Parameter(Mandatory=$false)]
	[switch]$PassThruFinal,
	
	[Parameter(Mandatory=$false)]
	[switch]$FinalCalc,

	[Parameter(Mandatory=$false)]
	[switch]$ForceDownload
)

# Creates the specified folder if it does not already exist.
function Create-Folder
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[string]$folderPath
	)
	
	if (!(Test-Path $folderPath)) {
		#Write-Output "Folder path, $folderPath, does not exist. Creating directory..."
		mkdir $folderPath
		
		if (!(Test-Path $folderPath)) {
			Write-Error "Directory creation failed. Exiting script..."
			Exit 1
		} else {
			#Write-Output "Directory creation succeeded..."
		}
	}
}

function getHTML ($uri, $timeoutNum=5, $postParams) {
	$count = 0
	$success = $false
	Write-Verbose "Function: getHTML; Input: uri; Value: *$uri*"
	while (($count -lt $timeoutNum-1) -and (!$success)) {
		try {
			if ($postParams)
			{
				$html = Invoke-WebRequest -Uri $uri -Method POST -Body $postParams
			}
			else
			{
				$html = Invoke-WebRequest -Uri $uri
			}
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

function Get-ActualMemberNumber ($UspsaNumber)
{
	$memberNumber = "PEN"
	if ($global:currentUspsanumber.$UspsaNumber)
	{
		Write-Verbose "Existing USPSA lookup found"
		$memberNumber = $global:currentUspsanumber.$UspsaNumber
	}
	else
	{
		Write-Verbose "Existing USPSA lookup NOT found"
		$postParams = @{Submit='lookup';number="$UspsaNumber"}
		$uri = "https://uspsa.org/uspsa-classifier-lookup-results.php"
		$html = getHTML -uri $uri -timeoutNum 5 -postParams $postParams
		#Write-Host "Retrieved html"
		
		$parsedHtmlA = $html.ParsedHtml.getElementsByTagName("a")
		#$parsedHtmlA | out-file c:\temp\href.txt

		$count = 0
		$length = $parsedHtmlA.length
		#Write-Host $length
		#$memberNumber = "NA"
		while (($count -lt $length) -and ($memberNumber -eq "PEN"))
		{
			#Write-Host "Count: $count"
			if ($parsedHtmlA[$count].href)
			{
				#Write-Host "not null"
				if ($parsedHtmlA[$count].href.Contains("?number="))
				{
					#Write-Host $parsedHtmlA[$count].href
					#$memberNumber = $parsedHtmlA[$count + 1]
					$memberNumber = $parsedHtmlA[$count].href.Split("=")[1].Split("&")[0]
					#$count = $length
				}
			}
			$count++
		}

		$global:currentUspsanumber.Add($UspsaNumber,$memberNumber)
	}
	
	Write-Verbose "$UspsaNumber - $memberNumber"
	
	return $memberNumber
}

function Write-Json ()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$Path
	)
	
}

function Get-MatchFiles ()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[string]$matchID,

		[Parameter(Mandatory=$true)]
		[string]$Name,

		[Parameter(Mandatory=$false)]
		[switch]$Force
	)

	$fileName = "$($global:seasonPath)\$($matchID)-$($Name).json"

	# If we don't want to force a redownload and the file was found locally, get the JSON from the local file.
	# Else, get the file from practiscore
	if (!$Force -and (Test-Path $fileName))
	{
		$json = Get-Content $fileName | ConvertFrom-Json
	}
	else
	{
		$uri = "https://s3.amazonaws.com/ps-scores/production/$matchID/$($Name).json"
		Write-Host "Match Def URI: $uri"
		
		$html = getHTML $uri
		$html.Content | Out-File -FilePath $fileName -Force
		$json = $html.Content | ConvertFrom-Json
	}

	return $json
}

# Blame practiscore for the following confusing set of bitwise operations. In the name of speed?
function Get-Alphas ($scoreFields)
{
	$A_MASK = 0x0000000F
	$A_MASK2 = 0x0000000F00000000
	$A_SHIFT = 0
	$A_SHIFT2 = 28

	return (($scoreFields -band $A_MASK) -shr $A_SHIFT) + (($scoreFields -band $A_MASK2) -shr $A_SHIFT2)
}

function Get-Bravos ($scoreFields)
{
	$B_MASK = 0x000000F0
	$B_MASK2 = 0x000000F000000000
	$B_SHIFT = 4
	$B_SHIFT2 = 32

	return (($scoreFields -band $B_MASK) -shr $B_SHIFT) + (($scoreFields -band $B_MASK2) -shr $B_SHIFT2)
}

function Get-Charlies ($scoreFields)
{
	$C_MASK = 0x00000F00
	$C_MASK2 = 0x00000F0000000000
	$C_SHIFT = 8
	$C_SHIFT2 = 36

	return (($scoreFields -band $C_MASK) -shr $C_SHIFT) + (($scoreFields -band $C_MASK2) -shr $C_SHIFT2)
}

function Get-Deltas ($scoreFields)
{
	$D_MASK = 0x0000F000
	$D_MASK2 = 0x0000F00000000000
	$D_SHIFT = 12
	$D_SHIFT2 = 40

	return (($scoreFields -band $D_MASK) -shr $D_SHIFT) + (($scoreFields -band $D_MASK2) -shr $D_SHIFT2)
}

function Get-NoShoots ($scoreFields)
{
	$NS_MASK = 0x000F0000
	$NS_MASK2 = 0x000F000000000000
	$NS_SHIFT = 16
	$NS_SHIFT2 = 44

	return (($scoreFields -band $NS_MASK) -shr $NS_SHIFT) + (($scoreFields -band $NS_MASK2) -shr $NS_SHIFT2)
}

function Get-Mikes ($scoreFields)
{
	$M_MASK = 0x00F00000
	$M_MASK2 = 0x00F0000000000000
	$M_SHIFT = 20
	$M_SHIFT2 = 48

	return (($scoreFields -band $M_MASK) -shr $M_SHIFT) + (($scoreFields -band $M_MASK2) -shr $M_SHIFT2)
}

function Get-NoPenaltyMikes ($scoreFields)
{
	$NPM_MASK = 0x0F000000
	$NPM_MASK2 = 0x0F00000000000000
	$NPM_SHIFT = 24
	$NPM_SHIFT2 = 52

	return (($scoreFields -band $NPM_MASK) -shr $NPM_SHIFT) 	+ (($scoreFields -band $NPM_MASK2) -shr $NPM_SHIFT2)
}

function Get-ShooterScores ()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$shooterUUID,

		[Parameter(Mandatory=$true)]
		$matchInfo
	)

	#Write-Host "Shooter UUID $shooterUUID"
	$shooterScores = $matchScoresJson.match_scores.stage_stagescores | Where-Object shtr -eq $shooterUUID
	$pophTotal = 0
	$popmTotal = 0
	$popnpmTotal = 0
	$popnsTotal = 0
	$alphas = 0
	$bravos = 0
	$charlies = 0
	$deltas = 0
	$mikes = 0
	$noShoots = 0
	$npms = 0
	$stageTime = 0.0

	$i = 1
	foreach ($shooterStageScore in $shooterScores)
	{
		#Write-Host $i
		$i++
		$poph = $shooterStageScore.poph
		$popm = $shooterStageScore.popm
		$popnpm = $shooterStageScore.popnpm
		$popns = $shooterStageScore.popns

		if ($shooterStageScore.ts)
		{
			foreach ($int in $shooterStageScore.ts)
			{
				$alphas += (Get-Alphas -scoreFields $int)
				#Write-Host $alphas
				$bravos += (Get-Bravos -scoreFields $int)
				$charlies += (Get-Charlies -scoreFields $int)
				$deltas += (Get-Deltas -scoreFields $int)
				$noShoots += (Get-NoShoots -scoreFields $int)
				$mikes += (Get-Mikes -scoreFields $int)
				$npms += (Get-NoPenaltyMikes -scoreFields $int)
			}
		}

		$alphas += $poph
		#Write-Host "End alphas stage:" -NoNewLine
		#Write-Host $alphas
		$mikes += $popm
		$npms += $popnpm
		$noShoots += $popns

		$pophTotal += $poph
		$popmTotal += $popm
		$popnpmTotal += $popnpm
		$popnsTotal += $popns

		foreach ($stringTime in $shooterStageScore.str)
		{
			$stageTime += $stringTime
		}
		if ($shooterUUID -eq "mmShooter_1987249") {Write-Host "Stage $i time: $($shooterStageScore.str[0])"}
	}

	#Write-Host "End alphas match:" -NoNewLine
	#	Write-Host $alphas

	$scores = [pscustomobject]@{
		shooterUUID = $shooterUUID
		A = $alphas
		B = $bravos
		C = $charlies
		D = $deltas
		M = $mikes
		NS = $noShoots
		NPM = $npms
		TotalTime = $stageTime
	}

	#Write-Host "Before Return:" -NoNewLine
	#Write-Host $scores.A

	return $scores
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
	
	foreach ($shooter in $matchInfo.matchShooters)
	{
		if ($null -ne $shooter.sh_id)
		{
			$uspsaNumber = $shooter.sh_id.Replace("-","")
		}
		else
		{
			$uspsaNumber = ""
		}
		$firstName = $shooter.sh_fn
		$lastName = $shooter.sh_ln
		$division = $shooter.sh_dvp
		$class = $shooter.sh_grd
		$shooterUUID = $shooter.sh_uuid
		$divPercent = ($matchInfo.matchResults.match."$division" | Where-Object shooter -eq $shooterUUID ).matchPercent
		$matchPoints = ($matchInfo.matchResults.match."$division" | Where-Object shooter -eq $shooterUUID ).matchPoints
		
		# Exceptions
		# Force-fix any data issues in input
		# eg.
		# - Shooter changes to lifetime or three year membership midseason (might have another fix to this by dropping prefix
		# - Typos in name or USPSA number
		# - Discrepencies between division or class names
		
		if ($division -eq "CO")
		{
			$division = "Carry Optics"
		}
		
		if ($uspsaNumber -eq "101809")
		{
			$uspsaNumber = "A101809"
		}
		
		if ($uspsaNumber.ToUpper().Contains("PEN"))
		{
			$uspsaNumber = "PEN"
		}
		
		if ($lastName -eq "Hong" -and $firstName -eq "Andrew")
		{
			$uspsaNumber = "A83199"
		}
		
		if ($lastName -eq "LeRoux #1" -and $firstName -eq "Scott")
		{
			$uspsaNumber = "L3253"
		}
		
		if ($lastName -eq "Niemann" -and $firstName -eq "Kamryn")
		{
			$uspsaNumber = "A101879"
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
		
		if ($lastName -eq "Domingo" -and $firstName -eq "Emilio")
		{
			$uspsaNumber = "TY86951"
		}
		
		if ($lastName -eq "Doster" -and $firstName -eq "Stephanie")
		{
			$uspsaNumber = "A96362"
		}
		
		if ($lastName -eq "Skubi" -and $firstName -eq "Bart")
		{
			$uspsaNumber = "L4061"
		}
		
		if ($lastName -eq "Tomasie" -and $firstName -eq "Squire")
		{
			$uspsaNumber = "L1145"
		}
		
		if ($lastName -eq "Blair" -and $firstName -eq "Bruce")
		{
			$uspsaNumber = "A47451"
		}
		
		if ($lastName -eq "Dong" -and $firstName -eq "James")
		{
			$uspsaNumber = "FY22573"
		}

		if ($lastName -eq "Novasky" -and $firstName -eq "Tim")
		{
			$uspsaNumber = "A100482"
		}

		if ($lastName -eq "Pukalo" -and $firstName -eq "Chris")
		{
			$uspsaNumber = "A89231"
		}
		
		#$uspsaNumber = Get-ActualMemberNumber -UspsaNumber $uspsaNumber
		
		$sectionMember = $false
		$sectionStatus = "Non-member"
		# Check to see if the shooter is in the section. Remove '-' to standardize.
		# TODO: Sanitize USPSA number to ignore membershp type prefix. Number seem to never change between TY, A, F, etc. Could use this as a truly unique value.
		
		switch ($uspsaNumber)
		{
			{$_.StartsWith("A")} { $uspsaNumberClean = $uspsaNumber.Substring(1) }
			{$_.StartsWith("B")} { $uspsaNumberClean = $uspsaNumber }
			{$_.StartsWith("F")} { $uspsaNumberClean = $uspsaNumber.Substring(1) }
			{$_.StartsWith("TY")} { $uspsaNumberClean = $uspsaNumber.Substring(2) }
		}
		#Write-Host "Uspsa number clean: " $uspsaNumberClean
		#if ($uspsaNumber -in $sectionShooters.USPSANumber.Replace("-","").Replace("A","").Replace("F","").Replace("TY",""))
		$uspsaClean = $uspsaNumber.ToUpper().Replace("-","").Replace("A","").Replace("TY","").Replace("L","").Replace("B","").Replace("FY","")
		if ($uspsaClean -in $sectionShooters.USPSANumber.ToUpper().Replace("-","").Replace("A","").Replace("TY","").Replace("L","").Replace("B","").Replace("FY",""))
		{
			$sectionMember = $true
			$sectionStatus = "Member"
		}

		$scores = Get-ShooterScores -shooterUUID $shooterUUID -matchInfo $matchInfo

		$matchShooters += [pscustomobject]@{
			USPSANumber = $uspsaNumber.ToUpper()
			USPSANumberClean = $uspsaClean
			FirstName = $firstName
			LastName = $lastName
			MatchName = $matchInfo.matchName
			Club = $matchInfo.Club
			ClubOrdered = $matchInfo.ClubOrdered
			Division = $division
			Class = $class
			DivisionPercent = [single]$divPercent
			DivisionPoints = $matchPoints
			SectionMember = $sectionMember
			SectionStatus = $sectionStatus
			A = $scores.A
			B = $scores.B
			C = $scores.C
			D = $scores.D
			M = $scores.M
			NS = $scores.NS
			NPM = $scores.NPM
			TotalTime = $scores.TotalTime
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

	if ($ForceDownload)
	{	
		Write-Host "ForceDownload option has been set. All match files will be downloaded from Practiscore."
		$matchDefJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $matchDefJsonName -Force
		$matchResultsJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $resultsJsonName -Force
		$matchScoresJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $matchScoresJsonName -Force
	}
	else
	{
		Write-Host "ForceDownload option has NOT been set. Local match files will be used if available."
		$matchDefJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $matchDefJsonName
		$matchResultsJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $resultsJsonName
		$matchScoresJson = Get-MatchFiles -matchID $sectionMatch.PractiScoreID -Name $matchScoresJsonName
	}

	$matchName = $matchDefJson.match_name
	$matchShooters = $matchDefJson.match_shooters
	$matchStages = $matchDefJson.match_stages
	$matchScores = $matchScoresJson.match_scores

	$matchInfo += [pscustomobject]@{
				matchName = $matchName
				Club = $sectionMatch.Club
				ClubOrdered = "$($sectionMatch.MatchNumber) - $($sectionMatch.Club)"
				matchShooters = $matchShooters
				matchStages = $matchStages
				matchResults = $matchResultsJson
				matchScores = $matchScores
				}
	$matchOverallByDivision = Get-OverallByDivisionPercent -sectionShooters $sectionShooters -matchInfo $matchInfo
	$matchOverallByDivision | Export-CSV "$($global:standingsDir)\$($sectionMatch.Club)-ovrbydiv.csv" -NoTypeInformation
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
	
	$clubs = $standingsRaw | Select-Object Club,ClubOrdered -Unique | Sort-Object ClubOrdered
	
	$sectionShooterResult = [pscustomobject]@{
			USPSANumber = ""
			FirstName = ""
			LastName = ""
			Division = ""
			Class = ""
			SectionScore = ""
			CurrentAverage = ""
			ScoresUsed = ""
			SectionMember = $false
			SectionStatus = "Non-Member"
			OverallAward = ""
			ClassAward = ""
			}
	
	foreach ($club in $clubs)
	{
		$sectionShooterResult | Add-Member -MemberType NoteProperty -Name $club.club -Value ""
	}

	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "MatchPoints" -Value 0.0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "TotalTime" -Value 0.0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "A" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "B" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "C" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "D" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "M" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "NS" -Value 0
	$sectionShooterResult | Add-Member -MemberType NoteProperty -Name "NPM" -Value 0
	
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
		$sectionShooters,
		
		[Parameter(Mandatory=$false)]
		[int]$BestXOf=4
	)
	
	$standingsRaw | foreach-object {$_.DivisionPercent = [single]$_.DivisionPercent}
	
	$uspsaNumbers = ($standingsRaw | Where-Object {($_.DivisionPercent -ne 0) -and ($_.USPSAnumber -ne "") -and ($_.USPSAnumber -ne "PEN")} | Select-Object USPSAnumber -Unique | Sort-Object).USPSANumber
	$shooterStandingObj = Build-MasterSheet -standingsRaw $standingsRaw
	$finalStandings = @()

	foreach ($uspsaNumber in $uspsaNumbers)
	{
		# This year shooters may have enough scores for multiple divisions. Make sure we separate out divisions.
		$shooterDivs = ($standingsRaw | Where-Object {($_.USPSANumber -eq $uspsaNumber) -and ($_.DivisionPercent -ne 0)} | Select-Object Division -Unique).Division
		
		foreach ($division in $shooterDivs)
		{
			#Write-Host "Calculating average scores for shooter, $uspsaNumber"
			$shooterStanding = $shooterStandingObj.PsObject.Copy()
			$shooterResults = $standingsRaw | Where-Object {($_.USPSANumber -eq $uspsaNumber) -and ($_.DivisionPercent -ne 0) -and $_.Division -eq $division} | Sort-Object ClubOrdered 
			$bestOfResults = @()
			$bestOfResults += $shooterResults | Sort-Object DivisionPercent -Descending | Select-Object -First $BestXOf
			
			if ($bestOfResults.length -lt $BestXOf)
			{
				[single]$average = $null
				$shooterStanding.ScoresUsed = "Not eligible for series score. $($bestOfResults.length) out of $BestXOf required matches."
				$averageObj = $shooterResults.DivisionPercent | Measure-Object -Average
				[single]$currentAverage = [single]([math]::Round($averageObj.Average,2))
				$shooterStanding.CurrentAverage = $currentAverage
				
				if ($uspsaNumber -eq  "A85001")
				{
					$shooterResults.DivisionPercent | Out-File C:\temp\a85001.txt -Append
					$averageObj | Out-File C:\temp\a85001.txt -Append
					$currentAverage  | Out-File C:\temp\a85001.txt -Append
					Write-Host "Done!" -foregroundcolor red
				}
			}
			else
			{
				$averageObj = $bestOfResults.DivisionPercent | Measure-Object -Average
				[single]$average = [single]([math]::Round($averageObj.Average,2))
				$shooterStanding.ScoresUsed = $bestOfResults.Club -join ';'
				$shooterStanding.CurrentAverage = $average
				if ($uspsaNumber -eq  "A85001")
				{
					$shooterResults | Out-File C:\temp\a85001.txt -Append
					$shooterDivs | Out-File C:\temp\a85001.txt -Append
					$bestOfResults | Out-File C:\temp\a85001.txt -Append
					$bestOfResults.DivisionPercent | Out-File C:\temp\a85001.txt -Append
					$averageObj | Out-File C:\temp\a85001.txt -Append
					$average  | Out-File C:\temp\a85001.txt -Append
					Write-Host "Done!" -foregroundcolor red
				}
			}
			#Write-Host "Average is, $average"
			
			
			foreach ($shooterResult in $shooterResults)
			{ 
				$shooterStanding."$($shooterResult.Club)" = $shooterResult.DivisionPercent
				$shooterStanding.A += $shooterResult.A
				$shooterStanding.B += $shooterResult.B
				$shooterStanding.C += $shooterResult.C
				$shooterStanding.D += $shooterResult.D
				$shooterStanding.M += $shooterResult.M
				$shooterStanding.NS += $shooterResult.NS
				$shooterStanding.NPM += $shooterResult.NPM
				$shooterStanding.TotalTime += $shooterResult.TotalTime
				$shooterStanding.MatchPoints += $shooterResult.DivisionPoints
			}
			
			$shooterStanding.USPSANumber = $uspsaNumber.Replace("-","")
			$shooterStanding.FirstName = $shooterResults[0].FirstName.Substring(0,1).ToUpper()+$shooterResults[0].FirstName.Substring(1).ToLower()
			$shooterStanding.LastName = $shooterResults[0].LastName.Substring(0,1).ToUpper()+$shooterResults[0].LastName.Substring(1).ToLower()
			$shooterStanding.Division = $division
			$shooterStanding.SectionScore = $average
			
			if ($shooterResults[0].Class -eq "U")
			{
				# Exception for known unclassified that now have classifications
				# Check with actual USPSA classifier will be added later
				$class = "U"
				
				$shooterStanding.Class = $class
			}
			else
			{
				$shooterStanding.Class = $shooterResults[0].Class
			}
			
			# Check to see if the shooter is in the section. Remove '-' to standardize.
			# TODO: Sanitize USPSA number to ignore membershp type prefix. Number seem to never change between TY, A, F, etc. Could use this as a truly unique value.
			$uspsaClean = $uspsaNumber.ToUpper().Replace("-","").Replace("A","").Replace("TY","").Replace("L","").Replace("B","").Replace("FY","")
			if ($uspsaClean -in $sectionShooters.USPSANumber.ToUpper().Replace("-","").Replace("A","").Replace("TY","").Replace("L","").Replace("B","").Replace("FY",""))
			{
				$shooterStanding.SectionMember = $true
				$shooterStanding.SectionStatus = "Member"
			}
			<#if ($uspsaNumber -in $sectionShooters.USPSANumber.Replace("-",""))
			{
				$shooterStanding.SectionMember = $true
				$shooterStanding.SectionStatus = "Member"
			}#>
			
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
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml "Overall Results By Division"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml "* indicates the shooter qualified for a division or class award. Refer to the awards section for details."
	#Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml "Overall Results By Division"
	
	foreach ($division in $global:divisions)
	{
		Write-Debug $division
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		
		$shooters = @()
		$shooters += $finalStandings | Where-Object {($_.Division -eq $division) -and ($_.SectionScore -gt 0)} | Sort-Object SectionScore -Descending
		if ($null -ne $shooters)
		{
			$numEligibleShooters = $shooters.Length
			#Write-Debug "$numEligibleShooters eligible shooters"

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
				Write-Debug $shooterOutput
				Generate-Html -elementType "html" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml $shooterOutput
				Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
				$place++
			}
			
		}
		else
		{
			Write-Debug "No eligible shooters."
			Generate-Html -elementType "html" -htmlOutputPath $global:standingByDivisionHtmlOutputPath -innerHtml "No eligible shooters."
			Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		}
		
		
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		Generate-Html -elementType "br" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:standingByDivisionHtmlOutputPath
}

function Calculate-ClassByDivisionPercent
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$finalStandings
	)
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:standingByClassHtmlOutputPath
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml "Class Results By Division"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml "* indicates the shooter qualified for a division or class award. Refer to the awards section for details."
	
	foreach ($division in $global:divisions)
	{
		Write-Debug $division
		
		
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:standingByClassHtmlOutputPath
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:standingByClassHtmlOutputPath
		Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByClassHtmlOutputPath
		
		foreach ($class in $global:classes)
		{
			$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			Write-Debug $fullName
			
			$uniqueShooters = @()
			$uniqueShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $class)} | Select USPSANumber -Unique
			$numUniqueShooters = $uniqueShooters.Count
			
			Generate-Html -elementType "divClass" -htmlOutputPath $global:standingByClassHtmlOutputPath
			Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml "$fullName <span class=`"classUniqueShooters`">($numUniqueShooters unique shooters)</span>"
			Generate-Html -elementType "divClassBody" -htmlOutputPath $global:standingByClassHtmlOutputPath
			Generate-Html -elementType "pStart" -htmlOutputPath $global:standingByClassHtmlOutputPath
			
			
			
			
			$eligibleShooters = @()
			$eligibleShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.Class -eq $class)} | Sort SectionScore -Descending
			if ($eligibleShooters -ne $null)
			{
				$numEligibleShooters = $eligibleShooters.Length
				Write-Debug "$numEligibleShooters eligible shooters"
				
				
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
					Write-Debug $shooterOutput
					Generate-Html -elementType "html" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml $shooterOutput
					Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtmlOutputPath
					$place++
				}
			}
			else
			{
				Write-Debug "No eligible shooters."
				Generate-Html -elementType "html" -htmlOutputPath $global:standingByClassHtmlOutputPath -innerHtml "No eligible shooters."
				Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtmlOutputPath
			}
			
			
			Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
		}
		
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
		Generate-Html -elementType "br" -htmlOutputPath $global:standingByClassHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:standingByClassHtmlOutputPath
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
		#Write-Debug $division
		
		$sectionShooterResult = [pscustomobject]@{
			Division = $division
			Class = "Overall"
			TotalUniqueShooters = @($rawStandings | Where {($_.Division -eq $division) -and ($_.USPSANumber -ne "") -and ($_.USPSANumber -ne "PEN")} | Select USPSANumber -Unique).Count
			TotalEligibleShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0)}).Count
			TotalEligibleSectionShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.SectionMember)}).Count
		}
		$sectionStats += $sectionShooterResult
		
		foreach ($class in $global:classes)
		{
			$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			#Write-Debug $fullName
			
			$uniqueShooters = @()
			$uniqueShooters += $rawStandings | Where {($_.Division -eq $division) -and ($_.USPSANumber -ne "") -and ($_.USPSANumber -ne "PEN") -and ($_.Class -eq $class)} | Select USPSANumber -Unique
			$numUniqueShooters = $uniqueShooters.Count
			
			$eligibleShooters = @()
			$eligibleShooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.Class -eq $class)} | Sort SectionScore -Descending
			$numEligibleShooters = $eligibleShooters.Length
			
			$eligibleShootersSection = @()
			$eligibleShootersSection += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.Class -eq $class) -and ($_.SectionMember)} | Sort SectionScore -Descending
			$numEligibleShootersSection = $eligibleShootersSection.Length
			
			$sectionShooterResult = [pscustomobject]@{
				Division = $division
				Class = $fullName
				TotalUniqueShooters = $numUniqueShooters
				TotalEligibleShooters = $numEligibleShooters
				TotalEligibleSectionShooters = $numEligibleShootersSection
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
	
	#Write-Debug "Overall Awards Calc"
	foreach ($division in $global:divisions)
	{
		#Write-Debug "Division: $division"
		$numberUniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalUniqueShooters
		$numberEligibleShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalEligibleSectionShooters
		
		if ($numberUniqueShooters -ge $global:overallMin)
		{
			#Write-Debug "The number of shooters in this division ($($numberUniqueShooters)) met the minimum required shoooters ($($global:overallMin))."
			$shooters = @()
			$shooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.SectionScore -gt 0) -and ($_.SectionMember)} | Sort SectionScore -Descending
			if ($shooters -ne $null)
			{
				$numShooters = $shooters.Length
				if ($numShooters -lt $global:overallPlaceLimit) { $placeLimit = $numShooters }
				else { $placeLimit = $global:overallPlaceLimit }
				
				if ($placeLimit -gt $numberEligibleShooters)
				{
					$placeLimit = $numberEligibleShooters
				}
				
				#Write-Debug "numShooters: $numShooters"
				#Write-Debug "PlaceLimit: $placeLimit"
				
				for ($i = 0; $i -lt $placeLimit; $i++)
				{
					#Write-Debug "Working on place $($i + 1) of $placeLimit"
					$uspsaNumber = $shooters[$i].USPSANumber
					#Write-Debug "placed $uspsaNumber"
					$place = Get-PlaceFull -place ([string]($i+1))
					($finalStandings | Where {($_.USPSANumber -eq $uspsaNumber) -and $_.Division -eq $division}).OverallAward = "$place Place $division Overall"
				}
			}
		}
		else
		{
			#Write-Debug "The number of shooters in this division ($($numberUniqueShooters)) did not meet the minimum required shoooters ($($global:overallMin))."
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
	
	Write-Debug "Class Awards Calc"
	foreach ($division in $global:divisions)
	{
		Write-Debug "Division: $division"
		
		foreach ($class in ($global:classes | Where {$_ -ne "U"}))
		{
			$classFullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
			Write-Debug "Class: $classFullName"
			
			$numberUniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $classFullName)}).TotalUniqueShooters
			$numberEligibleShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $classFullName)}).TotalEligibleSectionShooters

			
			if ($numberUniqueShooters -ge $global:classMinFirst)
			{
				Write-Debug "The number of shooters in this division ($($numberUniqueShooters)) met the minimum required shoooters ($($global:classMinFirst))."
				
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
				$shooters += $finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $class) -and ($_.SectionScore -gt 0) -and ($_.OverallAward -eq "") -and ($_.SectionMember)} | Sort SectionScore -Descending
				
				$finalStandings | Export-CSV "C:\Temp\Update-NWSectionResults\temp\final-$($division).$($class).csv"
				$shooters | Export-CSV "C:\Temp\Update-NWSectionResults\temp\shooters-$($division).$($class).csv"
				
				
				if ($placeLimit -gt $shooters.Count)
				{
					$placeLimit = $shooters.Count
				}
				
				if ($shooters -ne $null)
				{
					$numShooters = $shooters.Length
					
					Write-Debug "numShooters: $numShooters"
					Write-Debug "PlaceLimit: $placeLimit"
					
					for ($i = 0; $i -lt $placeLimit; $i++)
					{
						Write-Debug "Working on place $($i + 1) of $placeLimit"
						$uspsaNumber = $shooters[$i].USPSANumber
						Write-Debug "placed $uspsaNumber"
						$place = Get-PlaceFull -place ([string]($i+1))
						($finalStandings | Where {($_.USPSANumber -eq $uspsaNumber) -and ($_.Division -eq $division) -and ($_.Class -eq $class)}).ClassAward = "$place Place $division $classFullName"
					}
				}
			}
			else
			{
				#Write-Debug "The number of shooters in this division and class ($($numberUniqueShooters)) did not meet the minimum required shoooters ($($global:classMinFirst))."
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
	
	Generate-Html -elementType "bodystart" -htmlOutputPath $global:awardsHtmlOutputPath
	Generate-Html -elementType "divHeader" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml "Awards Qualification"
	Generate-Html -elementType "divDescription" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml $global:awardsDescription
	
	foreach ($division in $global:divisions)
	{
		Generate-Html -elementType "divDivision" -htmlOutputPath $global:awardsHtmlOutputPath
		Generate-Html -elementType "divDivisionHeader" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml $division
		Generate-Html -elementType "divDivisionBody" -htmlOutputPath $global:awardsHtmlOutputPath
		Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtmlOutputPath
			
		$overallShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.OverallAward -ne "")} | Sort OverallAward)
		
		if ($overallShooters -ne $null)
		{
			Write-Debug $division
			
			
			
			Write-Debug "Overall"
			
			$uniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq "Overall")}).TotalUniqueShooters
			Generate-Html -elementType "divClass" -htmlOutputPath $global:awardsHtmlOutputPath
			Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml "Overall <span class=`"classUniqueShooters`">($uniqueShooters unique shooters)</span>"
			Generate-Html -elementType "divClassBody" -htmlOutputPath $global:awardsHtmlOutputPath
			Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtmlOutputPath
			
		
			$place = 1
			foreach ($overallShooter in $overallShooters)
			{
				$firstName = $overallShooter.FirstName
				$lastName = $overallShooter.LastName
				$uspsaNumber = $overallShooter.USPSANumber
				$sectionScore = $overallShooter.SectionScore
				$placeFull = Get-PlaceFull -place ([string]$place)
				$shooterOutput = "$placeFull - $firstName $lastName ($uspsaNumber) - $sectionScore%"
				Write-Debug $shooterOutput
				Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml $shooterOutput
				Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtmlOutputPath
				$place++
			}
			
			Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtmlOutputPath
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
			Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
			
			foreach ($class in $global:classes)
			{
							
				$classShooters = @($finalStandings | Where {($_.Division -eq $division) -and ($_.Class -eq $class) -and ($_.ClassAward -ne "")} | Sort ClassAward)
				
				
				if ($classShooters  -ne $null)
				{
					$fullName = $global:uspsaConfigJson.ClassesAttributes.$class.FullName
					Write-Debug $fullName
					
					$uniqueShooters = ($sectionStats | Where {($_.Division -eq $division) -and ($_.Class -eq $fullName)}).TotalUniqueShooters
					Generate-Html -elementType "divClass" -htmlOutputPath $global:awardsHtmlOutputPath
					Generate-Html -elementType "divClassHeader" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml "$fullName <span class=`"classUniqueShooters`">($uniqueShooters unique shooters)</span>"
					Generate-Html -elementType "divClassBody" -htmlOutputPath $global:awardsHtmlOutputPath
					Generate-Html -elementType "pStart" -htmlOutputPath $global:awardsHtmlOutputPath
				
					$place = 1
					foreach ($classShooter in $classShooters)
					{
						$firstName = $classShooter.FirstName
						$lastName = $classShooter.LastName
						$uspsaNumber = $classShooter.USPSANumber
						$sectionScore = $classShooter.SectionScore
						$placeFull = Get-PlaceFull -place ([string]$place)
						$shooterOutput = "$placeFull - $firstName $lastName ($uspsaNumber) - $sectionScore%"
						Write-Debug $shooterOutput
						Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml $shooterOutput
						Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtmlOutputPath
						$place++
					}
					
					Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtmlOutputPath
					Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
					Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
				}
				else
				{
					#Write-Debug "Not enough shooters for this class award."
				}
				
				#Write-Debug
			}
		}
		else
		{
			Write-Debug "Not enough shooters for division or class awards."
			
			Generate-Html -elementType "html" -htmlOutputPath $global:awardsHtmlOutputPath -innerHtml "Not enough shooters for division or class awards."
			Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtmlOutputPath
		}
		
		Generate-Html -elementType "pEnd" -htmlOutputPath $global:awardsHtmlOutputPath
		Generate-Html -elementType "br" -htmlOutputPath $global:awardsHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
		Generate-Html -elementType "divEnd" -htmlOutputPath $global:awardsHtmlOutputPath
	}
	
	Generate-Html -elementType "bodyEnd" -htmlOutputPath $global:awardsHtmlOutputPath
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

function Generate-MatchListHtml
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$sectionMatchesConfigJson
	)
	
	$matchInfoList = @()

	foreach ($sectionMatch in $sectionMatchesConfigJson."$Season".Matches)
	{

		$matchInfo = [pscustomobject]@{
				"Match #" = $sectionMatch.MatchNumber
				Club = $sectionMatch.Club
				Date = $sectionMatch.MatchDate
				"PractiScore Link" = "practiLink-$($sectionMatch.MatchNumber)"
				"USPSA Link" = "uspsaLink-$($sectionMatch.MatchNumber)"
				"ChallengeMatch?" = $sectionMatch.Championship
			}
			
		$matchInfoList += $matchInfo
	}

	$matchInfoHtml = $matchInfoList | ConvertTo-HTML -Fragment
	#$matchInfoList
	
	foreach ($sectionMatch in $sectionMatchesConfigJson."$Season".Matches)
	{
		if ($sectionMatch.PractiScoreURL)
		{
			$matchInfoHtml = $matchInfoHtml -replace "practiLink-$($sectionMatch.MatchNumber)", "<a href=`"$($sectionMatch.PractiScoreURL)`">Link</a>"
		}
		else
		{
			$matchInfoHtml = $matchInfoHtml -replace "practiLink-$($sectionMatch.MatchNumber)", "NA"
		}
		
		if ($sectionMatch.UspsaURL)
		{
			$matchInfoHtml = $matchInfoHtml -replace "uspsaLink-$($sectionMatch.MatchNumber)", "<a href=`"$($sectionMatch.UspsaURL)`">Link</a>"
		}
		else
		{
			$matchInfoHtml = $matchInfoHtml -replace "uspsaLink-$($sectionMatch.MatchNumber)", "NA"
		}
	}
	
	return $matchInfoHtml
}

function Get-UspsaNumberWithoutPrefix
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[string]$UspsaNumber
	)

	$uspsaNoPrefix = ""

	if ($UspsaNumber.ToUpper().StartsWith("A"))
	{
		$uspsaNoPrefix = $UspsaNumber.Replace("A","")
	}
	elseif ($UspsaNumber.ToUpper().StartsWith("TY"))
	{
		$uspsaNoPrefix = $UspsaNumber.Replace("TY","")
	}
	else
	{

	}

	return $uspsaNoPrefix
}

function Scrub-UspsaNumbers
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		$standingsRaw
	)
	
	$uniqueUspsaNumberNoPrefix = $standingsRaw | Select-Object USPSANumberClean -Unique
	$replaceTy = $false

	foreach ($standing in $standingsRaw)
	{
		$uniqueShooter = $standingsRaw | Where-Object USPSANumberClean -eq $standing.USPSANumberClean
		$fullUspsaNumbers = $uniqueShooter.USPSANumber
		$tyNumber = "TY$($standing.USPSANumberClean)"

		if ($fullUspsaNumbers.Contains($tyNumber))
		{
			Write-Verbose "Changing $($standing.USPSANumber) to $($standing.USPSANumber.ToUpper().Replace("A","TY"))"
			$standing.USPSANumber = $standing.USPSANumber.ToUpper().Replace("A","TY")
		}
	}
	
	return $standingsRaw

	<#foreach ($uspsaNumber in $uniqueUspsaNumberNoPrefix)
	{
		$uniqueShooter = $standingsRaw | Where-Object USPSANumberClean -eq $uspsaNumber
		$fullUspsaNumbers = $uniqueShooter | Select-Object USPSANumber
		$aNumber = "A$uspsaNumber"
		$tyNumber = "TY$uspsaNumber"
		if ($fullUspsaNumbers.Contains("TY"))
		{
			
		}
	}#>
}

function Get-LeaderBoardHtml ()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)]
		[PSCustomObject[]]$FinalStandings,

		[Parameter(Mandatory=$true)]
		[string]$PropertyName,

		[Parameter(Mandatory=$true)]
		[string]$FriendlyColumnName,

		[Parameter(Mandatory=$false)]
		[int]$Top = 10
	)

	$shooterList = $finalStandings | Sort-Object -Property $PropertyName -Descending | Select-Object -First $Top

	$place = 1
	$leaderBoardList = @()
	foreach ($shooter in $shooterList)
	{
		$leaderBoardList += [pscustomobject]@{
			Place = $place
			"$FriendlyColumnName" = $shooter.$PropertyName
			Name = $shooter.FirstName + " " + $shooter.LastName
			Division = $shooter.Division
		}
		$place++
	}

	return ($leaderBoardList | ConvertTo-HTML -Fragment)
}



$date = (get-date -f yyyyMMdd-hhmmss)
$global:scriptName = "Update-NWSectionResults"
$global:tempDir = "C:\temp"
$global:outputDir = "$($global:tempDir)\$($global:scriptName)"
$global:standingsDir = "$($global:outputDir)\standings"
$global:workingDir = (Get-Location).Path
$sectionShooterCsvPath = "$($global:outputDir)\sectionShooters.csv"
$sectionShooters = Import-CSV $sectionShooterCsvPath
$sectionMatchesConfigJson = Get-Content $SectionMatchesConfigPath | ConvertFrom-Json
$standingsRawOutputCsvPath = "$($global:standingsDir)\sectionStandingsRaw-$($date).csv"
$finalStandingsCsvPath = "$($global:standingsDir)\finalStandingsRaw-$($date).csv"
$global:standingByDivisionHtmlOutputPath = "$($global:standingsDir)\standingByDivisionHtml-$($date).html"
$global:standingByClassHtmlOutputPath = "$($global:standingsDir)\standingByClassHtml-$($date).html"
$global:awardsHtmlOutputPath = "$($global:standingsDir)\awardsHtml-$($date).html"
$global:cssPath = "$($global:outputDir)\nwsectionresults.css"
$global:style = Get-Content $global:cssPath
$global:currentUspsanumber = @{}
$global:matchScoresJsonName = "match_scores"
$global:resultsJsonName = "results"
$global:matchDefJsonName = "match_def"



Create-Folder -folderPath $global:tempDir
Create-Folder -folderPath $global:outputDir
Create-Folder -folderPath $global:standingsDir


$finalStandingsExcel = "$($global:standingsDir)\finalStandings-$($date).xlsx"


$uspsaConfigPath = ".\uspsaconfig.json"
$global:uspsaConfigJson = Get-Content $uspsaConfigPath | ConvertFrom-Json

if ($Year)
{
	Write-Host "Season override to $Year"
	$global:season = $Year
}
else
{
	$global:season = $global:uspsaConfigJson.Season
	Write-Host "Season set to " $global:season
}

$global:divisions = $global:uspsaConfigJson.Divisions
$global:classes = $global:uspsaConfigJson.Classes
$global:overallPlaceLimit = $global:uspsaConfigJson.AwardParameters.OverallPlaceLimit
$global:overallMin = $global:uspsaConfigJson.AwardParameters.OverallMin
$global:classPlaceLimit = $global:uspsaConfigJson.AwardParameters.ClassPlaceLimit
$global:classMinFirst = $global:uspsaConfigJson.AwardParameters.ClassMinFirst
$global:ClassInterval = $global:uspsaConfigJson.AwardParameters.ClassInterval
$championshipYear = $sectionMatchesConfigJson.$Season.Championship


if ($championshipYear)
{
	Write-Host "This is a NW Challenge year" $championshipYear
	$global:bestXOf = $global:uspsaConfigJson.Eligibility.BestXOf
}
else
{
	Write-Host "This is NOT a NW Challenge year"
	$global:bestXOf = $global:uspsaConfigJson.Eligibility.BestXOfNoChallenge
}

Write-Host "Using best " $global:bestXOf " of n scores."

# Public HTML URLs
$htmlLocalRepoDir = "C:\Repos\Github\nwsectionresults"
Create-Folder -folderPath "$($htmlLocalRepoDir)\$season"

$indexHtmlNewPath = "$($htmlLocalRepoDir)\$season\index.html"
$shooterStatHtmlSourcePath = "$($htmlLocalRepoDir)\shooter-breakdown-source.html"
$shooterStatHtmlNewPath = "$($htmlLocalRepoDir)\$season\shooter-breakdown.html"
$awardsHtmlPath = "$($htmlLocalRepoDir)\$season\awards.html"
$standingByDivisionHtmlPath = "$($htmlLocalRepoDir)\$season\standingByDivision.html"
$standingByClassHtmlPath = "$($htmlLocalRepoDir)\$season\standingByClass.html"
$finalStandingsRawHtmlSourcePath = "$($htmlLocalRepoDir)\finalstandingsraw-source.html"
$finalStandingsRawHtmlNewPath = "$($htmlLocalRepoDir)\$season\finalstandingsraw.html"
$leaderboardHtmlSourcePath = "$($htmlLocalRepoDir)\leaderboard.html"
$leaderboardHtmlNewPath = "$($htmlLocalRepoDir)\$season\leaderboard.html"
$global:seasonPath = "$($htmlLocalRepoDir)\$season"

if ($FinalCalc)
{
	$indexHtmlSourcePath = "$($htmlLocalRepoDir)\index-source.html"
}
else
{
	$indexHtmlSourcePath = "$($htmlLocalRepoDir)\index-source-midseason.html"
}


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
<b>Who's Eligible?</b> Northwest Section Members</br>
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

Write-Host "Processing section matches"
foreach ($sectionMatch in $sectionMatchesConfigJson.$Season.Matches)
{
	Write-Host "Getting overall results by division for club, $($sectionMatch.Club)"
	if ($sectionMatch.InputType -eq "CSV")
	{
		$additionalMatchCSV = Import-CSV $sectionMatch.CSVPath
		foreach ($shooter in $additionalMatchCSV)
		{
			$shooter | Add-Member -MemberType NoteProperty -Name ClubOrdered -Value "$($sectionMatch.MatchNumber) - $($sectionMatch.Club)"
			$sectionMember = $false
			$sectionStatus = "Non-Member"
			# Check to see if the shooter is in the section. Remove '-' to standardize.
			# TODO: Sanitize USPSA number to ignore membershp type prefix. Number seem to never change between TY, A, F, etc. Could use this as a truly unique value.
			if ($shooter.USPSANumber -in $sectionShooters.USPSANumber.Replace("-",""))
			{
				$sectionMember = $true
				$sectionStatus = "Member"
			}
			$shooter | Add-Member -MemberType NoteProperty -Name SectionMember -Value $sectionMember
			$shooter | Add-Member -MemberType NoteProperty -Name SectionStatus -Value $sectionStatus
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

$standingsRaw | Export-CSV "$standingsRawOutputCsvPath-pre.csv" -NoTypeInformation
$standingsRaw = Scrub-UspsaNumbers -standingsRaw $standingsRaw

Write-Host "Writing raw standings to file."
$standingsRaw | Export-CSV $standingsRawOutputCsvPath -NoTypeInformation
$standingsRaw | Export-Excel $finalStandingsExcel -WorkSheetname RawStandings -FreezeTopRow -AutoSize

Write-Host "Processing final standings."
$finalStandings = Process-Standings -standingsRaw $standingsRaw -sectionShooters $sectionShooters -BestXOf $global:bestXOf

Write-Host "Calculating section stats."
$sectionStats = Calculate-SectionStats -finalStandings $finalStandings -rawStandings $standingsRaw

Write-Host "Calculating overall awards."
Calculate-OverallAwards -finalStandings $finalStandings -sectionStats $sectionStats

Write-Host "Calculating class awards."
Calculate-ClassAwards -finalStandings $finalStandings -sectionStats $sectionStats

Write-Host "Calculating division standings awards."
Calculate-OverallByDivisionPercent -finalStandings $finalStandings

Write-Host "Calculating class standings."
Calculate-ClassByDivisionPercent -finalStandings $finalStandings

Write-Host "Writing final standings to file."
$finalStandings | Export-Excel $finalStandingsExcel -WorkSheetname FinalStandings -FreezeTopRow -AutoSize
$finalStandings | Export-CSV $finalStandingsCsvPath -NoTypeInformation

Write-Host "Generating awards html."
Write-OverallAwards -finalStandings $finalStandings -sectionStats $sectionStats

Write-Host "Writing index.html file"
$newIndex = Get-Content $indexHtmlSourcePath
$matchListHtml = Generate-MatchListHtml -sectionMatchesConfigJson $sectionMatchesConfigJson
$newIndex = $newIndex -replace "\[matchListTable\]", $matchListHtml
$newIndex = $newIndex -replace "\[season\]", $season
$newIndex | Out-File $indexHtmlNewPath

Write-Host "Writing shooter-breakdown.html file"
$newShooterStat = Get-Content $shooterStatHtmlSourcePath
$sectionStatHtml = $sectionStats | ConvertTo-HTML -Fragment
$newShooterStat = $newShooterStat -replace "\[shooterBreakdown\]", $sectionStatHtml
$newShooterStat | Out-File $shooterStatHtmlNewPath

Write-Host "Writing final standings raw html file"
$newFinalHtml = Get-Content $finalStandingsRawHtmlSourcePath
$finalStandingsHtml = $finalStandings | Sort-Object LastName,FirstName | ConvertTo-HTML -Fragment
$newFinalHtml = $newFinalHtml -replace "\[finalStandingsRaw\]", $finalStandingsHtml
$newFinalHtml = $newFinalHtml -replace "\[season\]", $season
$newFinalHtml | Out-File $finalStandingsRawHtmlNewPath

Write-Host "Writing Leader Boards"
$newLeaderHtml = Get-Content $leaderboardHtmlSourcePath
$hoaHtml = Get-LeaderBoardHtml -FinalStandings $finalStandings -PropertyName "MatchPoints" -FriendlyColumnName "Total Match Points"
$mostAccurateHtml = Get-LeaderBoardHtml -FinalStandings $finalStandings -PropertyName "A" -FriendlyColumnName "Total Alphas"
$noshootHtml = Get-LeaderBoardHtml -FinalStandings $finalStandings -PropertyName "NS" -FriendlyColumnName "Total No Shoots"
$deltaHtml = Get-LeaderBoardHtml -FinalStandings $finalStandings -PropertyName "D" -FriendlyColumnName "Total Deltas"
$newLeaderHtml = $newLeaderHtml -replace "\[hoa\]", $hoaHtml
$newLeaderHtml = $newLeaderHtml -replace "\[mostaccurate\]", $mostAccurateHtml
$newLeaderHtml = $newLeaderHtml -replace "\[ns\]", $noshootHtml
$newLeaderHtml = $newLeaderHtml -replace "\[deltas\]", $deltaHtml
$newLeaderHtml | Out-File $leaderboardHtmlNewPath

Write-Host "Copying other web files to repo"
Copy-Item -Path $global:standingByDivisionHtmlOutputPath -Destination $standingByDivisionHtmlPath
Copy-Item -Path $global:standingByClassHtmlOutputPath -Destination $standingByClassHtmlPath
Copy-Item -Path $global:awardsHtmlOutputPath -Destination $awardsHtmlPath


if ($PassThruRaw)
{
	$standingsRaw
}

if ($PassThruFinal)
{
	$finalStandings
}
#$sectionStats | ft



