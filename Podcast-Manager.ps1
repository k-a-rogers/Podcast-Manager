#Podcast-Manager
param(
	[Parameter(Mandatory=$false)][boolean]$menu=$false
)
###############
# Variables etc
###############

# Module import

Ipmo BitsTransfer

# Config file import
[System.Collections.ArrayList]$global:config=Import-CSV -Path $($PSScriptRoot+"\Config.txt") -Delimiter ";"
$global:pendingepisodes=@()

# Local and remote directory configurations
$podcastdir="<top-level directory for podcast storage>"
$playerdir=$null


###############
# Functions
###############


function Show-Menu {
	# Define the available functions within the menu
	$global:menu= @"
Podcast Manager

1.	Check for available episodes of all podcasts 
2.	Check for available episodes of a specific podcast
3.	Show information for most recent available episodes of all podcasts
4.	Download available episodes of current podcasts
5.	View managed podcasts
6.	Add new managed podcast
7.	Remove managed podcast
8.	Quit

"@
	# Wrap menu in a while-loop
	$quit=$false
	do {
		Clear-Host
		Write-Host $menu
		$choice = Read-Host("Please choose from 1-8")

		switch ($choice) {
			"1" {
				for ($i=0;$i -lt $global:config.count;$i++) {
					Get-PendingEpisodes -podchoice $i
				}
				Start-sleep 5
			}
			"2" {
				# Display names of podcast, prompt for numerical selection, then invoke Get-PendingEpisodes.
				cls
				[boolean]$valid=$false
				while (!$valid) {
					Write-Host "The following podcasts are managed:`n"
					for ($i=0;$i -lt $global:config.count;$i++) {
						Write-Host "$($i+1): $($global:config[$i].Name)"
					}
					Write-Host ""
					[int]$choice=Read-Host("Please select the desired podcast")
					$choice--
					if ($global:config[$choice]) {
						$valid=$true;
						Get-PendingEpisodes -podchoice $choice
					} else {
						Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
						Start-sleep 3
						cls
					}
				}
				Start-Sleep 5
				Remove-Variable -Name valid,choice -Force -ErrorAction SilentlyContinue
			}
			"3" {
				Show-LatestEpisodes;
			}
			"4" {
				Download-PendingEpisodes;
			}
			"5" {
				View-Podcasts;
			}
			"6" {
				Add-Podcast;
			}
			"7" {
				Remove-Podcast;
			}
			"8" {
				Write-Host "Quitting..."
				$quit = $true
			}
			default {
				Write-Host "Invalid selection entered, please try again."
			}
		}
		Start-sleep 1
	} while (!$quit)
}

Function Get-PendingEpisodes {
	param(
		[Parameter(Mandatory=$true)][string]$podchoice
	)
	$podcast=$global:config[$podchoice]
	Write-Host "Podcast:`t$($podcast.Name)"
	# Feed retrieval and new episode identification goes here.
	[xml]$feed=Invoke-Webrequest -URI $podcast.RSSFeed
	if ($(Invoke-Expression $podcast.EpNumberInTitle)) {
		# Find most recent downloaded episode using episode number
		if ((Gci -Path $Podcast.Directory -filter "*.mp3" ).count -gt 0) {
			[int]$mostrecent=((GCI -Path $podcast.directory -filter "*.mp3" | Sort-object -Property Name -Descending | Select -First 1).Name -split " - ")[0]
		} else {
			$mostrecent=$null
		}
		# Find all episodes matching the title filter which have a greater episode number than the most recent episode.
		$eps=@($feed.rss.channel.item | ? {($_.Title -match $podcast.TitleFilter) -and (((($_.Title -split " ")[0] -as [int]) -is [int]) -and ([int16]($_.Title -split " ")[0] -gt $mostrecent))} | Sort-Object -Property Title)
		if ($eps) {
			$var = New-Object PSObject
			Add-Member -InputObject $var -MemberType NoteProperty -Name Name -Value $podcast.Name
			Add-Member -InputObject $var -MemberType NoteProperty -Name Episodes -Value $eps
			Write-Host "Number of available new episodes:`t$($eps.Count)`n"
			$global:pendingepisodes+=$var
			Remove-Variable -name eps,var
		} else {
			Write-Host "Number of available new episodes:`t0`n"
		}
		Remove-Variable -name mostrecent -Force -erroraction silentlycontinue
	} else {
		# Find most recent downloaded episode using episode datestring
		if ((Gci -Path $Podcast.Directory -filter "*.mp3" ).count -gt 0) {
			[datetime]$mostrecent=((GCI -Path $podcast.directory -filter "*.mp3" | Sort-object -Property Name -Descending | Select -First 1).Name).SubString(0,10)
			$mostrecent=$mostrecent.AddMinutes(1439)
		} else {
			$mostrecent=$null
		}
		# Find all episodes matching the title filter which have a newer publication date than the most recent episode.
		$eps=@($feed.rss.channel.item | ? {($_.Title -match $podcast.TitleFilter) -and ([datetime]$($_.Pubdate -replace " EST","") -gt $mostrecent)}| Sort-Object @{Expression={[datetime]$($_.Pubdate -replace " EST","")};Ascending=$true})

		if ($eps) {
			if ($eps.GetType().BaseType.Name -eq "System.Xml.XmlElement") {
				$var = New-Object PSObject
				Add-Member -InputObject $var -MemberType NoteProperty -Name Name -Value $podcast.Name
				Add-Member -InputObject $var -MemberType NoteProperty -Name Episodes -Value $eps
				Write-Host "Number of available new episodes:`t1`n"
				$global:pendingepisodes+=$var
				Remove-Variable -name eps,var
			} else {
				$var = New-Object PSObject
				Add-Member -InputObject $var -MemberType NoteProperty -Name Name -Value $podcast.Name
				Add-Member -InputObject $var -MemberType NoteProperty -Name Episodes -Value $eps
				Write-Host "Number of available new episodes:`t$($eps.Count)`n"
				$global:pendingepisodes+=$var
				Remove-Variable -name eps,var
			}
		} else {
			Write-Host "Number of available new episodes:`t0`n"
		}
		Remove-Variable -name mostrecent,eps,var -Force -erroraction silentlycontinue
	}
}

Function Show-LatestEpisodes {
	cls
	Write-Host "Latest episodes released for managed podcasts`n" -foregroundcolor white
	for ($i=0;$i -lt $global:config.count;$i++) {
		$podcast=$global:config[$i];
		Write-Host "Podcast:`t$($podcast.Name)"
		[xml]$feed=Invoke-Webrequest -URI $podcast.RSSFeed
		$ep=($feed.rss.channel.item | ? {($_.Title -match $podcast.TitleFilter) -and ([datetime]$($_.Pubdate -replace " EST","") -gt $mostrecent)}| Sort-Object @{Expression={[datetime]$($_.Pubdate -replace " EST","")};Ascending=$false})[0]
		Write-Host "Latest episode:`t$($ep.Title)" -foregroundcolor white
		if ([datetime]$($ep.Pubdate -replace " EST","") -lt ((Get-Date).AddMonths(-6))) {
			$color="red"
		} elseif ([datetime]$($ep.Pubdate -replace " EST","") -lt ((Get-Date).AddMonths(-3))) {
			$color="yellow"
		} else {
			$color="green"
		}
		Write-Host "Released:`t$($ep.Pubdate)`n" -foregroundcolor $color
	}
	Start-Sleep 5
}

Function Download-PendingEpisodes {
	$invalidchars=@("/",":","*","?","<",">","|")
	# Only start the download pr
	if ($global:pendingepisodes) {
		foreach ($podcast in $global:config) {
			foreach ($episode in ($global:pendingepisodes | ? {$_.Name -eq $podcast.Name}).Episodes) {
				# Assemble filename from title, adding pubdate where necessary
				$title=$episode.Title
				if ($podcast.TitleTransform) {
					$targetfile=Invoke-Expression $podcast.TitleTransform
				} else {
					$targetfile=$title
				}
				# Check filename for invalid characters.
				foreach ($char in $invalidchars) {
					if ($targetfile -match $("\$char")) {
						$targetfile=$targetfile -replace $("\$char"),"_"
					}
				}
				# Special case for replacing quotes:
				if ($targetfile -match '"') {
					$targetfile=$targetfile -replace '"',"'"
				}

				if (!(Invoke-Expression $podcast.EpNumberInTitle)) {
					[datetime]$pubdate=$episode.Pubdate -replace " EST",""
					[string]$year=$pubdate.Year.ToString()
					if ($pubdate.Month.ToString().Length -lt 2) {
						[string]$month="0"+$pubdate.Month.ToString()
					} else {
						[string]$month=$pubdate.Month.ToString()
					}
					if ($pubdate.Day.ToString().Length -lt 2) {
						[string]$day="0"+$pubdate.Day.ToString()
					} else {
						[string]$day=$pubdate.Day.ToString()
					}
					if ($targetfile) {
						$targetfile=$year+"-"+$month+"-"+$day+" - "+$targetfile
					} else {
						$targetfile=$year+"-"+$month+"-"+$day
					}
				}
				
				$targetpath=$podcast.Directory+"\"+$targetfile+".mp3"
				$link=Invoke-Expression [string]$("`$episode."+$($podcast.Mp3Link))
				if ($podcast.Mp3Transform) {
					$link=Invoke-Expression ($podcast.Mp3Transform)
				}
				Start-BitsTransfer -Displayname "$($podcast.Name) Download" -Source $link -Destination $targetpath -Asynchronous

				Remove-Variable -Name targettitle,targetpath,link -Force -ErrorAction SilentlyContinue
			}

			# Optional progress reporting
			$transfers=Get-BitsTransfer
			while (($transfers | % {$_.JobState.ToString()} | ? {$_ -eq "Transferring"}) -or ($transfers | % {$_.JobState.ToString()} | ? {$_ -eq "Connecting"})) {
				# Progress stuff goes here
				$percentages=@()
				foreach ($transfer in $transfers) {
					$percentages+=[double]($transfer.BytesTransferred/$transfer.BytesTotal)*100
				}
				$percent="{0:N2}" -f ($percentages | Measure-Object  -Average).Average
				Write-Progress -Activity "Downloading pending files..." -CurrentOperation "$percent% complete"
			}
		}
		# Complete transfer process
		Complete-BitsTransfer -BitsJob $transfers
	}
}

Function Add-Podcast {
	[boolean]$valid=$false
	while (!$valid) {
		cls
		Write-Host "Add Managed Podcast`n" -foregroundcolor white
		$newfeed=Read-Host("Please enter the RSS feed address for the podcast")
		try {
			[xml]$feed=Invoke-Webrequest -URI $newfeed -ErrorAction Stop
			$eps=@()
			for ($i=0;$i -lt 10;$i++) {
				$eps+=($feed.rss.channel.item[$i])
			}
			[boolean]$valid=$true
		} catch {
			Write-Host "Could not retrieve feed from the specified address, please try again!"
			Start-Sleep 2
		}
	}
	Remove-variable -name valid -force -erroraction silentlycontinue
	
	# Check if name on feed should be used, prompt for name if not. - DONE
	cls
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	Write-Host "RSS feed retrieved:`nName:`t$($feed.rss.channel.title)"
	$name=Read-Host("If you wish to change the name used to manage this podcast, please enter it now. To use the displayed podcast name, press Enter");
	if (!$name) {
		$name=$($feed.rss.channel.title)
	}
	
	# Display 10 most recent episode titles and check if Episode number at start of title. - DONE
	[boolean]$validepnumber=$false
	while (!$validepnumber) {
		cls
		Write-Host "Add Managed Podcast`n" -foregroundcolor white
		foreach ($ep in $eps) {
			Write-Host "Episode title:`t$($ep.title)`n"
		}
		[string]$epchoice=Read-Host("Based on the displayed episode titles, is an integer episode number included at the start of every episode title? Y/N")
		switch -regex ($epchoice) {
			"Y" {
				$epnumber="`$true"
				$validepnumber=$true
				break;
			}
			"N" {
				$epnumber="`$false"
				$validepnumber=$true
				break;
			}
			default {
				Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
				start-sleep 3			
			}
		}
	}
	Remove-Variable -name epchoice,validepnumber -Force -ErrorAction Silentlycontinue
	
	# Check if f title filtering is required, prompt for regex to use and verify before proceeding. - DONE
	cls
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $eps) {
		Write-Host "Episode title:`t$($ep.title)`n"
	}
	[string]$titlefilter=Read-Host("Based on the displayed episode titles, is any filtering by episode titles required? Y/N")
	if ($titlefilter -eq "Y") {
		[boolean]$validfilter=$false
		while (!$validfilter) {
			$filter=Read-Host("Please enter a regular expression to match for filtering episode titles")
			if ($filter) {
				Write-Host "Filtered episode titles are:`n"
				foreach ($ep in ($eps | ? {$_.Title -match $filter})) {
					Write-Host "Episode title:`t$($ep.Title)`n"
				}
				[string]$confirm=Read-Host("Is this filter expression correct? Y/N")
				if ($confirm -eq "Y") {
					$validfilter=$true
				} else {
					Remove-variable -name filter -force -erroraction SilentlyContinue
				}
			} else {
				Write-Host "Filter cannot be empty, please try again!)" -foregroundcolor red
			}
		}
	}
	Remove-variable -name validfilter,confirm -force -erroraction silentlycontinue
	
	# 6) Check if title transform is required, prompt for regex to use and require confirmation before proceeding - DONE
	cls
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $($eps | ? {$_.Title -match $filter})) {
		Write-Host "Episode title:`t$($ep.title)`n"
	}
	[string]$titletransform=Read-Host("Based on the displayed episode titles, is any modification of episode titles required? Y/N")
	if ($titletransform -eq "Y") {
		[boolean]$validtransform=$false
		while (!$validtransform) {
			$transform=Read-Host("Please enter a regular expression to use for modifying episode titles, using `$title as the variable")
			if ($transform) {
				Write-Host "Transformed episode titles are:`n"
				foreach ($ep in ($eps | ? {$_.Title -match $filter})) {
					$title=$ep.Title
					$newtitle=Invoke-Expression $transform
					Write-Host "Episode title:`t$($newtitle)`n"
					Remove-Variable -Name title,newtitle -Force -ErrorAction SilentlyContinue
				}
				[string]$confirm=Read-Host("Is this modification expression correct? Y/N")
				if ($confirm -eq "Y") {
					$validtransform=$true
				} else {
					Remove-variable -name transform -force -erroraction SilentlyContinue
				}
			} else {
				Write-Host "Modification cannot be empty, please try again!)" -foregroundcolor red
			}
		}
	}
	Remove-variable -name validtransform,confirm -force -erroraction silentlycontinue
	
	# 7) Check if Mp3transform is required, prompt for regex to use and require confirmation before proceeding. - DONE
	cls
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $($eps | ? {$_.Title -match $filter})) {
		Write-Host "Link:`t$($ep.enclosure.url)`n"
	}
	[string]$mp3transform=Read-Host("Based on the displayed download links, is any modification of URLs required? Y/N")
	if ($mp3transform -eq "Y") {
		cls
		[boolean]$validmp3=$false
		while (!$validmp3) {
			$mp3=Read-Host("Please enter a regular expression to use for modifying download links, using `$link as the variable")
			Write-Host "Transformed download links are:`n"
			foreach ($ep in ($eps | ? {$_.Title -match $filter})) {
				$link=$ep.enclosure.url
				$newurl=Invoke-Expression $mp3
				Write-Host "Modified link:`t$($newurl)`n"
				Remove-Variable -Name link,newlink -Force -ErrorAction SilentlyContinue
			}
			[string]$confirm=Read-Host("Is this modification expression correct? Y/N")
			if ($confirm -eq "Y") {
				Write-Host "Testing modified link..."
				try {
					$link=$eps[0].Enclosure.url
					$URI=Invoke-Expression $($mp3) 
					$test=Invoke-WebRequest -URI $URI -ErrorAction Stop
					if (($test.BaseResponse.StatusCode -eq "OK") -and ($test.BaseResponse.ContentType -eq "audio/mpeg")) {
						Write-Host "Modified link test successful." -foregroundcolor green
						Start-Sleep 3
						$validmp3=$true
					} else {
						Write-Host "Modified link test unsuccessful.`nStatus code:`t$($test.baseresponse.statuscode)`nContent Type:`t$($test.baseresponse.contenttype)`nContent Size:`t$('{0:N2}' -f $($test.BaseResponse.ContentLength/1mb))" -foregroundcolor red
						Start-Sleep 3
					}
				} catch {
					Write-Host "An error occured while trying to test the modified download link, message was:`n$($_.Exception.Message)" -foregroundcolor red
					Start-Sleep 3
				}
			} else {
				Remove-variable -name mp3 -force -erroraction SilentlyContinue
			}
		}
	}
	Remove-variable -name validmp3,confirm -force -erroraction silentlycontinue
	
	# 8) Check if F:\Podcasts\<name> exists, then check if it should be created or another path used. Create directory once confirmed.
	cls
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	[string]$fullpath=$podcastdir+$($name)
	[boolean]$validpath=$false
	While (!$validpath) {
		If (Test-Path $fullpath -ErrorAction SilentlyContinue) {
			Write-Host "Directory $($fullpath) already exists, no action needed."
			$validpath=$true
		} else {
			Write-Host "Directory $($fullpath) does not exist, creating..."
			try {
				New-Item -ItemType Directory -Path $fullpath -ErrorAction Stop
				Write-Host "Directory $($fullpath) created successfully." -foregroundcolor green
				$validpath=$true
			} catch {
				Write-Host "Directory $($fullpath) couldn't be created, error was:`n$($_.Exception.Message)." -foregroundcolor red
				$newname=Read-Host("Please enter another directory name to use for this podcast")
				[string]$fullpath=$podcastdir+$newname
			}
		}
	}
	
	# 9) Create object, add to config, write new contents of config file. DONE
	$var = New-Object PSObject

	Add-Member -InputObject $var -MemberType NoteProperty -Name Name -Value $name
	Add-Member -InputObject $var -MemberType NoteProperty -Name RSSFeed -Value $newfeed
	Add-Member -InputObject $var -MemberType NoteProperty -Name EpNumberInTitle -Value $epnumber
	Add-Member -InputObject $var -MemberType NoteProperty -Name TitleFilter -Value $filter
	if ($titletransform -eq "Y") {
		Add-Member -InputObject $var -MemberType NoteProperty -Name TitleTransform -Value $transform
	} else {
		Add-Member -InputObject $var -MemberType NoteProperty -Name TitleTransform -Value $null
	}	
	Add-Member -InputObject $var -MemberType NoteProperty -Name Mp3Link -Value "enclosure.url"
	if ($mp3transform -eq "Y") {
		Add-Member -InputObject $var -MemberType NoteProperty -Name Mp3Transform -Value $mp3
	} else {
		Add-Member -InputObject $var -MemberType NoteProperty -Name Mp3Transform -Value $null	
	}
	Add-Member -InputObject $var -MemberType NoteProperty -Name Directory -Value $fullpath
	
	# Add new object to $global:config array
	$global:config+=$var
	
	# Remove unnecessary parameters
	Remove-Variable -Name var,newfeed,titletransform,transform,mp3transform,fullpath -ErrorAction SilentlyContinue
	
	# Check if existing old config file exists, if so delete, then back up original config.txt file
	$oldpath=$PsScriptRoot+"\Config.old"
	If (Test-Path $oldpath) {
		Remove-Item -Path $oldpath -Force -Confirm:$false
		remove-variable -name oldpath -force -erroraction silentlycontinue
	}
	Rename-Item -Path $($PSScriptRoot+"\Config.txt") -NewName "Config.old" -Force -confirm:$false
	
	# Output updated $global:config array to text file
	$global:config | Export-CSV -Path $($PSScriptRoot+"\Config.txt") -Encoding UTF8 -Delimiter ";" -NoTypeInformation
	
	# Reload config file
	[System.Collections.ArrayList]$global:config=Import-CSV -Path $($PSScriptRoot+"\Config.txt") -Delimiter ";"

	# 10) Check for episodes to download: all available or only most recent. Default is most recent.
	[string]$epchoice=Read-Host("To download only the latest episode, press Enter. To download all episodes, type 'ALL'")
	if ($epchoice -eq "ALL") {
		Write-Host "Retrieving all available episodes..."
		Get-PendingEpisodes -podchoice $($global:config.count -1)
		Download-PendingEpisodes
	} else {
		Write-Host "Retrieving the most recent episode..."
		$podcast=$global:config[$($global:config.count -1)]
		[xml]$feed=Invoke-Webrequest -URI $podcast.RSSFeed
		if ($(Invoke-Expression $podcast.EpNumberInTitle)) {
			$eps=($feed.rss.channel.item | ? {($_.Title -match $podcast.TitleFilter) -and ((($_.Title -split " ")[0] -as [int]) -is [int])})[0]
		} else {
			$eps=@($feed.rss.channel.item | ? {$_.Title -match $podcast.TitleFilter})[0]
		}
		$var = New-Object PSObject
		Add-Member -InputObject $var -MemberType NoteProperty -Name Name -Value $podcast.Name
		Add-Member -InputObject $var -MemberType NoteProperty -Name Episodes -Value $eps
		Write-Host "Number of available new episodes:`t1`n"
		$global:pendingepisodes+=$var
		Remove-Variable -name eps,var
		Download-PendingEpisodes
	}
}

Function View-Podcasts {
	cls
	Write-Host "The following podcasts are currently managed:`n"
	for ($i=0;$i -lt $global:config.count;$i++) {
		$podcast=$global:config[$i];
		Write-Host "Podcast:`t$($podcast.Name)"
		if (Test-Path $podcast.Directory) {
			Write-Host "Podcast directory exists." -foregroundcolor green
		} else {
			Write-Host "Podcast directory does not exist!" -foregroundcolor red
		}
		try {
			[xml]$feed=Invoke-Webrequest -URI $podcast.RSSFeed -ErrorAction Stop;
			Write-Host "Podcast RSS feed is online." -foregroundcolor green
		} catch {
			Write-Host "An error occured retrieving the podcast RSS feed, message was:`n$($_.Exception.Message).`nIf this occurs repeatedly, the RSS feed may not be valid." -foregroundcolor red
		}
		Write-Host "`n"
	}
	Write-Host ""
	Read-Host("Press Enter to return to the main menu")
}

Function Remove-Podcast {
	# Display names of podcast, prompt for numerical selection, remove from config array and update config.txt file.
	[boolean]$valid=$false
	while (!$valid) {
		cls
		Write-Host "Add new podcast to managed podcast list`n"
		Write-Host "The following podcasts are managed:`n"
		for ($i=0;$i -lt $global:config.count;$i++) {
			Write-Host "$($i+1): $($global:config[$i].Name)"
		}
		Write-Host ""
		[int]$choice=Read-Host("Please select the podcast to remove")
		$choice--
		if ($global:config[$choice]) {
			$global:config.RemoveAt($choice)
			$valid=$true;
		} else {
			Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
			Start-sleep 3
			cls
		}
	}
	# Back up original config.txt file
	
	If (Test-Path $($PsScriptRoot+"\Config.old")) {
		Remove-Item -Path $($PsScriptRoot+"\Config.old") -Force -Confirm:$false
	}
	Rename-Item -Path $($PSScriptRoot+"\Config.txt") -NewName "Config.old" -Force -confirm:$false
	
	# Output updated $global:config array to text file
	$global:config | Export-CSV -Path $($PSScriptRoot+"\Config.txt") -Encoding UTF8 -Delimiter ";" -NoTypeInformation
	Write-Host "Managed podcast list has been updated. The previous configuration has been saved at $($PSScriptRoot+`"\Config.old`")."
	
	# Reload original config file
	[System.Collections.ArrayList]$global:config=Import-CSV -Path $($PSScriptRoot+"\Config.txt") -Delimiter ";"
	
	Remove-Variable -Name valid,choice -Force -ErrorAction SilentlyContinue
	Start-Sleep 3
}

Function Update-PlayerFiles {
	<#
	Logic for this function will be:
	Select podcast(s) to update on player
	for each selected podcast:
		if player folder exists
			Check for most recent ep on player
			Copy all episodes more recent than that
		else 
			Create folder on player
			Prompt user for how many episodes to copy
		
	#>
}
	

###############
# Main Script
###############

# Enable TLS 1.2 support
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;

if (!$menu) {
	# Unattended Mode
	cls
	Write-Host "Podcast Manager - Unattended Mode`n"
	if ($global:config) {
		# Iterate through the config array and call Get-PendingEpisodes to identify available downloads
		for ($i=0;$i -lt $global:config.count;$i++) {
			Get-PendingEpisodes -podchoice $i
		}
		# Download available episodes
		Download-PendingEpisodes
		Write-Host "All operations complete, terminating."
	} else {
		Write-Host "Halting as configuration could not be loaded." -foregroundcolor red
	}
} else {
	# Interactive Mode
	cls
	Write-host "Podcast Manager`n" -foregroundcolor white
	# Check config file has loaded, start menu
	if ($global:config) {
		Show-Menu
	} else {
		Write-Host "Halting as configuration could not be loaded." -foregroundcolor red
	}
}