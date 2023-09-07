#Podcast-Manager
param(
	[Parameter(Mandatory=$false)]
	[string]$config
)
Import-Module BitsTransfer

# Configuration import
if (-not $config) {
	if ($PSScriptRoot -ne "") {
		$config = "$($PSScriptRoot)\Config.xml"
	} else {
		$config = ".\Config.xml"
	}
}

if (Test-path $config) {
	[xml]$settings = Get-Content $config
} else {
	Write-Output "No configuration file found at $($config). Starting configuration process..."
	Start-sleep 2
	Set-ConfigFile
}

if ($PSScriptRoot -ne "") {
	[System.Collections.ArrayList]$global:podcasts = Import-CSV -Path $($PSScriptRoot+"$($settings.Settings.PodcastList)") -Delimiter $settings.Settings.PodcastDelimiter;
	[System.Collections.ArrayList]$invalidchars = Import-CSV -Path $($PSScriptRoot+"$($settings.Settings.InvalidChars)")
} else {
	[System.Collections.ArrayList]$global:podcasts = Import-CSV -Path $settings.Settings.PodcastList -Delimiter $settings.Settings.PodcastDelimiter;
	[System.Collections.ArrayList]$invalidchars = Import-CSV -Path $settings.Settings.InvalidChars
}
[System.collections.arraylist]$global:pendingepisodes = @()

# Local and remote directory configurations
$podcastdir = $settings.Settings.PodcastDirectory
$playerpath = $settings.Settings.PlayerDirectory

# Interactive mode
if ($settings.Settings.Interactive -eq "Enabled") {
	[boolean]$interactive = $true
} 

###############
# Functions
###############


function Show-Menu {
	# Define the available functions within the menu
	$global:menu = @"
Podcast Manager

1.	Check for available episodes of all podcasts 
2.	Check for available episodes of a specific podcast
3.	Show information for most recent available episodes of all podcasts
4.	Download available episodes of current podcasts
5.	View managed podcasts
6.	Add new managed podcast
7.	Remove managed podcast
8.	Update player files
9.	Quit

"@
	# Wrap menu in a while-loop
	$quit = $false
	do {
		#Clear-Host
		Write-Host $menu
		$choice = Read-Host("Please choose from 1-9")

		switch ($choice) {
			"1" {
				for ($i = 0;$i -lt $global:podcasts.count;$i++) {
					Get-PendingEpisodes -podchoice $i
				}
				Start-sleep 5
			}
			"2" {
				# Display names of podcast, prompt for numerical selection, then invoke Get-PendingEpisodes.
				Clear-Host
				[boolean]$valid = $false
				while (!$valid) {
					Write-Host "The following podcasts are managed:`n"
					for ($i = 0;$i -lt $global:podcasts.count;$i++) {
						Write-Host "$($i+1): $($global:podcasts[$i].Name)"
					}
					Write-Host ""
					[int]$choice = Read-Host("Please select the desired podcast")
					$choice--
					if ($global:podcasts[$choice]) {
						$valid = $true; 
						Get-PendingEpisodes -podchoice $choice
					} else {
						Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
						Start-sleep 3
						Clear-Host
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
				Update-PlayerFiles;
			}
			"9" {
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

Function Set-ConfigFile {
	Clear-Host
	$prompt = Read-Host -Prompt "Is interactive mode required? Y/N"
	if	($prompt -eq "Y") {
		$interactive = "Enabled"
	} else {
		$interactive = "Disabled"
	}
	$podcastlist = ".\podcasts.csv"
	if (-not Test-path $podcastlist) {
		New-Item -Name 'podcasts.csv' -Path '.' -Type File
	}
	$delimiter = ";"
	$invalid = ".\invalidchars.csv"
	[boolean]$valid = $false
	while (!$valid) {
		$prompt = Read-Host -Prompt "Enter full path for top-level folder where podcasts will be stored"
		if	(Test-Path $prompt) {
			$podpath = $prompt
			$valid = $true
		} else {
			Write-Output "Invalid path specified, please try again"
		}
	}
	$prompt = Read-Host -Prompt "Enter full path to folder on player where episodes will be copied. Press Enter to skip"
	if ($prompt -ne "") {
		$playerpath = $prompt
	}
	$config = @"
	<?xml version = "1.0" encoding = "UTF-8"?>
	<Settings>
		<Interactive>$($interactive)</Interactive>
		<PodcastList>$($podcastlist)</PodcastList>
		<PodcastDelimiter>$($delimiter)</PodcastDelimiter>
		<InvalidChars>$($invalid)</InvalidChars>
		<PodcastDirectory>$($podpath)</PodcastDirectory>
		<PlayerDirectory>$($playerpath)</PlayerDirectory>
	</Settings>
"@
	Out-file -FilePath .\Config.xml -Encoding UTF8 -InputObject $config
}

Function Get-PendingEpisodes {
	param(
		[Parameter(Mandatory = $true)][string]$podchoice
	)
	$podcast = $global:podcasts[$podchoice]
	Write-Host "Podcast:`t$($podcast.Name)"
	# Feed retrieval and new episode identification goes here.
	[xml]$feed = Invoke-Webrequest -URI $podcast.RSSFeed
	if ($(Invoke-Expression $podcast.EpNumber)) {
		# If cutoff is populated, use that
		if ($podcast.cutoff) {
			[int]$mostrecent = $podcast.cutoff
		} elseif ((Get-ChildItem -Path "$($podcastdir)\$($Podcast.Directory)" -filter "*.mp3").count -gt 0) {
			[int]$mostrecent = ((Get-ChildItem -Path "$($podcastdir)\$($Podcast.Directory)" -filter "*.mp3" | Sort-object -Property Name -Descending | Select-Object -First 1).Name -split " - |:")[0]
		} else {
			$mostrecent = $null
		}
		# Check available episodes that match the episode filter until one exceeds the cutoff.
		
		# OLD CODE $eps = @($feed.rss.channel.item | Where-Object {($_.Title -match $podcast.TitleFilter) -and (((($_.Title -split " |:")[0] -as [int]) -is [int]) -and ([int16]($_.Title -split " |:")[0] -gt $mostrecent))} | Sort-Object -Property Title)
		[System.collections.arraylist]$eps = @()
		for ($i = 0; $i -lt $feed.rss.channel.item.count; $i++) {
			$episode = $feed.rss.channel.item[$i] | Where-Object{Invoke-Expression $podcast.EpFilter} |Select-Object enclosure,@{N = "EpisodeNumber";E = {[int](Invoke-Expression $podcast.EpNumberTransform)}} | Where-Object {$_.EpisodeNumber -gt $podcast.cutoff}
			if ($episode) {
				$eps.Add($episode) | out-Null
			} else {
				break;
			}
		}
		
		if ($eps) {
			$var = New-Object -Type PSObject @{
				Name = $podcast.Name
				Episodes = $eps
			}
			Write-Host "Number of available new episodes:`t$($eps.Count)`n"
			$global:pendingepisodes.Add($var) | Out-Null
			Remove-Variable -name eps,var
		} else {
			Write-Host "Number of available new episodes:`t0`n"
		}
		Remove-Variable -name mostrecent -Force -erroraction silentlycontinue
	} else {
		# Find most recent downloaded episode using episode datestring
		if ((Get-ChildItem -path "$($podcastdir)\$($Podcast.Directory)" -filter "*.mp3" ).count -gt 0) {
			[datetime]$mostrecent = ((Get-ChildItem -path "$($podcastdir)\$($Podcast.Directory)" -filter "*.mp3" | Sort-object -Property Name -Descending | Select-Object -First 1).Name).SubString(0,10)
			$mostrecent = $mostrecent.AddMinutes(1439)
		} else {
			$mostrecent = $null
		}
		# Find all episodes matching the title filter which have a newer publication date than the most recent episode.
		[system.collections.arraylist]$eps = @($feed.rss.channel.item | Where-Object {($_.Title -match $podcast.TitleFilter) -and ([datetime]$($_.Pubdate -replace " EST","") -gt $mostrecent)}| Sort-Object @{Expression = {[datetime]$($_.Pubdate -replace " EST","")};Ascending = $true})

		if ($eps) {
			$var = New-Object -Type PSObject @{
				Name = $podcast.Name
				Episodes = $eps
			}
			$global:pendingepisodes.Add($var) | Out-Null
			if ($eps.GetType().BaseType.Name -eq "System.Xml.XmlElement") {
				Write-Host "Number of available new episodes:`t1`n"
			} else {
				Write-Host "Number of available new episodes:`t$($eps.Count)`n"
			}
		} else {
			Write-Host "Number of available new episodes:`t0`n"
		}
		Remove-Variable -name mostrecent,eps,var -Force -erroraction silentlycontinue
	}
}

Function Show-LatestEpisodes {
	Clear-Host
	Write-Host "Latest episodes released for managed podcasts`n" -foregroundcolor white
	for ($i = 0;$i -lt $global:podcasts.count;$i++) {
		$podcast = $global:podcasts[$i];
		Write-Host "Podcast:`t$($podcast.Name)"
		[xml]$feed = Invoke-Webrequest -URI $podcast.RSSFeed
		$ep = ($feed.rss.channel.item | Where-Object {($_.Title -match $podcast.TitleFilter) -and ([datetime]$($_.Pubdate -replace " EST","") -gt $mostrecent)}| Sort-Object @{Expression = {[datetime]$($_.Pubdate -replace " EST","")};Ascending = $false})[0]
		Write-Host "Latest episode:`t$($ep.Title)" -foregroundcolor white
		if ([datetime]$($ep.Pubdate -replace " EST","") -lt ((Get-Date).AddMonths(-6))) {
			$color = "red"
		} elseif ([datetime]$($ep.Pubdate -replace " EST","") -lt ((Get-Date).AddMonths(-3))) {
			$color = "yellow"
		} else {
			$color = "green"
		}
		Write-Host "Released:`t$($ep.Pubdate)`n" -foregroundcolor $color
	}
	Start-Sleep 5
}

Function Download-PendingEpisodes {
	# Only start the download pr
	if ($global:pendingepisodes) {
		foreach ($podcast in $global:podcasts) {
			foreach ($episode in ($global:pendingepisodes | Where-Object {$_.Name -eq $podcast.Name}).Episodes) {
				# Assemble filename from title, adding pubdate where necessary
				$title = $episode.Title
				if ($podcast.TitleTransform) {
					$targetfile = Invoke-Expression $podcast.TitleTransform
				} else {
					$targetfile = $title
				}
				# Check filename for invalid characters.
				foreach ($char in $invalidchars) {
					if ($char.Replace -eq "") {
						$replace = "_"
					} else {
						$replace = $char.Replace
					}
					if ($targetfile -match $char.Escaped) {
						$targetfile = $targetfile -replace $($char.Escaped),$replace
					}
					Remove-Variable -name replace -Force -ErrorAction SilentlyContinue
				}
				# Special case for replacing quotes:
				if ($targetfile -match '"') {
					$targetfile = $targetfile -replace '"',"'"
				}
				
				if (!(Invoke-Expression $podcast.EpNumberInTitle)) {
					[datetime]$pubdate = $episode.Pubdate -replace " EST",""
					[string]$year = $pubdate.Year.ToString()
					if ($pubdate.Month.ToString().Length -lt 2) {
						[string]$month = "0"+$pubdate.Month.ToString()
					} else {
						[string]$month = $pubdate.Month.ToString()
					}
					if ($pubdate.Day.ToString().Length -lt 2) {
						[string]$day = "0"+$pubdate.Day.ToString()
					} else {
						[string]$day = $pubdate.Day.ToString()
					}
					if ($targetfile) {
						$targetfile = $year+"-"+$month+"-"+$day+" - "+$targetfile
					} else {
						$targetfile = $year+"-"+$month+"-"+$day
					}
				}
				
				$targetpath = $podcast.Directory+"\"+$targetfile+".mp3"
				$link = Invoke-Expression [string]$("`$episode."+$($podcast.Mp3Link))
				if ($podcast.Mp3Transform) {
					$link = Invoke-Expression ($podcast.Mp3Transform)
				}
				Start-BitsTransfer -Displayname "$($podcast.Name) Download" -Source $link -Destination $targetpath -Asynchronous

				Remove-Variable -Name targettitle,targetpath,link -Force -ErrorAction SilentlyContinue
			}

			# Optional progress reporting
			$transfers = Get-BitsTransfer
			while (($transfers | Foreach-Object {$_.JobState.ToString()} | Where-Object {$_ -eq "Transferring"}) -or ($transfers | Foreach-Object {$_.JobState.ToString()} | Where-Object {$_ -eq "Connecting"})) {
				# Progress stuff goes here
				[System.Collections.Arraylist]$percentages = @()
				foreach ($transfer in $transfers) {
					$percentages.Add([double]($transfer.BytesTransferred/$transfer.BytesTotal)*100) | Out-Null
				}
				$percent = "{0:N2}" -f ($percentages | Measure-Object  -Average).Average
				Write-Progress -Activity "Downloading pending files..." -CurrentOperation "$percent% complete"
			}
		}
		# Complete transfer process
		Complete-BitsTransfer -BitsJob $transfers
	}
}

Function Add-Podcast {
	[boolean]$valid = $false
	while (!$valid) {
		Clear-Host
		Write-Host "Add Managed Podcast`n" -foregroundcolor white
		$newfeed = Read-Host("Please enter the RSS feed address for the podcast")
		try {
			[xml]$feed = Invoke-Webrequest -URI $newfeed -ErrorAction Stop
			[System.Collections.ArrayList]$eps = @()
			for ($i = 0;$i -lt 10;$i++) {
				$eps.Add($feed.rss.channel.item[$i]) | Out-Null
			}
			[boolean]$valid = $true
		} catch {
			Write-Host "Could not retrieve feed from the specified address, please try again!"
			Start-Sleep 2
		}
	}
	Remove-variable -name valid -force -erroraction silentlycontinue
	
	# Check if name on feed should be used, prompt for name if not. - DONE
	Clear-Host
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	Write-Host "RSS feed retrieved:`nName:`t$($feed.rss.channel.title)"
	$name = Read-Host("If you wish to change the name used to manage this podcast, please enter it now. To use the displayed podcast name, press Enter");
	if (!$name) {
		$name = $($feed.rss.channel.title)
	}
	
	# Display 10 most recent episode titles and check if Episode number at start of title. - DONE
	[boolean]$validepnumber = $false
	while (!$validepnumber) {
		Clear-Host
		Write-Host "Add Managed Podcast`n" -foregroundcolor white
		foreach ($ep in $eps) {
			Write-Host "Episode title:`t$($ep.title)`n"
		}
		[string]$epchoice = Read-Host("Based on the displayed episode titles, is an integer episode number included at the start of every episode title? Y/N")
		switch -regex ($epchoice) {
			"Y" {
				$epnumber = "`$true"
				$validepnumber = $true
				break;
			}
			"N" {
				$epnumber = "`$false"
				$validepnumber = $true
				break;
			}
			default {
				Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
				start-sleep 3			
			}
		}
	}
	Remove-Variable -name epchoice,validepnumber -Force -ErrorAction Silentlycontinue
	
	# Check if title filtering is required, prompt for regex to use and verify before proceeding. - DONE
	Clear-Host
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $eps) {
		Write-Host "Episode title:`t$($ep.title)`n"
	}
	[string]$titlefilter = Read-Host("Based on the displayed episode titles, is any filtering by episode titles required? Y/N")
	if ($titlefilter -eq "Y") {
		[boolean]$validfilter = $false
		while (!$validfilter) {
			$filter = Read-Host("Please enter a regular expression to match for filtering episode titles")
			if ($filter) {
				Write-Host "Filtered episode titles are:`n"
				foreach ($ep in ($eps | Where-Object {$_.Title -match $filter})) {
					Write-Host "Episode title:`t$($ep.Title)`n"
				}
				[string]$confirm = Read-Host("Is this filter expression correct? Y/N")
				if ($confirm -eq "Y") {
					$validfilter = $true
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
	Clear-Host
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $($eps | Where-Object {$_.Title -match $filter})) {
		Write-Host "Episode title:`t$($ep.title)`n"
	}
	[string]$titletransform = Read-Host("Based on the displayed episode titles, is any modification of episode titles required? Y/N")
	if ($titletransform -eq "Y") {
		[boolean]$validtransform = $false
		while (!$validtransform) {
			$transform = Read-Host("Please enter a regular expression to use for modifying episode titles, using `$title as the variable")
			if ($transform) {
				Write-Host "Transformed episode titles are:`n"
				foreach ($ep in ($eps | Where-Object {$_.Title -match $filter})) {
					$title = $ep.Title
					$newtitle = Invoke-Expression $transform
					Write-Host "Episode title:`t$($newtitle)`n"
					Remove-Variable -Name title,newtitle -Force -ErrorAction SilentlyContinue
				}
				[string]$confirm = Read-Host("Is this modification expression correct? Y/N")
				if ($confirm -eq "Y") {
					$validtransform = $true
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
	Clear-Host
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	foreach ($ep in $($eps | Where-Object {$_.Title -match $filter})) {
		Write-Host "Link:`t$($ep.enclosure.url)`n"
	}
	[string]$mp3transform = Read-Host("Based on the displayed download links, is any modification of URLs required? Y/N")
	if ($mp3transform -eq "Y") {
		Clear-Host
		[boolean]$validmp3 = $false
		while (!$validmp3) {
			$mp3 = Read-Host("Please enter a regular expression to use for modifying download links, using `$link as the variable")
			Write-Host "Transformed download links are:`n"
			foreach ($ep in ($eps | Where-Object {$_.Title -match $filter})) {
				$link = $ep.enclosure.url
				$newurl = Invoke-Expression $mp3
				Write-Host "Modified link:`t$($newurl)`n"
				Remove-Variable -Name link,newlink -Force -ErrorAction SilentlyContinue
			}
			[string]$confirm = Read-Host("Is this modification expression correct? Y/N")
			if ($confirm -eq "Y") {
				Write-Host "Testing modified link..."
				try {
					$link = $eps[0].Enclosure.url
					$URI = Invoke-Expression $($mp3) 
					$test = Invoke-WebRequest -URI $URI -ErrorAction Stop
					if (($test.BaseResponse.StatusCode -eq "OK") -and ($test.BaseResponse.ContentType -eq "audio/mpeg")) {
						Write-Host "Modified link test successful." -foregroundcolor green
						Start-Sleep 3
						$validmp3 = $true
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
	
	# 8) Check if target directory exists, then check if it should be created or another path used. Create directory once confirmed.
	Clear-Host
	Write-Host "Add Managed Podcast`n" -foregroundcolor white
	[string]$fullpath = $podcastdir+"\"+$($name)
	[boolean]$validpath = $false
	While (!$validpath) {
		If (Test-Path $fullpath -ErrorAction SilentlyContinue) {
			Write-Host "Directory $($fullpath) already exists, no action needed."
			$validpath = $true
		} else {
			Write-Host "Directory $($fullpath) does not exist, creating..."
			try {
				New-Item -ItemType Directory -Path $fullpath -ErrorAction Stop
				Write-Host "Directory $($fullpath) created successfully." -foregroundcolor green
				$validpath = $true
			} catch {
				Write-Host "Directory $($fullpath) couldn't be created, error was:`n$($_.Exception.Message)." -foregroundcolor red
				$newname = Read-Host("Please enter another directory name to use for this podcast")
				[string]$fullpath = $podcastdir+"\"+$newname
			}
		}
	}
	
	# 9) Create object, add to config, write new contents of config file. DONE
	if ($titletransform -ne "Y") {
		$transform = $null
	}	
	if ($mp3transform -ne "Y") {
		$mp3 = $null	
	}
	$var = New-Object -Type PSObject @{
		Name = $name
		RSSFeed = $newfeed
		EpNumberInTitle = $epnumber
		TitleFilter = $filter
		TitleTransform = $transform	
		Mp3Link = "enclosure.url"
		Mp3Transform = $mp3
		Directory = $fullpath
	}
	# Add new object to $global:podcasts array
	$global:podcasts.Add($var) | Out-Null
	
	# Remove unnecessary parameters
	Remove-Variable -Name var,newfeed,titletransform,transform,mp3transform,fullpath -ErrorAction SilentlyContinue
	
	# Rename original podcast list file, then save new list under original filename
	if ($PSScriptRoot -ne "") {
		$podlist = $($PSScriptRoot+"$($settings.Settings.PodcastList)");
		$oldlist = $($PSScriptRoot+"$($settings.Settings.PodcastList)").Replace(".csv","")+"_"+(Get-Date -format 'yyyyMMdd_HHmm').ToString()+".old"
	} else {
		$podlist = $settings.Settings.PodcastList;
		$oldlist = ($settings.Settings.PodcastList).Replace(".csv","")+"_"+(Get-Date -format 'yyyyMMdd_HHmm').ToString()+".old"
	}
	
	Rename-Item -Path $podlist -NewName $oldlist -Force -confirm:$false
	$global:podcasts | Export-CSV -Path $podlist -Encoding UTF8 -Delimiter $settings.Settings.PodcastDelimiter -NoTypeInformation
	[System.Collections.ArrayList]$global:podcasts = Import-CSV -Path $podlist -Delimiter $settings.Settings.PodcastDelimiter
	Remove-Variable -name podlist,oldlist -Force ErrorAction SilentlyContinue
	
	# 10) Check for episodes to download: all available or only most recent. Default is most recent.
	[string]$epchoice = Read-Host("To download only the latest episode, press Enter. To download all episodes, type 'ALL'")
	if ($epchoice -eq "ALL") {
		Write-Host "Retrieving all available episodes..."
		Get-PendingEpisodes -podchoice $($global:podcasts.count -1)
		Download-PendingEpisodes
	} else {
		Write-Host "Retrieving the most recent episode..."
		$podcast = $global:podcasts[$($global:podcasts.count -1)]
		[xml]$feed = Invoke-Webrequest -URI $podcast.RSSFeed
		$eps = ($feed.rss.channel.item | Where-Object {$_.Title -match $podcast.TitleFilter})[0]
		$var = New-Object PSObject -@{
			Name = $podcast.Name
			Episodes = $eps
		}
		Write-Host "Number of available new episodes:`t1`n"
		$global:pendingepisodes.Add($var) | Out-Null
		Remove-Variable -name eps,var
		Download-PendingEpisodes
	}
}

Function View-Podcasts {
	Clear-Host
	Write-Host "The following podcasts are currently managed:`n"
	for ($i = 0;$i -lt $global:podcasts.count;$i++) {
		$podcast = $global:podcasts[$i];
		Write-Host "Podcast:`t$($podcast.Name)"
		if (Test-path "$($Podcast.Directory)") {
			Write-Host "Podcast directory exists." -foregroundcolor green
		} else {
			Write-Host "Podcast directory does not exist!" -foregroundcolor red
		}
		try {
			[xml]$feed = Invoke-Webrequest -URI $podcast.RSSFeed -ErrorAction Stop;
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
	[boolean]$valid = $false
	while (!$valid) {
		Clear-Host
		Write-Host "Add new podcast to managed podcast list`n"
		Write-Host "The following podcasts are managed:`n"
		for ($i = 0;$i -lt $global:podcasts.count;$i++) {
			Write-Host "$($i+1): $($global:podcasts[$i].Name)"
		}
		Write-Host ""
		[int]$choice = Read-Host("Please select the podcast to remove")
		$choice--
		if ($global:podcasts[$choice]) {
			$global:podcasts.RemoveAt($choice)
			$valid = $true;
		} else {
			Write-Host "Invalid selection entered, please try again!" -foregroundcolor red
			Start-sleep 3
			Clear-Host
		}
	}
	# Rename original podcast list file, then save new list under original filename
	if ($PSScriptRoot -ne "") {
		$podlist = $($PSScriptRoot+"$($settings.Settings.PodcastList)");
		$oldlist = $($PSScriptRoot+"$($settings.Settings.PodcastList)").Replace(".csv","")+"_"+(Get-Date -format 'yyyyMMdd_HHmm').ToString()+".old"
	} else {
		$podlist = $settings.Settings.PodcastList;
		$oldlist = ($settings.Settings.PodcastList).Replace(".csv","")+"_"+(Get-Date -format 'yyyyMMdd_HHmm').ToString()+".old"
	}
	
	Rename-Item -Path $podlist -NewName $oldlist -Force -confirm:$false
	$global:podcasts | Export-CSV -Path $podlist -Encoding UTF8 -Delimiter $settings.Settings.PodcastDelimiter -NoTypeInformation
	[System.Collections.ArrayList]$global:podcasts = Import-CSV -Path $podlist -Delimiter $settings.Settings.PodcastDelimiter
	Remove-Variable -name valid,choice,podlist,oldlist -Force ErrorAction SilentlyContinue
	Start-Sleep 3
}

Function Update-PlayerFiles {
	param(
		[Parameter(Mandatory = $false)][string]$mode = "Current"
	)
	switch ($mode) {
		"Current" {
			<# 	Current:
			Check player directory
			for each directory found, check if a podcast in $podcast matches the name
			If found:
				Check for latest episode on player, either by episode or date.
				Identify newer episodes in podcast directory
				Copy sequentially to player.
			#>
			if (!$playerpath) {
					$playerpath = Read-Host -Prompt "Enter full path of player podcast directory"
			}
			if (Test-path $playerpath) {
				Write-Host "Updating player with latest episodes..."
				$podmatch = ($podcastdir -replace "\\","\\") -replace ":","\:"
				foreach ($podcast in $global:podcasts) {
				
					$playerdir = $podcast.directory -replace $podmatch,$playerpath
					# Check if the podcast directory exists on the player
					if (Test-Path $playerdir) {
						Write-host "$($podcast.Name)"
						# Check for latest episode depending on whether podcast uses episode numbers or dates.
						if (Invoke-Expression ($podcast.EpNumberInTitle)) {
							[int]$mostrecent = ((Get-ChildItem -Path $playerdir -filter "*.mp3" | Sort-object -Property Name -Descending | Select-Object -First 1).Name -split " - ")[0]
							$files = Get-ChildItem $podcast.Directory -Filter "*.mp3" | Where-Object {[int](($_.Name -split " - ")[0]) -gt $mostrecent}
							if ($files.count -gt 0) {
								Write-host "New episodes of $($podcast.Name) found, starting copy process..."
								foreach ($file in $files) {
									Copy-Item -Path $file.Fullname -Destination $playerdir
								}
							} else {
								Write-host "No new episodes available."
							}
							remove-variable -name mostrecent,files -Force -Erroraction SilentlyContinue	
						} else {
							[datetime]$mostrecent = ((Get-ChildItem -Path $playerdir -filter "*.mp3" | Sort-object -Property Name -Descending | Select-Object -First 1).Name).SubString(0,10)
							$files = Get-ChildItem $podcast.Directory -Filter "*.mp3" | Where-Object {[datetime](($_.Name).SubString(0,10)) -gt $mostrecent} 
							if ($files.count -gt 0) {
								Write-host "New episodes of $($podcast.Name) found, starting copy process..."
								foreach ($file in $files) {
									Copy-Item -Path $file.Fullname -Destination $playerdir
								}
							} else {
								Write-host "No new episodes available."
							}
							remove-variable -name mostrecent,files -Force -Erroraction SilentlyContinue	
						}
					}
					Remove-Variable -name playerdir -force -erroraction SilentlyContinue
				}
				remove-variable -name podmatch -force -ErrorAction SilentlyContinue
			} else {
				Write-Host "Cannot find podcast directory with specified path $($playerpath)!" -foregroundcolor red
			}
		}
		"Single" {
			<#	Single:
			Select podcast to update on player
			if player folder exists
				Check for most recent ep on player
				Copy all episodes more recent than that
			else 
				Create folder on player
				Prompt user for how many episodes to copy
				Copy selected episodes sequentially
			#>
			Write-Host "Not implemented yet!" -foregroundcolor yellow
		}
		"All" {
			<#	All:
			Foreach podcast:
				if player folder exists
					Check for most recent ep on player
					Copy all episodes more recent than that
				else 
					Create folder on player
					Prompt user for how many episodes to copy
					Copy selected episodes sequentially
			#>
			Write-Host "Not implemented yet!" -foregroundcolor yellow
		}
		default {
			Write-Host "Invalid player update mode specified!" -foregroundcolor red
		}
	}
	Start-Sleep 3
}
	

###############
# Main Script
###############

# Enable TLS 1.2 support
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;

if (!$interactive) {
	# Unattended Mode
	Clear-Host
	Write-Host "Podcast Manager - Unattended Mode`n"
	if ($global:podcasts) {
		# Iterate through the config array and call Get-PendingEpisodes to identify available downloads
		for ($i = 0;$i -lt $global:podcasts.count;$i++) {
			Get-PendingEpisodes -podchoice $i
		}
		# Download available episodes
		Download-PendingEpisodes
		# Update portable player with new episodes
		Update-PlayerFiles
		Write-Host "All operations complete, terminating."
	} else {
		Write-Host "Halting as configuration could not be loaded." -foregroundcolor red
	}
} else {
	# Interactive Mode
	Clear-Host
	Write-host "Podcast Manager`n" -foregroundcolor white
	# Check config file has loaded, start menu
	if ($global:podcasts) {
		Show-Menu
	} else {
		Write-Host "Halting as configuration could not be loaded." -foregroundcolor red
	}
}