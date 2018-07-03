function Get-WsusUpdates {
	[CmdletBinding()]
	Param([switch]$Silent)
	
	$wsusScope = "IsInstalled=0 and Type='Software'"
	
	$wsusSearch = New-Object -ComObject Microsoft.Update.Searcher
	
	$wsusSearchResults = $wsusSearch.Search($wsusScope).Updates
	
	if (!$Silent) {
		Write-Host ""
		Write-Host "$($wsusSearchResults.Count) Updates Available:"
		Write-Host ""
		
		$i = 1
		foreach ($wsusSearchResult in $wsusSearchResults) {
			$downloaded = "Not Yet Downloaded"
			if ($wsusSearchResult.IsDownloaded) {
				$downloaded = "Downloaded, Ready To Install"
			}
			Write-Host "${i}. $($wsusSearchResult.Title) - ${downloaded}"
			$i++
		}
		Write-Host ""
	}
	else {
		$wsusSearchResults
	}
}

function Install-WsusUpdates {
	[CmdletBinding()]
	Param(
		[switch]$DownloadOnly,
		[string[]]$Include,
		[string[]]$Exclude
	)
	
	if (($Include.Count -gt 0) -and ($Exclude.Count -gt 0)) {
		Write-Error ""
		Write-Error "You Can't Include and Exclude... Do One or the Other."
		Write-Error ""
		return 2
	}
	
	$wsus = New-Object -ComObject Microsoft.Update.Session
	
	$wsusUpdates = Get-WsusUpdates -Silent
	
	$wsusUpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
	
	foreach ($wsusUpdate in $wsusUpdates) {
		if (!wsusUpdate.IsDownloaded) {
			if ($Exclude.Count -gt 0) {
				foreach ($e in $Exclude) {
					if ($wsusUpdate.KBArticleIDs -notcontains ($e -replace '[a-zA-z]')) { $wsusUpdatesToDownload.Add($wsusUpdate) }
				}
			}
			elseif ($Include.Count -gt 0) {
				foreach ($i in $Include) {
					if ($wsusUpdate.KBArticleIDs -contains ($i -replace '[a-zA-z]')) { $wsusUpdatesToDownload.Add($wsusUpdate) }
				}
			}
			else {
				$wsusUpdatesToDownload.Add($wsusUpdate)
			}
		}
	}
	
	if ($wsusUpdatesToDownload.Count -gt 0) {
		Write-Host ""
		Write-Host "Downloading $($wsusUpdatesToDownload.Count) Updates."
		Write-Host ""
		
		$wsusDownloader = $wsus.CreateUpdateDownloader()
		$wsusDownloader.Updates = $wsusUpdatesToDownload
		$wsusDownloadStatus = try { 
			$wsusDownloader.Download() 
			Write-Host ""
			Write-Host "$($wsusUpdatesToDownload.Count) Updates Downloaded Successfully."
			Write-Host ""
		}
		catch { 
			Write-Error ""
			Write-Error "Error Downloading Updates. Make Sure You are Running as Admin."
			Write-Error ""
			return 3
		}
	}
	
	if (!$DownloadOnly) {
		$wsusUpdates = Get-WsusUpdates -Silent
	
		$wsusUpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
		
		foreach ($wsusUpdate in $wsusUpdates) {
			if (wsusUpdate.IsDownloaded) {
				if ($Exclude.Count -gt 0) {
					foreach ($e in $Exclude) {
						if ($wsusUpdate.KBArticleIDs -notcontains ($e -replace '[a-zA-z]')) { $wsusUpdatesToInstall.Add($wsusUpdate) }
					}
				}
				elseif ($Include.Count -gt 0) {
					foreach ($i in $Include) {
						if ($wsusUpdate.KBArticleIDs -contains ($i -replace '[a-zA-z]')) { $wsusUpdatesToInstall.Add($wsusUpdate) }
					}
				}
				else {
					$wsusUpdatesToInstall.Add($wsusUpdate)
				}
			}
		}
		
		$wsusInstaller = $wsus.CreateUpdateInstaller()
		$wsusInstaller.Updates = $wsusUpdatesToInstall
		
		Write-Host ""
		Write-Host "Installing $($wsusUpdatesToInstall.Count) Updates."
		Write-Host ""
		
		$wsusInstallStatus = try {
			$wsusInstaller.Install()
		}
		catch {
			Write-Error ""
			Write-Error "Error Installing Updates. Make Sure You are Running as Admin."
			Write-Error ""
			return 1
		}
		
		if ($wsusInstallStatus.rebootRequired) {
			Write-Host ""
			Write-Host "$($wsusUpdatesToInstall.Count) Updates Installed Successfully."
			Write-Host ""
			Write-Host ""
			Write-Host -ForegroundColor Yellow "!! A Reboot Is Required. Please do so ASAP. !!"
			Write-Host ""
		}
		elseif (!$wsusInstallStatus.rebootRequired) {
			Write-Host ""
			Write-Host "$($wsusUpdatesToInstall.Count) Updates Installed Successfully."
			Write-Host ""
			Write-Host ""
			Write-Host -ForegroundColor Green "No Reboot is Required."
			Write-Host ""
		}
	}
	else {
		Write-Host ""
		Write-Host "Cmdlet Called with -DownloadOnly; Exiting Without Installing Updates."
		Write-Host ""
		return 0
	}
}
