# ===================================================================================================
#                              FUNCTION: Clear_XE_Folder
# ---------------------------------------------------------------------------------------------------
# Description:
#   Keeps the lane XE (urgent messages) folder clean. At startup it deletes everything except
#   valid S*.??? health files (also clearing old FATAL* errors), then starts a background
#   watcher job that keeps newly arriving FATAL* files and deletes everything else except
#   current health files on a regular sweep.
#
# Parameters:
#   - FolderPath           [string]: XE folder (default <OfficePath>\XE<Store>901)
#   - SweepIntervalSeconds [int]: periodic sweep interval in seconds, 2-3600 (default 30)
#   - HealthMaxAgeDays     [int]: max age in days of health files kept, 1-3650 (default 30)
#
# Returns:
#   The watcher job when started; $null if the folder is missing or the job fails.
#
# Writes:
#   - Deletes files under the XE folder (everything except valid health files)
# ===================================================================================================

function Clear_XE_Folder
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$FolderPath = "$OfficePath\XE${StoreNumber}901",
		[Parameter(Mandatory = $false)]
		[ValidateRange(2, 3600)]
		[int]$SweepIntervalSeconds = 30,
		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 3650)]
		[int]$HealthMaxAgeDays = 30
	)

	# ------------------------------------------------------------------------------------------------
	# 0) Validate folder exists
	# ------------------------------------------------------------------------------------------------
	if (-not (Test-Path -LiteralPath $FolderPath))
	{
		Write_Log "Folder '$FolderPath' (Urgent Messages) does not exist." "red"
		return $null
	}

	# ------------------------------------------------------------------------------------------------
	# 1) STARTUP CLEANUP
	#    - Deletes EVERYTHING except valid S*.??? health files
	#    - ALSO deletes FATAL* (clear old errors)
	# ------------------------------------------------------------------------------------------------
	try
	{
		$now = Get-Date

		$items = Get-ChildItem -LiteralPath $FolderPath -Recurse -Force -ErrorAction SilentlyContinue |
		Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 }

		foreach ($item in $items)
		{
			$keep = $false

			# Startup mode: delete FATAL*
			if ($item.Name -like 'FATAL*')
			{
				$keep = $false
			}
			elseif ($item.PSIsContainer)
			{
				$keep = $false
			}
			elseif ($item.Name -match '^S.*\.\w{3}$')
			{
				if (($now - $item.LastWriteTime).TotalDays -le $HealthMaxAgeDays)
				{
					$content = $null
					try { $content = Get-Content -LiteralPath $item.FullName -ErrorAction Stop }
					catch { $content = $null }

					if ($content)
					{
						$fromLine = $content | Where-Object { $_ -like 'From:*' } | Select-Object -First 1
						$subjectLine = $content | Where-Object { $_ -like 'Subject:*' } | Select-Object -First 1
						$msgLine = $content | Where-Object { $_ -like 'MSG:*' } | Select-Object -First 1
						$lastStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' } | Select-Object -First 1

						if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
						{
							$fileStoreNumber = $Matches[1]
							$fileLaneNumber = $Matches[2]

							if ($fileStoreNumber -eq $StoreNumber -and
								$item.Name -match '^S.*\.(\d{3})$' -and
								$fileLaneNumber -eq $Matches[1] -and
								$subjectLine -match 'Subject:\s*Health' -and
								$msgLine -match 'MSG:\s*This application is not running\.' -and
								$lastStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
							{
								$keep = $true
							}
						}
					}
				}
			}

			if (-not $keep)
			{
				for ($i = 0; $i -lt 8; $i++)
				{
					try
					{
						if (Test-Path -LiteralPath $item.FullName)
						{
							Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop
						}
						break
					}
					catch { Start-Sleep -Milliseconds 150 }
				}
			}
		}
	}
	catch
	{
		Write_Log "An error occurred during startup cleaning of '$FolderPath': $_" "red"
	}

	# ------------------------------------------------------------------------------------------------
	# 2) MONITOR (FileSystemWatcher in a background job)
	#    - Keeps NEW FATAL* while running
	#    - Deletes everything else except valid health files
	# ------------------------------------------------------------------------------------------------
	try
	{
		$jobName = "XEWatcherJob_$StoreNumber"

		# Replace any existing watcher job
		$existing = Get-Job -Name $jobName -ErrorAction SilentlyContinue
		if ($existing)
		{
			try { Stop-Job -Job $existing -Force -ErrorAction SilentlyContinue | Out-Null }
			catch { }
			try { Remove-Job -Job $existing -Force -ErrorAction SilentlyContinue | Out-Null }
			catch { }
		}

		$job = Start-Job -Name $jobName -ArgumentList $FolderPath, $StoreNumber, $HealthMaxAgeDays, $SweepIntervalSeconds -ScriptBlock {
			param (
				[string]$FolderPath,
				[string]$StoreNumber,
				[int]$HealthMaxAgeDays,
				[int]$SweepIntervalSeconds
			)

			$ErrorActionPreference = 'SilentlyContinue'

			# --- Watcher object ---
			$fsw = New-Object System.IO.FileSystemWatcher
			$fsw.Path = $FolderPath
			$fsw.IncludeSubdirectories = $true
			$fsw.NotifyFilter = [IO.NotifyFilters]'FileName, DirectoryName, LastWrite, CreationTime'

			# --- Shared processor (MONITOR MODE: keep FATAL*) ---
			$ProcessPath = {
				param ([string]$Path)

				if (-not $Path) { return }

				# Retry Get-Item because events can fire before writer finishes.
				$item = $null
				for ($t = 0; $t -lt 8; $t++)
				{
					try
					{
						if (-not (Test-Path -LiteralPath $Path)) { return }
						$item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
						break
					}
					catch { Start-Sleep -Milliseconds 120 }
				}
				if (-not $item) { return }

				# Keep NEW FATAL* while running
				if ($item.Name -like 'FATAL*') { return }

				# Skip reparse points
				if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return }

				$now = Get-Date
				$keep = $false

				if ($item.PSIsContainer)
				{
					$keep = $false
				}
				elseif ($item.Name -match '^S.*\.\w{3}$')
				{
					if (($now - $item.LastWriteTime).TotalDays -le $HealthMaxAgeDays)
					{
						$content = $null
						try { $content = Get-Content -LiteralPath $item.FullName -ErrorAction Stop }
						catch { $content = $null }

						if ($content)
						{
							$fromLine = $content | Where-Object { $_ -like 'From:*' } | Select-Object -First 1
							$subjectLine = $content | Where-Object { $_ -like 'Subject:*' } | Select-Object -First 1
							$msgLine = $content | Where-Object { $_ -like 'MSG:*' } | Select-Object -First 1
							$lastStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' } | Select-Object -First 1

							if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
							{
								$fileStoreNumber = $Matches[1]
								$fileLaneNumber = $Matches[2]

								if ($fileStoreNumber -eq $StoreNumber -and
									$item.Name -match '^S.*\.(\d{3})$' -and
									$fileLaneNumber -eq $Matches[1] -and
									$subjectLine -match 'Subject:\s*Health' -and
									$msgLine -match 'MSG:\s*This application is not running\.' -and
									$lastStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
								{
									$keep = $true
								}
							}
						}
					}
				}

				if (-not $keep)
				{
					for ($i = 0; $i -lt 8; $i++)
					{
						try
						{
							if (Test-Path -LiteralPath $item.FullName)
							{
								Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop
							}
							break
						}
						catch { Start-Sleep -Milliseconds 150 }
					}
				}
			}

			# Register events (no -Action) and consume them with Wait-Event (best in jobs)
			$createdId = "XE_Created_$StoreNumber"
			$changedId = "XE_Changed_$StoreNumber"
			$renamedId = "XE_Renamed_$StoreNumber"

			Register-ObjectEvent -InputObject $fsw -EventName Created -SourceIdentifier $createdId | Out-Null
			Register-ObjectEvent -InputObject $fsw -EventName Changed -SourceIdentifier $changedId | Out-Null
			Register-ObjectEvent -InputObject $fsw -EventName Renamed -SourceIdentifier $renamedId | Out-Null

			$fsw.EnableRaisingEvents = $true
			$nextSweep = (Get-Date).AddSeconds($SweepIntervalSeconds)

			while ($true)
			{
				# Wait briefly for an event
				$e = Wait-Event -Timeout 1

				if ($e)
				{
					try
					{
						$path = $e.SourceEventArgs.FullPath
						& $ProcessPath -Path $path
					}
					catch { }
					finally
					{
						# Always remove handled event from queue
						try { Remove-Event -EventIdentifier $e.EventIdentifier -ErrorAction SilentlyContinue }
						catch { }
					}

					# Drain burst
					while ($true)
					{
						$e2 = Wait-Event -Timeout 0
						if (-not $e2) { break }
						try
						{
							$path2 = $e2.SourceEventArgs.FullPath
							& $ProcessPath -Path $path2
						}
						catch { }
						finally
						{
							try { Remove-Event -EventIdentifier $e2.EventIdentifier -ErrorAction SilentlyContinue }
							catch { }
						}
					}
				}

				# Periodic sweep fallback
				if ((Get-Date) -ge $nextSweep)
				{
					$nextSweep = (Get-Date).AddSeconds($SweepIntervalSeconds)

					try
					{
						if (Test-Path -LiteralPath $FolderPath)
						{
							Get-ChildItem -LiteralPath $FolderPath -Recurse -Force -ErrorAction SilentlyContinue |
							Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 } |
							ForEach-Object { & $ProcessPath -Path $_.FullName }
						}
					}
					catch { }
				}
			}
		}

		# Write_Log "XE watcher job started: $jobName" "green"
		return $job
	}
	catch
	{
		Write_Log "Failed to start XE watcher job: $_" "red"
		return $null
	}
}
