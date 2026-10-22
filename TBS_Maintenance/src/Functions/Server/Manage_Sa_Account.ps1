# ===================================================================================================
#                              FUNCTION: Manage_Sa_Account
# ---------------------------------------------------------------------------------------------------
# Description:
#   Form with Enable/Disable buttons for the 'sa' login on the local SQL Server. Enable sets
#   the sa password to 'TB$upp0rT'; button states reflect the current sa state (queried from
#   sys.sql_logins). Closes after a successful action.
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:FunctionResults: SqlModuleName, ConnectionString, DBSERVER (from SQL detection)
#
# Notes:
#   - Assumes the default SQL instance and sysadmin privileges.
# ===================================================================================================

function Manage_Sa_Account
{
	[CmdletBinding()]
	param ()
	
	Write_Log "`r`n==================== Starting Manage_Sa_Account Function ====================`r`n" "blue"
	
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	# Use already detected SQL module and connection info
	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$serverInstance = $script:FunctionResults['DBSERVER']
	$database = "master"
	
	if (-not $SqlModuleName -or $SqlModuleName -eq "None")
	{
		Write_Log "No valid SQL module detected for SQL operations!" "red"
		return
	}
	Import-Module $SqlModuleName -ErrorAction Stop
	$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
	
	$supportsConnectionString = $false
	if ($InvokeSqlCmd)
	{
		$supportsConnectionString = $InvokeSqlCmd.Parameters.Keys -contains 'ConnectionString'
	}
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Manage SQL 'sa' Account"
	$form.Size = New-Object System.Drawing.Size(300, 150)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	
	$btnEnable = New-Object System.Windows.Forms.Button
	$btnEnable.Text = "Enable 'sa'"
	$btnEnable.Location = New-Object System.Drawing.Point(50, 30)
	$btnEnable.Size = New-Object System.Drawing.Size(200, 30)
	$form.Controls.Add($btnEnable)
	
	$btnDisable = New-Object System.Windows.Forms.Button
	$btnDisable.Text = "Disable 'sa'"
	$btnDisable.Location = New-Object System.Drawing.Point(50, 70)
	$btnDisable.Size = New-Object System.Drawing.Size(200, 30)
	$form.Controls.Add($btnDisable)
	
	# --- Initial state update ---
	try
	{
		$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
		if ($supportsConnectionString)
		{
			$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
		}
		else
		{
			$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
		}
		if ($result)
		{
			$isEnabled = ($result.is_disabled -eq 0)
		}
		else
		{
			Write_Log "'sa' account not found." "red"
			$isEnabled = $false
		}
	}
	catch
	{
		Write_Log "Error checking 'sa' status: $_" "red"
		$isEnabled = $false
	}
	$btnEnable.Enabled = -not $isEnabled
	$btnDisable.Enabled = $isEnabled
	if ($isEnabled)
	{
		Write_Log "'sa' is currently enabled. Enable button greyed out." "yellow"
	}
	else
	{
		Write_Log "'sa' is currently disabled. Disable button greyed out." "yellow"
	}
	
	# --- Enable 'sa' button click event ---
	$btnEnable.Add_Click({
			Write_Log "Enable button clicked. Attempting to enable 'sa'..." "blue"
			try
			{
				$enableQuery = "ALTER LOGIN sa ENABLE; ALTER LOGIN sa WITH PASSWORD = 'TB`$upp0rT';"
				if ($supportsConnectionString)
				{
					Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $enableQuery -ErrorAction Stop
				}
				else
				{
					Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $enableQuery -ErrorAction Stop
				}
				Write_Log "'sa' account enabled and password set successfully." "green"
				# Refresh state and close
				try
				{
					$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
					if ($supportsConnectionString)
					{
						$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
					}
					else
					{
						$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
					}
					if ($result) { $isEnabled = ($result.is_disabled -eq 0) }
					else { $isEnabled = $false }
				}
				catch { $isEnabled = $false }
				$btnEnable.Enabled = -not $isEnabled
				$btnDisable.Enabled = $isEnabled
				$form.Close()
			}
			catch
			{
				Write_Log "Error enabling 'sa' account: $_" "red"
			}
		})
	
	# --- Disable 'sa' button click event ---
	$btnDisable.Add_Click({
			Write_Log "Disable button clicked. Attempting to disable 'sa'..." "blue"
			try
			{
				$disableQuery = "ALTER LOGIN sa DISABLE;"
				if ($supportsConnectionString)
				{
					Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $disableQuery -ErrorAction Stop
				}
				else
				{
					Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $disableQuery -ErrorAction Stop
				}
				Write_Log "'sa' account disabled successfully." "green"
				# Refresh state and close
				try
				{
					$query = "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'"
					if ($supportsConnectionString)
					{
						$result = Invoke-Sqlcmd -ConnectionString $ConnectionString -Query $query -ErrorAction Stop
					}
					else
					{
						$result = Invoke-Sqlcmd -ServerInstance $serverInstance -Database $database -Query $query -ErrorAction Stop
					}
					if ($result) { $isEnabled = ($result.is_disabled -eq 0) }
					else { $isEnabled = $false }
				}
				catch { $isEnabled = $false }
				$btnEnable.Enabled = -not $isEnabled
				$btnDisable.Enabled = $isEnabled
				$form.Close()
			}
			catch
			{
				Write_Log "Error disabling 'sa' account: $_" "red"
			}
		})
	
	$form.ShowDialog() | Out-Null
	Write_Log "`r`n==================== Manage_Sa_Account Function Completed ====================" "blue"
}
