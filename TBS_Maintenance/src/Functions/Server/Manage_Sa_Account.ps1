# ===================================================================================================
#                              FUNCTION: Manage_Sa_Account
# ---------------------------------------------------------------------------------------------------
# Description:
#   Form with Enable/Disable buttons for the 'sa' login on the store SQL Server. Enable also
#   sets the sa password to the standard support password ($script:BackupSqlPass). The button
#   for the current state is greyed out (read from sys.sql_logins); the form closes after a
#   successful change.
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:FunctionResults: SqlModuleName, ConnectionString, DBSERVER
#   - $script:BackupSqlPass: the password set when enabling sa
#
# Notes:
#   - Needs sysadmin rights on the instance. Uses Invoke-Sqlcmd when a SQL module is
#     available, else the built-in SqlClient.
# ===================================================================================================

function Manage_Sa_Account
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Manage_Sa_Account Function ====================`r`n" "blue"

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$SqlModuleName = $script:FunctionResults['SqlModuleName']
	$ConnectionString = $script:FunctionResults['ConnectionString']
	$serverInstance = $script:FunctionResults['DBSERVER']
	if ((-not $ConnectionString -or $ConnectionString -eq 'N/A') -and (-not $serverInstance -or $serverInstance -eq 'N/A'))
	{
		Write_Log "No SQL server / connection string known for this store - cannot manage 'sa'." "red"
		return
	}

	$InvokeSqlCmd = $null
	if ($SqlModuleName -and $SqlModuleName -ne "None")
	{
		try
		{
			Import-Module $SqlModuleName -ErrorAction Stop
			$InvokeSqlCmd = Get-Command Invoke-Sqlcmd -Module $SqlModuleName -ErrorAction Stop
		}
		catch { $InvokeSqlCmd = $null }
	}
	$supportsConnectionString = ($InvokeSqlCmd -and $InvokeSqlCmd.Parameters.ContainsKey('ConnectionString'))

	# Run one statement against the instance; returns rows (Invoke-Sqlcmd or SqlClient)
	$runSql = {
		param ([string]$Query)
		if ($InvokeSqlCmd)
		{
			if ($supportsConnectionString) { return & $InvokeSqlCmd -ConnectionString $ConnectionString -Query $Query -ErrorAction Stop }
			return & $InvokeSqlCmd -ServerInstance $serverInstance -Database master -Query $Query -ErrorAction Stop
		}
		$cs = if ($ConnectionString -and $ConnectionString -ne 'N/A') { $ConnectionString } else { "Server=$serverInstance;Database=master;Integrated Security=True;TrustServerCertificate=True;" }
		$cn = New-Object System.Data.SqlClient.SqlConnection $cs
		try
		{
			$cn.Open()
			$cmd = $cn.CreateCommand()
			$cmd.CommandText = $Query
			$dt = New-Object System.Data.DataTable
			[void](New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($dt)
			return , $dt
		}
		finally { $cn.Dispose() }
	}
	$getSaEnabled = {
		$rows = @(& $runSql "SELECT is_disabled FROM sys.sql_logins WHERE name = 'sa'" | ForEach-Object { $_ })
		if ($rows.Count -eq 0) { return $null }
		return ([int]$rows[0].is_disabled -eq 0)
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

	# --- Current state ---
	$isEnabled = $false
	try
	{
		$state = & $getSaEnabled
		if ($null -eq $state) { Write_Log "'sa' account not found." "red" }
		else { $isEnabled = $state }
	}
	catch { Write_Log "Error checking 'sa' status: $_" "red" }
	$btnEnable.Enabled = -not $isEnabled
	$btnDisable.Enabled = $isEnabled
	if ($isEnabled) { Write_Log "'sa' is currently enabled. Enable button greyed out." "yellow" }
	else { Write_Log "'sa' is currently disabled. Disable button greyed out." "yellow" }

	# --- Enable / Disable ---
	$apply = {
		param ([bool]$Enable)
		$verb = if ($Enable) { 'enable' } else { 'disable' }
		Write_Log "$((Get-Culture).TextInfo.ToTitleCase($verb)) button clicked. Attempting to $verb 'sa'..." "blue"
		try
		{
			if ($Enable)
			{
				$saPassword = ([string]$script:BackupSqlPass).Replace("'", "''")
				$null = & $runSql "ALTER LOGIN sa ENABLE; ALTER LOGIN sa WITH PASSWORD = '$saPassword';"
				Write_Log "'sa' account enabled and password set successfully." "green"
			}
			else
			{
				$null = & $runSql "ALTER LOGIN sa DISABLE;"
				Write_Log "'sa' account disabled successfully." "green"
			}
			$form.Close()
		}
		catch { Write_Log "Error trying to $verb 'sa' account: $_" "red" }
	}
	$btnEnable.Add_Click({ & $apply $true })
	$btnDisable.Add_Click({ & $apply $false })

	[void]$form.ShowDialog()
	$form.Dispose()
	Write_Log "`r`n==================== Manage_Sa_Account Function Completed ====================" "blue"
}
