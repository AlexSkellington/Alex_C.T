############################################################################
# 9) Update Database Button (simple click call)
############################################################################
$updateSQLDatabaseButton = New-Object System.Windows.Forms.Button
$updateSQLDatabaseButton.Text = "Update SQL Database"
$updateSQLDatabaseButton.Location = New-Object System.Drawing.Point(10, 240)
$updateSQLDatabaseButton.Size = New-Object System.Drawing.Size(150, 35)

$updateSQLDatabaseButton.Add_Click({
		Update_SQL_Database `
								 -CurrentMachineName $currentMachineName `
								 -NewMachineName $script:newMachineName `
								 -OperationStatus $operationStatus
	})
