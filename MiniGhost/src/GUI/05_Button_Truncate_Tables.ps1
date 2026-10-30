############################################################################
# 4) Truncate Tables Button
############################################################################
$truncateTablesButton = New-Object System.Windows.Forms.Button
$truncateTablesButton.Text = "Truncate Tables"
$truncateTablesButton.Location = New-Object System.Drawing.Point(10, 160)
$truncateTablesButton.Size = New-Object System.Drawing.Size(150, 35)
$truncateTablesButton.Add_Click({
		# Define the tables to truncate
		$tablesToTruncate = @(
			"TRS_LOG", "TRS_FIN", "TRS_DPT", "TRS_SUB", "TRS_CLK",
			"TRS_CLT", "TRS_VND", "TRS_ITM",
			"RPT_CLT_D", "RPT_CLT_W", "RPT_CLT_M", "RPT_CLT_N", "RPT_CLT_Y", "RPT_CLT_P", "RPT_CLT_F", "RPT_CLT_ITM_D", "RPT_CLT_ITM_N",
			"RPT_FIN", "RPT_DPT", "RPT_SUB", "RPT_HOU", "RPT_VND",
			"COST_REV", "POS_REV", "OBJ_REV", "PRICE_REV", "REV_HDR", "SAL_REG_SAV", "SAL_HDR_SAV", "SAL_TTL_SAV", "SAL_DET_SAV",
			"RPT_CLK_D", "RPT_CLK_W", "RPT_CLK_M", "RPT_CLK_Y", "RPT_CLK_P", "RPT_CLK_F", "RPT_CLK_N",
			"RPT_ITM_D", "RPT_ITM_W", "RPT_ITM_M", "RPT_ITM_Y", "RPT_ITM_P", "RPT_ITM_F", "RPT_ITM_N",
			"DATA_REG",
			"RPT_CHK", "RENT_TAB", "TIM_TAB", "GAS_INVENT", "GAS_COUNT", "GAS_TRANS",
			"SAL_BAT", "SAL_HDR", "SAL_REG", "SAL_DET", "SAL_TTL",
			"REC_BAT", "REC_HDR", "REC_REG", "REC_TTL", "INV_HDR", "INV_REG", "INV_TTL",
			"RPT_PAY_N", "RPT_PAY_M", "TRS_PAY", "OFR_TAB"
		)
		
		$result = [System.Windows.Forms.MessageBox]::Show("Do you want to truncate the specified tables?", "Truncate Tables", [System.Windows.Forms.MessageBoxButtons]::YesNo)
		if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Truncate tables
			$failedTruncateTables = RPT_Blank -tables $tablesToTruncate
			
			if ($failedTruncateTables.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("Tables truncated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["TableTruncation"].Status = "Successful"
				$operationStatus["TableTruncation"].Message = "Tables truncated successfully."
				$operationStatus["TableTruncation"].Details = "All tables were truncated successfully."
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to truncate some tables: $($failedTruncateTables -join ', ')", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["TableTruncation"].Status = "Failed"
				$operationStatus["TableTruncation"].Message = "Failed to truncate some tables."
				$operationStatus["TableTruncation"].Details = "Failed tables: $($failedTruncateTables -join ', ')"
			}
		}
		else
		{
			$operationStatus["TableTruncation"].Status = "Skipped"
			$operationStatus["TableTruncation"].Message = "User chose not to truncate tables."
		}
	})
