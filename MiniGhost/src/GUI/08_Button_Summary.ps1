############################################################################
# 10) Summary Button
############################################################################
$summaryButton = New-Object System.Windows.Forms.Button
$summaryButton.Text = "Show Summary"
$summaryButton.Location = New-Object System.Drawing.Point(170, 240)
$summaryButton.Size = New-Object System.Drawing.Size(150, 35)
$summaryButton.Add_Click({
		# Display the summary in a new form
		# Build the summary text
		$summaryText = ""
		foreach ($operationKey in $operationStatus.Keys)
		{
			$operation = $operationStatus[$operationKey]
			$status = $operation.Status
			$message = $operation.Message
			$details = $operation.Details
			
			$summaryText += "${operationKey}: $status`r`n"
			if ($message -ne "")
			{
				$summaryText += "  Message: $message`r`n"
			}
			if ($details -ne "")
			{
				$summaryText += "  Details: $details`r`n"
			}
			$summaryText += "`r`n"
		}
		
		Show_Text_Report_Form -Title "Operation Summary" -Text $summaryText -Width 600 -Height 400
	})
