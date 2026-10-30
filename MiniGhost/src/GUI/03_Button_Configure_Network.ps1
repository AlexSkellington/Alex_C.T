############################################################################
# 2) Configure Network Button
############################################################################
$configureNetworkButton = New-Object System.Windows.Forms.Button
$configureNetworkButton.Text = "Configure Network"
$configureNetworkButton.Location = New-Object System.Drawing.Point(170, 120)
$configureNetworkButton.Size = New-Object System.Drawing.Size(150, 35)
$configureNetworkButton.Add_Click({
		# Implement the ConfigureNetwork function with GUI elements
		
		if ($currentConfigs -eq $null -or $currentConfigs.Count -eq 0)
		{
			[System.Windows.Forms.MessageBox]::Show("No valid active IP configuration found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
		
		# If multiple adapters are found, ask the user to select one
		if ($currentConfigs.Count -gt 1)
		{
			# Create a form to select the network adapter
			$adapterForm = New-Object System.Windows.Forms.Form
			$adapterForm.Text = "Select Network Adapter"
			$adapterForm.Size = New-Object System.Drawing.Size(400, 200)
			$adapterForm.StartPosition = "CenterParent"
			
			$label = New-Object System.Windows.Forms.Label
			$label.Text = "Select the network adapter:"
			$label.Location = New-Object System.Drawing.Point(10, 10)
			$label.Size = New-Object System.Drawing.Size(380, 20)
			$label.AutoSize = $true
			
			$listBox = New-Object System.Windows.Forms.ListBox
			$listBox.Location = New-Object System.Drawing.Point(10, 40)
			$listBox.Size = New-Object System.Drawing.Size(360, 80)
			
			for ($i = 0; $i -lt $currentConfigs.Count; $i++)
			{
				$adapterName = $currentConfigs[$i].InterfaceAlias
				$ipAddress = $currentConfigs[$i].IPv4Address.IPAddress
				$listBox.Items.Add("$adapterName - IP: $ipAddress")
			}
			
			$okButton = New-Object System.Windows.Forms.Button
			$okButton.Text = "OK"
			$okButton.Location = New-Object System.Drawing.Point(150, 130)
			$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
			
			$adapterForm.AcceptButton = $okButton
			$adapterForm.Controls.AddRange(@($label, $listBox, $okButton))
			$adapterForm.ShowDialog() | Out-Null
			
			$selectedIndex = $listBox.SelectedIndex
			if ($selectedIndex -ge 0 -and $selectedIndex -lt $currentConfigs.Count)
			{
				$currentConfig = $currentConfigs[$selectedIndex]
			}
			else
			{
				[System.Windows.Forms.MessageBox]::Show("No network adapter selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				return
			}
		}
		else
		{
			# Only one adapter found
			$currentConfig = $currentConfigs[0]
		}
		
		# Proceed with configuring network
		# Display current IP Address
		$currentIP = $currentConfig.IPv4Address.IPAddress
		$currentGateway = $currentConfig.IPv4DefaultGateway.NextHop
		$adapterName = $currentConfig.InterfaceAlias
		
		# Ask user for network type
		$networkTypeResult = [System.Windows.Forms.MessageBox]::Show("Will the adapter use DHCP?", "Network Configuration", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
		
		if ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::Yes)
		{
			# Configure DHCP
			try
			{
				# netsh never throws: check its exit code so a failure is not reported as success
				# ("DHCP is already enabled on this interface" also exits 1 - that one is fine)
				$netshOutput = netsh interface ip set address name="$adapterName" source=dhcp 2>&1
				if ($LASTEXITCODE -ne 0 -and (($netshOutput | Out-String) -notmatch '(?i)already')) { throw ("netsh (address) exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
				$netshOutput = netsh interface ip set dns name="$adapterName" source=dhcp 2>&1
				if ($LASTEXITCODE -ne 0 -and (($netshOutput | Out-String) -notmatch '(?i)already')) { throw ("netsh (dns) exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
				[System.Windows.Forms.MessageBox]::Show("DHCP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				$operationStatus["IPConfiguration"].Status = "Successful"
				$operationStatus["IPConfiguration"].Message = "DHCP configuration applied."
				$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName"
			}
			catch
			{
				[System.Windows.Forms.MessageBox]::Show("Failed to configure DHCP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				$operationStatus["IPConfiguration"].Status = "Failed"
				$operationStatus["IPConfiguration"].Message = "Failed to configure DHCP."
				$operationStatus["IPConfiguration"].Details = "Error: $_"
			}
			
		}
		elseif ($networkTypeResult -eq [System.Windows.Forms.DialogResult]::No)
		{
			# Configure Static IP
			# Loop until valid input or user cancels
			$validInput = $false
			while (-not $validInput)
			{
				# Ask for last octet
				$gatewayParts = $currentGateway.Split('.')
				if ($gatewayParts.Length -lt 4)
				{
					[System.Windows.Forms.MessageBox]::Show("Invalid gateway format: $currentGateway", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					$operationStatus["IPConfiguration"].Status = "Failed"
					$operationStatus["IPConfiguration"].Message = "Invalid gateway format."
					$operationStatus["IPConfiguration"].Details = "Gateway: $currentGateway"
					return
				}
				$gatewayPrefix = $gatewayParts[0 .. 2] -join '.'
				$ipForm = New-Object System.Windows.Forms.Form
				$ipForm.Text = "Enter Last Octet for Static IP"
				$ipForm.Size = New-Object System.Drawing.Size(350, 170)
				$ipForm.StartPosition = "CenterParent"
				
				$label = New-Object System.Windows.Forms.Label
				$label.Text = "Enter the last octet (1-254) of the static IP address $gatewayPrefix."
				$label.Location = New-Object System.Drawing.Point(10, 20)
				$label.Size = New-Object System.Drawing.Size(320, 40)
				$label.AutoSize = $true
				
				$textBox = New-Object System.Windows.Forms.TextBox
				$textBox.Location = New-Object System.Drawing.Point(10, 70)
				$textBox.Size = New-Object System.Drawing.Size(320, 20)
				
				$okButton = New-Object System.Windows.Forms.Button
				$okButton.Text = "OK"
				$okButton.Location = New-Object System.Drawing.Point(130, 100)
				$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
				
				$ipForm.AcceptButton = $okButton
				
				$ipForm.Controls.AddRange(@($label, $textBox, $okButton))
				$dialogResult = $ipForm.ShowDialog()
				
				if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
				{
					$lastOctet = $textBox.Text.Trim()
					# Validate last octet
					if ($lastOctet -match '^\d{1,3}$' -and [int]$lastOctet -ge 1 -and [int]$lastOctet -le 254)
					{
						$ipAddress = "$gatewayPrefix.$lastOctet"
						
						# Check if IP is in use
						$pingResult = Test-Connection -ComputerName $ipAddress -Count 1 -Quiet
						if ($pingResult)
						{
							[System.Windows.Forms.MessageBox]::Show("The IP address '$ipAddress' is already in use. Please choose a different one.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
						}
						else
						{
							try
							{
								# Keep the adapter's current subnet size (fall back to /24 when it is unknown)
								$subnetMask = '255.255.255.0'
								$currentPrefix = 0
								if ([int]::TryParse([string](@($currentConfig.IPv4Address)[0].PrefixLength), [ref]$currentPrefix) -and $currentPrefix -ge 8 -and $currentPrefix -le 30)
								{
									$subnetMask = Convert_Prefix_Length_To_Subnet_Mask -PrefixLength $currentPrefix
								}
								
								# netsh never throws: check its exit code so a failure is not reported as success
								$netshOutput = netsh interface ip set address name="$adapterName" static $ipAddress $subnetMask $currentGateway 2>&1
								if ($LASTEXITCODE -ne 0) { throw ("netsh exit code {0}: {1}" -f $LASTEXITCODE, (($netshOutput | Out-String).Trim())) }
								
								# Update the label
								$ipAddressLabel.Text = "IP Address changed from: $currentIP to $ipAddress"
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								
								# Set DNS
								Set-DnsClientServerAddress -InterfaceAlias $adapterName -ServerAddresses ("8.8.8.8", "8.8.4.4")
								[System.Windows.Forms.MessageBox]::Show("Static IP configuration applied.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
								$operationStatus["IPConfiguration"].Status = "Successful"
								$operationStatus["IPConfiguration"].Message = "Static IP configuration applied."
								$operationStatus["IPConfiguration"].Details = "Adapter: $adapterName, IP: $ipAddress"
								$validInput = $true
							}
							catch
							{
								[System.Windows.Forms.MessageBox]::Show("Failed to configure static IP: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
								$operationStatus["IPConfiguration"].Status = "Failed"
								$operationStatus["IPConfiguration"].Message = "Failed to configure static IP."
								$operationStatus["IPConfiguration"].Details = "Error: $_"
								$validInput = $true # Exit loop since an error occurred
							}
						}
					}
					else
					{
						[System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a valid number between 1 and 254.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					}
				}
				else
				{
					# User canceled the dialog
					$operationStatus["IPConfiguration"].Status = "Skipped"
					$operationStatus["IPConfiguration"].Message = "User canceled static IP configuration."
					$operationStatus["IPConfiguration"].Details = "" # Empty details
					$validInput = $true # Exit loop
				}
			}
			
		}
		else
		{
			$operationStatus["IPConfiguration"].Status = "Skipped"
			$operationStatus["IPConfiguration"].Message = "User chose not to configure network settings."
		}
	})
