############################################################################
# Handle form closing event (X button)
############################################################################
# Closes straight away (no "Are you sure?" prompt since v1.4.1): every action already asks before it
# changes anything, and nothing runs in the background that closing could cut off.
$form.add_FormClosing({
		# Clean the temp copies of MiniGhost
		[void](Remove_MiniGhost_Temp_Copies -TempPath $TempDir)
	})
