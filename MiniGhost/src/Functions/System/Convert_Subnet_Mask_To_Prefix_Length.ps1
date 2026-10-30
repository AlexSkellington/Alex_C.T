# ===================================================================================================
#                          FUNCTION: Convert_Subnet_Mask_To_Prefix_Length
# ---------------------------------------------------------------------------------------------------
# Description:
#   Converts a dotted IPv4 subnet mask (255.255.255.0) to its prefix length (24).
#
# Returns:
#   The prefix length (0-32), or $null when the mask is not a valid contiguous IPv4 mask.
# ===================================================================================================

function Convert_Subnet_Mask_To_Prefix_Length
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$SubnetMask
	)
	
	if ([string]::IsNullOrWhiteSpace($SubnetMask) -or $SubnetMask.Trim() -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $null }
	
	$bits = ""
	foreach ($octet in $SubnetMask.Trim().Split('.'))
	{
		$value = [int]$octet
		if ($value -gt 255) { return $null }
		$bits += [Convert]::ToString($value, 2).PadLeft(8, '0')
	}
	
	# A mask is ones followed by zeros only
	if ($bits -notmatch '^1*0*$') { return $null }
	$firstZero = $bits.IndexOf('0')
	if ($firstZero -lt 0) { return 32 }
	return $firstZero
}
