# ===================================================================================================
#                          FUNCTION: Convert_Prefix_Length_To_Subnet_Mask
# ---------------------------------------------------------------------------------------------------
# Description:
#   Converts an IPv4 prefix length (24) to a dotted subnet mask (255.255.255.0).
#
# Returns:
#   The dotted mask, or $null when the prefix length is not 0-32.
# ===================================================================================================

function Convert_Prefix_Length_To_Subnet_Mask
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		$PrefixLength
	)
	
	$prefix = 0
	if ($null -eq $PrefixLength -or -not [int]::TryParse(([string]@($PrefixLength)[0]).Trim(), [ref]$prefix)) { return $null }
	if ($prefix -lt 0 -or $prefix -gt 32) { return $null }
	
	$bits = ('1' * $prefix).PadRight(32, '0')
	$octets = for ($i = 0; $i -lt 32; $i += 8) { [Convert]::ToInt32($bits.Substring($i, 8), 2) }
	return ($octets -join '.')
}
