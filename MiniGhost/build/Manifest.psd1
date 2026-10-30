@{
	# ---------------------------------------------------------------------------------------------
	# Build manifest for MiniGhost.
	#
	# Build.ps1 concatenates every file matched below, in this order, into ONE script: the file
	# that 'irm https://bit.ly/MiniGhost | iex' runs (root MiniGhost_V38.ps1 launcher ->
	# MiniGhost/MiniGhost.ps1 on main) and that PowerShell Studio packages into the .exe.
	#
	#   - Paths are relative to the MiniGhost folder.
	#   - Each pattern expands sorted by full path, so numeric prefixes (01_, 02_ ...) set the
	#     order inside a folder. Function files have no prefix: definition order does not matter.
	#   - Every .ps1 under src\ must be matched by exactly one pattern (the build fails on
	#     orphans and duplicates), so a new function file only needs to be dropped into the
	#     right Functions\<Category> folder.
	# ---------------------------------------------------------------------------------------------

	# Output, relative to the repository root. Unlike the TBS script this file IS committed:
	# the launcher downloads it from main (a GitHub Release would clash with the TBS releases,
	# which own releases/latest). Build.ps1 -Check fails when it is out of date.
	OutputFile = 'MiniGhost/MiniGhost.ps1'

	Sources    = @(
		# 1. Header, version, assemblies, elevation, globals and paths (runs before anything else)
		'src/Bootstrap/*.ps1'

		# 2. Function library (definitions only - nothing runs here)
		'src/Functions/UI/*.ps1'
		'src/Functions/Store/*.ps1'
		'src/Functions/Sql/*.ps1'
		'src/Functions/System/*.ps1'
		'src/Functions/Workflow/*.ps1'

		# 3. Read store data and node role
		'src/Main/01_Load_Store_Data.ps1'

		# 4. Main window and its buttons
		'src/GUI/*.ps1'

		# 5. Show the window, exit
		'src/Main/02_Show_Form.ps1'
	)

	# Appended after the final 'exit' as a comment block (never executed).
	Changelog  = 'CHANGELOG.txt'
}
