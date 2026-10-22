@{
	# ---------------------------------------------------------------------------------------------
	# Build manifest for the TBS Maintenance Script.
	#
	# Build.ps1 concatenates every file matched below, in this order, into ONE script:
	# the file that 'irm https://bit.ly/TBS_Maintenance_Script | iex' downloads and that
	# PowerShell Studio packages into the .exe.
	#
	#   - Paths are relative to the TBS_Maintenance folder.
	#   - Each pattern expands sorted by full path, so numeric prefixes (01_, 02_ ...) set the
	#     order inside a folder. Function files have no prefix: definition order does not matter.
	#   - Every .ps1 under src\ must be matched by exactly one pattern (the build fails on
	#     orphans and duplicates), so a new function file only needs to be dropped into the
	#     right Functions\<Category> folder.
	# ---------------------------------------------------------------------------------------------

	# Output, relative to the repository root. dist\ is git-ignored: the built file is not
	# committed, it is attached to a GitHub Release under exactly this name, and the short link
	# points at releases/latest/download/TBS_Maintenance_Script.ps1.
	OutputFile = 'TBS_Maintenance/dist/TBS_Maintenance_Script.ps1'

	Sources    = @(
		# 1. Launch detection, elevation, globals, storeman detection (runs before anything else)
		'src/Bootstrap/*.ps1'

		# 2. Function library (definitions only - nothing runs here)
		'src/Functions/Core/*.ps1'
		'src/Functions/Store/*.ps1'
		'src/Functions/Background/*.ps1'
		'src/Functions/Dialogs/*.ps1'
		'src/Functions/Server/*.ps1'
		'src/Functions/Lanes/*.ps1'
		'src/Functions/Scales/*.ps1'
		'src/Functions/Nodes/*.ps1'
		'src/Functions/General/*.ps1'

		# 3. Runtime init + headless worker mode (exits before the GUI when used)
		'src/Init/*.ps1'

		# 4. Main window
		'src/GUI/*.ps1'

		# 5. Load store data, warmup, show the window, exit
		'src/Main/*.ps1'
	)

	# Appended after the final 'exit' as a comment block (never executed).
	Changelog  = 'CHANGELOG.txt'
}
