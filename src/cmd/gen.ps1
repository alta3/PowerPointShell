[CmdletBinding()]
param(
    [string]$inFile
)

# Create a PowerPoint presentation and keep it invisible until automation is complete
$ppt = Add-Presentation

