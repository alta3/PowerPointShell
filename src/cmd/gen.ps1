$ModulesDirectory = [Environment]::GetEnvironmentVariable("PSModulePath")
$ModulePathOne = $PSScriptRoot "../lib/yaml/powershell-yaml.psm1"
$ModulePathTwo = $PSScriptRoot "../lib/yaml/powershell-yaml.psd1"
Import-Module -name $ModulePathOne
Import-Module -name $ModulePathTwo

[CmdletBinding()]
param(
    [string]$filepath
)

$yamlfile = Get-Content $filepath
$yamlobj = ConvertFrom-Yaml $yamlfile

# Create a PowerPoint presentation and keep it invisible until automation is complete
$powerpoint = Add-PowerPoint
$ppt = Add-Presentation $powerpoint 0

