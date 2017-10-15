$ModulesDirectory = [Environment]::GetEnvironmentVariable("PSModulePath")
$ModulePathOne = "$PSScriptRoot../lib/yaml/powershell-yaml.psm1"
$ModulePathTwo = "$PSScriptRoot../lib/yaml/powershell-yaml.psd1"
Import-Module -name $ModulePathOne
Import-Module -name $ModulePathTwo

[CmdletBinding()]
param(
    [string]$filepath
)

$yamlfile = Get-Content $filepath
$yamldict = ConvertFrom-Yaml $yamlfile

# Start PowerPoint and open it invisably as a template presentation
$powerpoint = Start-PowerPoint
$ppt = Open-Presentation $powerpoint "$PSScriptRoot../mod/mod.pptm" 0

Insert-IntroSlide $ppt $yamldict.'course'
Insert-TOCSlide $ppt $yamldict.'chapters' $yamldict.'labs'

$yamldict.'chapters' | foreach-object {
    Insert-ChapterSlide $ppt $_.name
    $	
}

# Save the PowerPoint presentation as the course title to the wrk directory
Save-Presentation $ppt $yamldict.'course'
