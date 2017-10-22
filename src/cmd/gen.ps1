function pps-generate {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$filepath
    )
    # Dot Source the core library
    . "$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\lib\core.ps1"

    #Ensure the $filepath exists and is in its long form 
    $newfilepath = Resolve-Path -Path $filepath | Select -ExpandProperty Path

    # Convert YAML into useable dictionary
    $yamlfile = Get-Content $newfilepath | Out-String
    $yamldict = ConvertFrom-Yaml -yaml $yamlfile

    # Start PowerPoint
    $powerpoint = Start-PowerPoint

    # Open Working PowerPoint
    $ppt = Open-Presentation -powerpoint $powerpoint -path "$PSScriptRoot\..\mod\mod.pptm" -visable $true
    SaveAs-Presentation -ppt $ppt -name $yamldict.course

    # Get the slide templates
    $SlideTemplates = Get-SlideTemplates -powerpoint $powerpoint
    $slideTemplates

    # Add the first two slides to the PowerPoint
    Add-Slide -ppt $ppt -slide $SlideTemplates.'course' -title $yamldict.course
    Add-Slide -ppt $ppt -slide $SlideTemplates.'split' -title "Table of Contents" -chapter $yamldict.course -subchapter "TOC" 

    # Run through the dictionary and the required slides
    $yamldict.chapters | ForEach-Object {
         $chapter = $_
         Add-Slide -ppt $ppt -slide $SlideTemplates.'chapter' -title "Vocabulary" -chapter $chapter
         $chapter.subchapters | ForEach-Object {
             $subchapter = $_
             $subchapter.slides | ForEach-Object {
                 $slide = $_
                 Add-Slide -ppt $ppt -slide $SlideTemplates[$slide.type] -title $slide.title -chapter $chapter.title -subchapter $subchapter.title 
            }
        }
     }

     Add-Slide -ppt $ppt -slide $SlideTemplates.'question' -title "Knowledge Check"

     $count = 1 
     while ($count -le 6) {
         $ID = Get-SlideID -ppt $ppt -index 1
         Remove-Slide -ppt $ppt -id $ID
         $count = $count + 1
     }

     
     # Save the PowerPoint presentation as the course title to the wrk directory
     Save-Presentation -ppt $ppt
}
pps-generate