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
    $ppt = Open-Presentation -powerpoint $powerpoint -path "$PSScriptRoot\..\mod\mod.pptm" -visable $false
    SaveAs-Presentation -ppt $ppt -name $yamldict.course

    # Get the slide templates & a handle to the model_ppt that is open as well
    $SlideTemplates, $mod_ppt = Get-SlideTemplates -powerpoint $powerpoint

    # Add the first two slides to the PowerPoint
    Add-Slide -ppt $ppt -slide $SlideTemplates.'course' -title $yamldict.course
    Add-Slide -ppt $ppt -slide $SlideTemplates.'split' -title "Table of Contents" -chapter $yamldict.course -subchapter "TOC" 

    # Run through the dictionary and the required slides
    $yamldict.chapters | ForEach-Object {
         $chapter = $_
         Add-Slide -ppt $ppt -slide $SlideTemplates.'chapter' -title "Vocabulary" -chapter $chapter.title
         $chapter.subchapters | ForEach-Object {
             $subchapter = $_
             $subchapter.slides | ForEach-Object {
                 $slide = $_
                 if ($slide) {Add-Slide -ppt $ppt -slide $SlideTemplates[$slide.type] -title $slide.title -chapter $chapter.title -subchapter $subchapter.title} 
            }
        }
     }

     Add-Slide -ppt $ppt -slide $SlideTemplates.'question' -title "Knowledge Check"
    
     $ppt.slides.item(6).delete()
     $ppt.slides.item(5).delete()
     $ppt.slides.item(4).delete()
     $ppt.slides.item(3).delete()
     $ppt.slides.item(2).delete()
     $ppt.slides.item(1).delete()
   
     # Save the PowerPoint presentation as the course title to the wrk directory
     Close-Presentation -ppt $mod_ppt
     $PresentationName = $ppt.name
     Save-Presentation -ppt $ppt
     Close-Presentation -ppt $ppt

     $ppt = Open-Presentation -powerpoint $powerpoint -path "$HOME\Documents\Alta3 PowerPointShell\wrk\$PresentationName" -visable $True
}
pps-generate -filepath .\src\mod\mod.yml