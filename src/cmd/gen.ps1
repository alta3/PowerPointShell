[CmdletBinding()]
param(
    [string]$filepath
)

$yamlfile = Get-Content $filepath | Out-String
$yamldict = ConvertFrom-Yaml -yaml $yamlfile

# Start PowerPoint and open it invisibly as a template presentation
$powerpoint = Start-PowerPoint
$ppt = Open-Presentation $powerpoint "$PSScriptRoot\..\mod\mod.pptm" 0

# Add the first two slides to the PowerPoint
Add-IntroSlide $ppt $yamldict.course
Add-TOCSlide $ppt $yamldict.chapters $yamldict.labs

$yamldict.chapters | foreach-object {
    $chapter = $_
    Add-ChapterSlide $ppt $chapter.title
    $chapter.subchapters | foreach-object {
        $subchapter = $_
        $subchapter.slides | ForEach-Object {
            $slide = $_
            Add-GenericSlide $ppt $chapter.title $subchapter.title $slide.title $slide.type
        }
    }
}

# Save the PowerPoint presentation as the course title to the wrk directory
Save-Presentation $ppt $yamldict.course
