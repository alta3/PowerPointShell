# Load Neccessary Assemblies
Add-Type -AssemblyName office
Add-type -AssemblyName microsoft.office.interop.powerpoint

# PowerPoint Commands
function Close-PowerPoint {

}

function Start-PowerPoint {
    $powerpoint = New-Object -ComObject powerpoint.application
    $powerpoint	
}

# Presentation Commands
function Close-Presentation {
    [CmdletBinding()]    
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object] $ppt
    )
    $ppt.close()
}

function Open-Presentation {
    [CmdletBinding()]    
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object] $powerpoint,
	    [Parameter(Position=1,Mandatory=$true)]
	    [string] $path,
	    [Parameter(Position=2,Mandatory=$true)]
	    [bool] $visable
    )
    if ($visable) {$vis = 1} else {$vis = 0} 
    $ppt = $powerpoint.Presentations.open($path,0,0,$vis)
    $ppt
}

function Get-SlideTemplates {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object]$powerpoint
    )

    $ppt = Open-Presentation -powerpoint $powerpoint -path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\mod\mod.pptm" -visable $false

    $CourseTitleSlide = $ppt.slides.item(1)
    $ChapterSlide = $ppt.slides.item(2)
    $TitleSlide = $ppt.slides.item(3)
    $SplitSlide = $ppt.slides.item(4)
    $BlankSlide = $ppt.slides.item(5)
    $QuestionSlide = $ppt.slides.item(6)

    $SlideTemplates = @{ 
         "course" = $CourseTitleSlide; 
         "chapter" = $ChapterSlide;
         "title" = $TitleSlide;
         "split" = $SplitSlide;
         "blank" = $BlankSlide;
         "question" = $QuestionSlide;
    }
    $SlideTemplates, $ppt
}

function Add-Slide {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object]$ppt,
        [Parameter(Position=1,Mandatory=$true)]
        [object]$slide,
        [Parameter(Position=2,Mandatory=$true)]
        [string] $title,
        [Parameter(Mandatory=$false)]
        [string] $chapter,
        [string] $subchapter,
        [string] $note,
        [string] $text_content,
        [string] $content_path,
	    [int]$index
    )

    # Copy the sample slide either to the end of the presenation or at the specified index
    if ($index) {
        $slide.copy()
        $CurrentSlide = $ppt.slides.paste($index)
    } else {
        $slide.copy()
        $CurrentSlide = $ppt.slides.paste()
    }

    # Set the title of the slide
    if ($CurrentSlide.shapes("title")) {
        $CurrentSlide.shapes("title").textframe.textrange.text = "$title"
    }
    
    # Figure out what the SCRUBBER line should look like
    if ($chapter -and $subchapter) {
        $SCRUBBER = $chapter + ": " + $subchapter
    } else {
        $SCRUBBER = $chapter
    }
    
    # If this slide has a scrubber spot add the scrubber line
    if ($SCRUBBER -and $CurrentSlide.shapes("SCRUBBER")) {
        $CurrentSlide.shapes("SCRUBBER").textframe.textrange.text = "$SCRUBBER"     
    }
}

function SaveAs-Presentation {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object] $ppt,
	    [Parameter(Position=1,Mandatory=$true)]
	    [string] $name
    )
    if (Test-Path "$HOME/Documents/Alta3 PowerPointShell/wrk/$name.pptm") {
        $count = 0 
        $newname = $name + $count
        while (Test-Path "$HOME/Documents/Alta3 PowerPointShell/wrk/$newname.pptm") {
             $count = $count + 1 
             $newname = $name + $count
        }
        $ppt.SaveAs("$HOME/Documents/Alta3 PowerPointShell/wrk/$newname.pptm")
    } else {
        $ppt.SaveAs("$HOME/Documents/Alta3 PowerPointShell/wrk/$name.pptm")
    }
}

function Save-Presentation {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object] $ppt
    )
    $ppt.save()
}

# Slide Commands
function Add-Note {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object]$ppt
    )	
	
}

