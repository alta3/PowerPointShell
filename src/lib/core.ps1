# Load Neccessary Assemblies


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
    $ppt.save()
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
    $ppt = $powerpoint.open($path,,,$vis)
    $ppt
}

function Get-SlideTemplates {
    [CmdletBinding()]
    param(
    [Parameter(Position=0,Mandatory=$true)]
	[object]$powerpoint
    )

    $ppt = Open-Presentation -powerpoint $powerpoint -path "$env:PSModulePath\PowerPointShell\src\mod\mod.pptm" -visable $false

    $CourseTitleSlide = $ppt.slides.item(1)
    $ChapterSlide = $ppt.slides.item(3)
    $TitleSlide = $ppt.slides.item(4)
    $SplitSlide = $ppt.slides.item(5)
    $BlankSlide = $ptt.slides.item(6)
    $QuestionSlide = $ppt.slides.item(7)

    $SlideLayouts = @{ 
         "course" = $CourseTitleLayout; 
         "chapter" = $ChapterLayout;
         "title" = $TitleLayout;
         "split" = $SplitLayout;
         "blank" = $BlankLayout;
         "question" = $QuestionLayout;
    }

    Close-Presentation -ppt $ppt

    $SlideLayouts
}

function Add-Slide {
    [CmdletBinding()]
    param(
    [Parameter(Position=0,Mandatory=$true)]
	[object]$ppt,
    [Parameter(Position=1,Mandatory=$true)]
	[ValidateSet("intro","toc","chapter","title","split","blank","question")]
	[string]$type,
	[Parameter(Position=2,Mandatory=$true)]
    [object] $layout,
    [Parameter(Position=3,Mandatory=$true)]
    [string] $title,
    [Parameter(Mandatory=$false)]
    [string] $chapter,
    [Parameter(Mandatory=$false)]
    [string] $subchapter
    )

    if ($type -eq "Intro") {
        $slide = $ppt.slides.addslide($index, $layout)
        $slide.shapes.item("Title 1").TextFrame.TextRange.Text = $title
    } elseif ($type -eq "chapter") {
        $slide = $ppt.slides.addslide($index, $layout)
        $slide.shapes.item()
    } else if ($type -eq "title") {
        $slide = $ppt.slides.addslide($index, $layout)
    } else if ($type -eq "split") {
        $slide = $ppt.slides.addslide($index, $layout)
    } else if ($type -eq "blank") {
        $slide = $ppt.slides.addslide($index, $layout)
    } else if ($type -eq "question") {
        $slide = $ppt.slides.addslide($index, $layout)
    } else {
        write-host "Something went very wrong!"
    }
}

function Save-Presentation {
    [CmdletBinding()]
    param(
    [Parameter(Position=0,Mandatory=$true)]
	[object] $ppt
	[Parameter(Position=1,Mandatory=$true)]
	[string] $name
    )
    $ppt.SaveAs("$HOME/Documents/Alta3 PowerPointShell/wrk/$name.pptm")
}

# Slide Commands
function Add-Note {}

