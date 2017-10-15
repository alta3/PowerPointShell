# Load Neccessary Assemblies


# PowerPoint Commands
function Close-PowerPoint {

}

function Start-PowerPoint {
    $powerpoint = New-Object -ComObject powerpoint.application
    $powerpoint	
}

# Presentation Commands
function Open-Presentation {
    [CmdletBinding()]    
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object] $powerpoint
	[Parameter(Position=1,Mandatory=$true)]
	[string] $path
	[Parameter(Position=2,Mandatory=$true)]
	[int] $visable
    ) 
    $ppt = $powerpoint.open($path,,,$visable)
    $ppt
}

function Insert-Slide {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	[object]$ppt
	[Parameter(Position=1,Mandatory=$true)]
	[ValidateSet("intro","toc","chapter","title","split","blank","question")]
	[string]$type
	[Parameter(Position=2,Mandatory=$true)]
	[int]$index
    )
    if ($type -eq "title") {
        $ppt.slides.addslide($index, ppLayoutTitle)
    } elseif ($type -eq "split") {
        $ppt.slides.addslide($index, ppLayoutTwoContent)
    } elseif ($type -eq "blank") {
        $ppt.slides.addslide($index, ppLayoutBlank)
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
    $ppt.SaveAs("../wrk/$name.pptm")
}

# Slide Commands
function Add-Note {}
function Add-Textbox {}
