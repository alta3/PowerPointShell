# Load Neccessary Assemblies


# PowerPoint Commands
function Add-Presentation {
    $powerpoint = New-Object -ComObject powerpoint.application 
    $ppt = $powerpoint.presentations.add(0)
    $ppt	
}
function Close-PowerPoint {

}

# Presentatoin Commands
function Close-Presentation {}
function Insert-Slide {
    param(
        [Parameter(Mandatory=$true)]
	[object]$ppt
	[int]$index
    )
}
function Save-Presentation {}

# Slide Commands
function Add-Note {}
function Add-Textbox {}
