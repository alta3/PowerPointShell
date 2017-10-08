# PowerPoint Commands
function Add-Presentation {
    $powerpoint = New-Object -ComObject powerpoint.application 
    $ppt = powerpoint.presentations.add(0)

    write-object $ppt	
}

# Presentatoin Commands
function Add-Slide {}

# Slide Commands
function Add-Note {}
function Add-Textbox {}
