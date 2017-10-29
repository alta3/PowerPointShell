# Load Neccessary Assemblies
Add-Type -AssemblyName office
Add-type -AssemblyName microsoft.office.interop.powerpoint

function Start-PowerPoint {
    $powerpoint = New-Object -ComObject powerpoint.application
    $powerpoint	
}
function Close-PowerPoint {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object]$powerpoint
    )

    $powerpoint.quit()
    $powerpoint = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
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
function Close-Presentation {
    [CmdletBinding()]    
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object] $ppt
    )
    $ppt.close()
}
function Save-Presentation {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object] $ppt
    )
    $ppt.save()
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

function Add-Slide {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object]$ppt,
        [Parameter(Position=1,Mandatory=$true)]
        [int]$slide,
        [Parameter(Position=2,Mandatory=$true)]
        [string] $title,
        [Parameter(Mandatory=$false)]
        [string] $chapter,
        [string] $subchapter,
        [string] $note,
	    [int]$index
    )

    # Determine if the index should be at the end of the slide show or was specified by the user
    if (!$index) { $index = $ppt.slides.count() }

    # Insert the specified slide at the correct index
    $throwaway = $ppt.slides.insertfromfile("$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\mod\mod.pptm",$index,$slide,$slide)
    $CurrentSlide = $ppt.slides.item($index + 1)

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
function Add-Note {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object]$ppt,
        [Parameter(Position=1,Mandatory=$true)]
	    [int]$index,
        [Parameter(Position=2,Mandatory=$true)]
	    [string]$note
    )
	$slide = $ppt.slides.item($index)
    
    $slide.NotesPage.shapes("notes").textframe.textrange.text = $note

    #TO-DO
    # Figure out how to judge how much space will be utilized and either compensate with the text box 
    # or we are good with the way it generates a blank notes page? 
}

function Add-Section {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [object]$ppt,
        [Parameter(Position=1,Mandatory=$true)]
	    [int]$index,
        [Parameter(Position=2,Mandatory=$true)]
	    [string]$title
    )
    $ppt.SectionProperties.AddBeforeSlide($index, $title) | Out-Null
}
function Add-Vocab {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
	    [string]$word,
        [Parameter(Position=1,Mandatory=$true)]
	    [string]$def,
        [Parameter(Position=2,Mandatory=$true)]
	    [int]$chapter
    )

    $line = $word + " [" + $chapter + "] " + $def
     
    Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\vocab.txt" -Append -Encoding ascii -InputObject $line
}
function Add-Question {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
	    [string]$chapter,
        [Parameter(Mandatory=$true)]
	    [string]$subchapter,
        [Parameter(Mandatory=$true)]
        [string]$question,
        [Parameter(Mandatory=$true)]
	    [string]$correct,
        [Parameter(Mandatory=$true)]
        [string]$incortext1,
        [Parameter(Mandatory=$true)]
        [string]$incortext2,
        [Parameter(Mandatory=$true)]
        [string]$incortext3,
        [Parameter(Mandatory=$true)]
        [string]$incorexp1,
        [Parameter(Mandatory=$true)]
        [string]$incorexp2,
        [Parameter(Mandatory=$true)]
        [string]$incorexp3,
        [Parameter(Mandatory=$true)]
	    [int]$value
    )
    
    $id = -join ((33..126) | Get-Random -Count 16 | % {[char]$_})
    $num = Get-Random -Minimum 1 -Maximum 5 

    $incordict = @{
        1 = @{$incortext1 = $incorexp1};
        2 = @{$incortext2 = $incorexp2};
        3 = @{$incortext3 = $incorexp3}
    }

    $numdict = @{
        1 = "A";
        2 = "B";
        3 = "C";
        4 = "D"
    }

    Write-Output ("id: {0}" -f $id) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
    Write-Output ("chapsubchap: {0}: {1}" -f $chapter,$subchapter) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
    Write-Output "MediaURL:" | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii  
    Write-Output ("Points: {0}" -f $value) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
    Write-Output ("Question: {0}" -f $question) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
    
    $linecount = 1
    $incorrectcount = 1
    while ($linecount -le 4) {
        if ($linecount -eq $num) {
            Write-Output ("Choice{0}: {1}" -f $numdict.Item($linecount),$correct) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii  
        } else {
            $currentdict = $incordict.item($incorrectcount)
            $answer = $currentdict.keys | Out-String
            $answer = $answer.TrimEnd("`r?`n")
            Write-Output ("Choice{0}: {1}" -f $numdict.Item($linecount),$answer) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
            $incorrectcount = $incorrectcount + 1 
        }
        $linecount = $linecount + 1
    }

    $linecount = 1
    $incorrectcount = 1
    while ($linecount -le 4) {
        if ($linecount -eq $num) {
            Write-Output ("Correct{0}: 1" -f $numdict.Item($linecount)) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
        } else {
            Write-Output ("Correct{0}: 0" -f $numdict.Item($linecount)) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
            $incorrectcount = $incorrectcount + 1 
        }
        $linecount = $linecount + 1
    }

    $linecount = 1
    $incorrectcount = 1
    while ($linecount -le 4) {
        if ($linecount -eq $num) {
            Write-Output ("Why{0}: Correct!" -f $numdict.Item($linecount)) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii
        } else {
            $currentdict = $incordict.item($incorrectcount)
            $explanation = $currentdict.values | Out-String
            $explanation = $explanation.TrimEnd("`r?`n")
            Write-Output ("Why{0}: {1}" -f $numdict.Item($linecount),$explanation) | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii 
            $incorrectcount = $incorrectcount + 1 
        }
        $linecount = $linecount + 1
    }

    Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\quiz.txt" -Append -Encoding ascii -InputObject "------"
}

function Merge-Vocab {
    Get-Content "$HOME\Documents\Alta3 PowerPointShell\wrk\vocab.txt" | sort | Get-Unique | Out-File -FilePath "$HOME\Documents\Alta3 PowerPointShell\wrk\vocab.merge.txt" -Encoding ascii -Force
    Remove-Item "$HOME\Documents\Alta3 PowerPointShell\wrk\vocab.txt"
    Rename-Item -Path "$HOME\Documents\Alta3 PowerPointShell\wrk\vocab.merge.txt"  -NewName "vocab.txt"
}

function Generate-Guids {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [object]$ppt
    )
    
    $slideguids = @{}

    $ppt.slides | ForEach-Object {
        $slide = $_ 

        $slideguid = $slide.shapes("GUID").textframe.textrange.text
        if ($slideguid -ne "GUID") {
            $slideguids.Add($slide.slideindex, $slideguid)
        } else {
            $guid = New-Guid
            $slideguid = $guid.Guid
            while ($slideguids.ContainsValue($slideguid)) {
                $guid = New-Guid
                $slideguid = $guid.Guid
            } 
            $slideguids.Add($slide.slideindex, $slideguid)
            $slide.shapes("GUID").textframe.textrange.text = $slideguid
        }
    }

    $ppt.Slides | ForEach-Object {
        $slide = $_
        $slideguid = $slide.shapes("GUID").textframe.textrange.text
    }
}

function Clear-Errors {}