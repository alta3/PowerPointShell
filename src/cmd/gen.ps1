function pps-generate {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$filepath
    )
    # Dot Source the core library
    . "$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\lib\core.ps1"

    #Ensure the $filepath exists and is in its long form
    Write-Host "Ensuring $filepath exists and is reachable."
    $newfilepath = Resolve-Path -Path $filepath | Select -ExpandProperty Path
    Write-Host "Utilizing $newfilepath as full path."

    # Convert YAML into useable dictionary
    Write-Host "Converting from YAML to usable dictionary."
    $yamlfile = Get-Content $newfilepath | Out-String
    $yamldict = ConvertFrom-Yaml -yaml $yamlfile
    Write-Host "Conversion Complete!"

    # Start PowerPoint
    Write-Host "Starting PowerPoint"
    $powerpoint = Start-PowerPoint

    # Open Working PowerPoint
    Write-Host "Opening Blank Presentation"
    $ppt = Open-Presentation -powerpoint $powerpoint -path "$home\documents\alta3 powerpointshell\src\mod\blank.pptm" -visable $false
    SaveAs-Presentation -ppt $ppt -name $yamldict.course
    $PresentationName = $ppt.name
    Write-Host "Saved presentation as " + $PresentationName

    # Set Slide Template Indexes
    $SlideTemplates = @{ 
        "course" = 1; 
        "chapter" = 2;
        "title" = 3;
        "split" = 4;
        "blank" = 5;
        "question" = 6;
    } 

    # Add the first two slides to the PowerPoint
    Write-Host "Adding Course Title & TOC slides"
    Add-Slide -ppt $ppt -slide $SlideTemplates.'course' -title $yamldict.course
    Add-Slide -ppt $ppt -slide $SlideTemplates.'split' -title "Table of Contents" -chapter $yamldict.course -subchapter "TOC" 

    # Initialize chapter tracking index
    $chaptercount = 1
    
    # Loop through each chapter
    $yamldict.chapters | ForEach-Object {
         $chapter = $_
         
         # Add the chapter slide with the title Vocabulary 
         Write-Host "Adding chapter " + $chaptercount + ": " + $chapter.title
         Add-Slide -ppt $ppt -slide $SlideTemplates.'chapter' -title "Vocabulary" -chapter $chapter.title
         
         # If there are notes for the chapter slide include them here
         if ($chapter.note) {Add-Note -ppt $ppt -type $SlideTemplates[$slide.type] -index $ppt.count -note $chapter.note}
         
         # For each chapter add in the vocabulary words to the associated text file text file
         $chapter.vocab | ForEach-Object {
             $vocab = $_ 

             # Add vocabulary word to the dictionary of words
             Add-Vocab -name $PresentationName -word $vocab.keys -def $vocab.values -chapter $chaptercount
         }

         # Loop through each subchapter
         $chapter.subchapters | ForEach-Object {
             $subchapter = $_
             
             # Loop through all the slides in each subchapter
             $subchapter.slides | ForEach-Object {
                 $slide = $_
                 # If the slide exists add a slide to the presentation
                 if ($slide) {Add-Slide -ppt $ppt -slide $SlideTemplates[$slide.type] -title $slide.title -chapter $chapter.title -subchapter $subchapter.title}
                 
                 # If the slide has notes add the notes to the current slide
                 if ($slide.note) {Add-Note -ppt $ppt -type $SlideTemplates[$slide.type] -index $ppt.count -note $slide.note} 
            }

            # Loop through each question in the subchapter and add it to the appropriate text file 
            $subchapters.questions | ForEach-Object {
                $question = $_
                Add-Question -name $PresentationName -question $question.question -correct $question.correct -incortext1 $question.incorrect[0].text -incorexp1 $question.incorrect[0].explanation -incortext2 $question.incorrect[1].text -incorexp2 $question.incorrect[1].explanation-incortext3 $question.incorrect[3].text -incorexp3 $question.incorrect[2].explanation -value $question.value
            }
        }
        $chaptercount = $chaptercount + 1 
     }

     # Add the slide that indicates the end of the slide deck
     Add-Slide -ppt $ppt -slide $SlideTemplates.'title' -title "End of Slide Deck" -Chapter $yamldict.course

     # Add the final question slide
     Add-Slide -ppt $ppt -slide $SlideTemplates.'question' -title "Knowledge Check"
    
     # Save the PowerPoint presentation as the course title to the wrk directory
     Write-Host "Saving newly generated slide deck"
     Save-Presentation -ppt $ppt
     Close-Presentation -ppt $ppt

     Write-Host "Opening presentation for editing"
     $ppt = Open-Presentation -powerpoint $powerpoint -path "$HOME\Documents\Alta3 PowerPointShell\wrk\$PresentationName" -visable $True
}
pps-generate -filepath C:\users\Michael\Documents\Github\PowerPointShell\src\mod\mod.yml