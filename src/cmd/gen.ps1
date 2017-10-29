function pps-generate {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$filepath
    )
    # Dot Source the core library
    . "$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\lib\core.ps1"

    # PPS-GENERATE LOG STARTS
    Write-Host "-------------------------------------------------------------"
    Write-Host "PPS-GENERATE PROCESS LOG - START TIME: $(Get-Date)"
    Write-Host "-------------------------------------------------------------"

    #Ensure the $filepath exists and is in its long form
    Write-Host "Ensuring $filepath exists and is reachable."
    $newfilepath = Resolve-Path -Path $filepath | Select -ExpandProperty Path
    Write-Host "Utilizing $newfilepath as full path."

    # Convert YAML into useable dictionary
    Write-Host "Converting from YAML to usable dictionary."
    $yamlfile = Get-Content $newfilepath -Encoding Ascii | Out-String
    $yamldict = ConvertFrom-Yaml -yaml $yamlfile
    Write-Host "Conversion Complete!"

    # Start PowerPoint
    Write-Host "Starting PowerPoint"
    $powerpoint = Start-PowerPoint

    # Open Working PowerPoint
    Write-Host "Opening Blank Presentation"
    $ppt = Open-Presentation -powerpoint $powerpoint -path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell\src\mod\blank.pptm" -visable $false
    SaveAs-Presentation -ppt $ppt -name $yamldict.course
    $PresentationName = $ppt.name
    Write-Host ("Saved presentation as {0}" -f $PresentationName)

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
    Add-Section -ppt $ppt -index 1 -title $yamldict.course 

    # Initialize chapter tracking index
    $chaptercount = 1
    
    # Loop through each chapter
    $yamldict.chapters | ForEach-Object {
         $chapter = $_
         
         # Add the chapter slide with the title Vocabulary 
         Write-Host ("Adding chapter {0}: {1}" -f $chaptercount, $chapter.title)
         Add-Slide -ppt $ppt -slide $SlideTemplates.'chapter' -title "Vocabulary" -chapter $chapter.title
         
         # Add section title for this chapter
         Add-Section -ppt $ppt -index $ppt.slides.count -title $chapter.title
       
         # For each chapter add in the vocabulary words to the associated text file text file
         $chapter.vocab | ForEach-Object {
             $vocab = $_ 

             # Add vocabulary word to the dictionary of words
             if ($vocab) {Add-Vocab -word $vocab.keys -def $vocab.values -chapter $chaptercount}
         }

         # Loop through each subchapter
         $chapter.subchapters | ForEach-Object {
             $subchapter = $_
             
             # Check to see if the subchapter has defined slides
             if ($subchapter.slides.count -gt 0) {
                 # Loop through all the slides in each subchapter
                 $subchapter.slides | ForEach-Object {
                     $slide = $_

                     # If the slide has a type defined add the slide with that type else add a slide with a single title
                     if ($slide.type) {
                         Add-Slide -ppt $ppt -slide $SlideTemplates[$slide.type] -title $slide.title -chapter $chapter.title -subchapter $subchapter.title
                     } else {
                         Add-Slide -ppt $ppt -slide $SlideTemplates.'title' -title $slide.title -chapter $chapter.title -subchapter $subchapter.title
                     }
                     
                     # If the slide has notes add the notes to the current slide
                     if ($slide.notes) {Add-Note -ppt $ppt -index $ppt.slides.count -note $slide.notes} 
                }
            } else {
                Add-Slide -ppt $ppt -slide $SlideTemplates.'title' -title "SLIDE TITLE" -chapter $chapter.title -subchapter $subchapter.title
            } 

            # Loop through each question in the subchapter and add it to the appropriate text file 
            $subchapter.questions | ForEach-Object {
               $question = $_
               if ($question.incorrect.count -eq 3) {Add-Question -chapter $chapter.title -subchapter $subchapter.title -question $question.question -correct $question.correct -incortext1 $question.incorrect[0].text -incorexp1 $question.incorrect[0].explanation -incortext2 $question.incorrect[1].text -incorexp2 $question.incorrect[1].explanation -incortext3 $question.incorrect[2].text -incorexp3 $question.incorrect[2].explanation -value $question.value }
            }
        }
        $chaptercount = $chaptercount + 1 
        $chaptercount | Out-Null
     }

     # Add the slide that indicates the end of the slide deck
     Add-Slide -ppt $ppt -slide $SlideTemplates.'title' -title "End of Slide Deck" -Chapter $yamldict.course

     # Add the final question slide within its own section
     Add-Slide -ppt $ppt -slide $SlideTemplates.'question' -title "Knowledge Check"
     Add-Section -ppt $ppt -index $ppt.slides.count -title "Knowledge Check"

     # Generate GUIDs for each slide
     Generate-Guids -ppt $ppt

     # Merge the vocab words to a unique set of words
     Merge-Vocab

     # Save the PowerPoint presentation as the course title to the wrk directory
     Write-Host "Saving newly generated slide deck"
     Save-Presentation -ppt $ppt
     Close-Presentation -ppt $ppt

     # Open the newly Generated presentation so that it can be edited 
     Write-Host "Opening presentation for editing"
     $ppt = Open-Presentation -powerpoint $powerpoint -path "$HOME\Documents\Alta3 PowerPointShell\wrk\$PresentationName" -visable $True

     # PPS-GENERATE LOG ENDS
     Write-Host "-------------------------------------------------------------"
     Write-Host "PPS-GENERATE PROCESS LOG - END TIME: $(Get-Date)"
     Write-Host "-------------------------------------------------------------"
}
pps-generate -filepath C:\users\Michael\Documents\Github\PowerPointShell\src\mod\mod.yml