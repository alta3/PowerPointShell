param(
    $root,
    $name
)
Set-Location $root
pdflatex main.tex 
move-item .\main.pdf ..\$name.pdf