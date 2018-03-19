param(
    $markdownPath,
    $latexPath,
    $activeGuid
)

Set-Location $latexPath
pandoc -f markdown -t latex -o out.tex $markdownPath\$activeGuid.md 