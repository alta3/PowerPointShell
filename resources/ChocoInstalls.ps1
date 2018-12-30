Set-ExecutionPolicy Bypass -Scope Process -Force; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
choco install miktex --force -y
choco install pandoc --force -y 
choco install texstudio --force -y
choco install python --force -y 
pip install yamllint