# PowerPointShell
PowerShell library for interacting with PowerPoint from PowerShell -- Designed for Alta3 Research

## Installation
### Build from source:
1. Download and install Visual Studio IDE 2017
2. Clone the PowerPointShell GitHub repo to a destination of your choice
3. Open the **PowerPointShell.sln** in VS-2017
4. Compile the solution under **Release** mode (not debug mode)
5. Run the **buildpkg.ps1** script
6. Change directories to the newly created Build directory
7. Run the **install.ps1** script within this directory
8. In a new PowerShell console type **A3-Help** and press enter to see if the installation was successful

### Download & Install Pre-built Package
1. Download the **build.zip** file and unzip to a location of your choice
2. Run the **Install.ps1** script
3. In a new PowerShell console type **A3-Help** and press enter to see if the installation was succesful

## Command Help
### A3-Help
**Description:** Produces a list of commands currently available in the library as well as a short description of what the command is used for. To get more in-depth information about any given command in the library utilize the built in PowerShell help command (Get-Help <A3-COMMAND>) with any of the commands currently available in the PowerPointShell library.

**Usage:**
`PS> A3-Help`

### A3-Generate
**Description:** Takes in a YAML Outline file and produces a 

**Usage:** 
