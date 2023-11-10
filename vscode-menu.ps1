# Params
param(
  [string]$menuName = "Open with VSCode"
)

function File-Exists {
  param(
    [string]$filePath
  )
  Test-Path -Path $filePath -PathType Leaf
}

Write-Output "Context menu name: $menuName"
Write-Output "If you want other names, Please ^C and pass first parameter(with quote) as name."
Start-Sleep -Seconds 3

# cd to ps script
Set-Location -Path $PSScriptRoot

Write-Output $PSScriptRoot

$dataPath = "./data"
# Path of 7-zip Minimal
$szipRPath = "$dataPath/7zr.exe"
# Path of 7-zip Extra zip
$szipExtraPath = "$dataPath/7zextra.7z"
# Path of 7-zip Command
$szipADirectory = "$($dataPath)/7za"
$szipAPath = "$szipADirectory/x64/7za.exe"
# Path of code-explorer zip
$expCommandPath = "$dataPath/code_explorer.zip"

# Path of code_explorer extracted
$extractedCEPath = "$dataPath/codeExplorer" 

# Create a web client object
$webClient = New-Object System.Net.WebClient

# Check directory exists
if (!(File-Exists $dataPath)) {
  New-Item -ItemType Directory -Path $dataPath
}

if (!(File-Exists $szipRPath)) {
  Write-Output "Downloading 7-zip Minimal..."
  $webClient.DownloadFile("https://7-zip.org/a/7zr.exe", $szipRPath)
}


if (!(File-Exists $szipExtraPath)) {
  Write-Output "Downloading 7-zip Extra..."
  $webClient.DownloadFile("https://www.7-zip.org/a/7z2301-extra.7z", $szipExtraPath)
}

if (!(File-Exists $expCommandPath)) {
  Write-Output "Downloading code_explorer..."
  $webClient.DownloadFile("https://github.com/microsoft/vscode-explorer-command/releases/latest/download/code_explorer_x64.zip", $expCommandPath)
}

if (!(File-Exists $szipAPath)) {
  Write-Output "Extracting 7-zip command..."
  Start-Process -FilePath $szipRPath -NoNewWindow -Wait -WorkingDirectory "." -ArgumentList "x", "-y", $szipExtraPath, "-o$szipADirectory"
}

# Simple function for extracting
function Extract-Zip {
  param (
    [string]$zipPath,
    [string]$extractPath
  )
  Start-Process -FilePath $szipAPath -NoNewWindow -Wait -WorkingDirectory "." -ArgumentList "x", "-y", $zipPath, "-o$extractPath"
}

# Extract code_explorer.zip
if (!(Test-Path -Path $extractedCEPath)) {
  Write-Output "Extracting CodeExplorer.."
  Extract-Zip $expCommandPath $extractedCEPath

  Write-Output "Extracting appx.."
  # Extract code_explorer_x64.appx
  Extract-Zip "$extractedCEPath/code_explorer_x64.appx" $extractedCEPath

  # Replace appxmanifest company
  $manifestPath = "$extractedCEPath/AppxManifest.xml"
  Write-Output "Replacing company name.."
  (Get-Content -Path $manifestPath).Replace("=Microsoft Corporation","=Hifumi Daisuki") | Set-Content $manifestPath
}

# Create context menu reg
$contextReg = "HKCU:/Software/Classes/VSCodeContextMenu"
$currentMenuName = Get-ItemPropertyValue -Path "$contextReg" -Name "Title"
if (($null -eq $currentMenuName) -or ($menuName -ne $currentMenuName)) {
  Write-Output "Registering Context menu as $menuName..."
  New-Item -Path $contextReg
  Set-ItemProperty -Path "$contextReg" -Name "Title" -Value $menuName
}

$vscodePath = Resolve-Path "$(
  Split-Path -parent (Get-Command code).Source
)/.."
# Check shell
$vscodeShellPath = "$vscodePath/shell"

Write-Output "Copying to $vscodeShellPath folder..."
if (File-Exists $vscodeShellPath) {
  Remove-Item $vscodeShellPath -y
}
New-Item -Path (Split-Path -parent $vscodeShellPath) -Name "VSCode-Shell" -ItemType "directory"
Copy-Item -Path "$extractedCEPath/*" -Destination $vscodeShellPath

# Check already installed
$menuAppx = Get-AppxPackage Microsoft.VSCode
if ($null -ne $menuAppx) {
  Write-Output "Removing $menuAppx..."
  Remove-AppxPackage $menuAppx.PackageFullName
}

# Install appx
Write-Output "Installing context menu..."
Add-AppxPackage -Path "$vscodeShellPath/AppxManifest.xml" -Register -ExternalLocation $vscodeShellPath
# $runCmd = "Add-AppxPackage -Path '$vscodeShellPath/AppxManifest.xml' -Register -ExternalLocation '$vscodeShellPath'"
# Start-Process "powershell" -verb runas -Wait -NoExit -ArgumentList "-command $runCmd"

Write-Output "Complete!"