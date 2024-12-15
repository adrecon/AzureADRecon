# Function to check and install requirements
function Install-Requirements {
  try {
      # Example: Installation of the AzureAD module
      if (!(Get-Module -ListAvailable -Name "AzureAD")) {
          Write-Host "Installing AzureAD module..."
          Install-Module AzureAD -Scope CurrentUser -Force -ErrorAction Stop
          Write-Host "AzureAD module successfully installed."
      } else {
          Write-Host "AzureAD module is already installed."
      }
    

  } catch {
      Write-Error "Error during requirements installation: $($_.Exception.Message)"
      exit 1
  }
}

# Function to download the script from GitHub
function Download-Script {
  $DownloadUrl = "https://raw.githubusercontent.com/adrecon/AzureADRecon/master/AzureADRecon.ps1"

  try {
      Write-Host "Downloading AzureADRecon script..."
      $directoryPath = [System.IO.Path]::GetDirectoryName($ScriptPath)
      
      if (!(Test-Path $directoryPath)) {
          Write-Host "Directory does not exist. Creating directory: $directoryPath"
          New-Item -ItemType Directory -Path $directoryPath -Force
      }
      
      if (Test-Path $ScriptPath) {
          Write-Host "Existing script will be overwritten..."
      }
      Invoke-WebRequest -Uri $DownloadUrl -OutFile $ScriptPath -UseBasicParsing -ErrorAction Stop
      Write-Host "Script saved to $ScriptPath."
  } catch {
      Write-Error "Error downloading the script: $($_.Exception.Message)"
      exit 1
  }
}

# User-friendly prompts at the beginning
Write-Host "Welcome to the AzureADRecon script installation."

# Ask if requirements should be installed automatically
$installRequirements = Read-Host "Do you want the requirements to be installed automatically? (Yes/No) [Default: Yes]"
if ([string]::IsNullOrWhiteSpace($installRequirements) -or $installRequirements.ToUpper() -in @("YES", "Y")) {
  Install-Requirements
}

# Prompt for the script installation path
$defaultScriptPath = "C:\Scripts\AzureADRecon.ps1"
$ScriptPath = Read-Host "Please specify the installation path for the script (Default: $defaultScriptPath)"
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
  $ScriptPath = $defaultScriptPath
}

# Ask if the script should be executed directly after installation
$runAfterInstall = Read-Host "Do you want to run the script immediately after installation? (Yes/No) [Default: Yes]"
if ([string]::IsNullOrWhiteSpace($runAfterInstall)) {
  $runAfterInstall = "Yes"
}

# Main logic
try {
  Write-Host "Starting the installation process..."
  
  # Download the script
  Download-Script
  Write-Host "Installation complete."

  if ($runAfterInstall.ToUpper() -in @("YES", "Y")) {
      Write-Host "Starting the AzureADRecon script..."
      
      # Check if the script exists
      if (Test-Path $ScriptPath) {
          try {
              & $ScriptPath -ErrorAction Stop
          } catch {
              Write-Error "Error starting the AzureADRecon script: $($_.Exception.Message)"
              exit 1
          }
      } else {
          Write-Error "The script '$ScriptPath' was not found. Please check the specified installation path."
          exit 1
      }
  } else {
      Write-Host "The script was not started."
  }
} catch {
  Write-Error "General error: $($_.Exception.Message)"
  exit 1
}
