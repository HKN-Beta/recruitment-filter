# HKN Recruitment Filter - All-in-One Setup and Run Script
# This script automates the entire process: installs dependencies, clones/updates the repo, sets up the Python environment, and runs the application.

# --- CONFIGURATION ---
$repoUrl = "https://github.com/HKN-Beta/recruitment-filter.git"
$activeBranch = "Fa2025"
$venvName = "hknRecruitmentEnv"
# --- END CONFIGURATION ---


# Function to check if a command exists
function Test-Command {
    param($CommandName)
    # Use where.exe for robust executable checking in PATH
    where.exe $CommandName 2>$null 1>$null
    return $LASTEXITCODE -eq 0
}

# Function to run a command and exit on failure
function Invoke-CommandOrExit {
    param(
        [string]$Command,
        [string]$Arguments,
        [string]$ErrorMessage,
        [switch]$ShowOutput
    )
    # Use Invoke-Expression to correctly handle commands with multiple arguments
    $FullCommand = "$Command $Arguments"
    if ($ShowOutput) {
        Invoke-Expression $FullCommand
    }
    else {
        Invoke-Expression $FullCommand > $null 2>&1
    }
    if ($LASTEXITCODE -ne 0) {
        Write-Host "`r`nERROR: $ErrorMessage" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# 1. Check and Install Dependencies if needed
Write-Host "Checking dependencies..." -NoNewline
$needsInstall = $false
if (-not (Test-Command "git")) {
    $needsInstall = $true
}
if (-not (Test-Command "python")) {
    $needsInstall = $true
}

if ($needsInstall) {
    Write-Host "`r$(' ' * 50)`rInstalling missing software..." -NoNewline
    
    # 1a. Check for Administrator Privileges
    if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "`r`nRequesting Administrator privileges to install software..." -ForegroundColor Yellow
        
        # Restart the script with Administrator privileges
        try {
            Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            exit
        }
        catch {
            Write-Host "ERROR: Failed to elevate privileges. Please run PowerShell as Administrator manually." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }

    # 1b. Check for Winget
    if (-not (Test-Command "winget")) {
        Write-Host "`r`nERROR: Winget is not available on this system." -ForegroundColor Red
        Write-Host "Please install The following to use this script:" -ForegroundColor Yellow
        if (-not (Test-Command "python")) {
            Write-Host "Python " -ForegroundColor Cyan
        }
        if (-not (Test-Command "git")) {
            Write-Host "Git" -ForegroundColor Cyan
        }
        Read-Host "Press Enter to exit"
        exit 1
    }

    # 1c. Install missing software
    if (-not (Test-Command "git")) {
        Invoke-CommandOrExit "winget" "install --id Git.Git -e --source winget" "Failed to install Git."
    }
    if (-not (Test-Command "python")) {
        Invoke-CommandOrExit "winget" "install --id Python.Python.3 -e --source winget" "Failed to install Python."
    }
    Write-Host "`r$(' ' * 50)`rDependencies installed successfully!" -ForegroundColor Green
}
else {
    Write-Host "`r$(' ' * 50)`rDependencies verified!" -ForegroundColor Green
}

# 4. Git Repository Setup
Write-Host "Setting up repository..." -NoNewline

# Check if we're already in a git repository
if (-not (Test-Path ".git")) {
    # Check if repository folder already exists
    $repoName = (Split-Path -Leaf $repoUrl) -replace '\.git$'
    $repoPath = Join-Path $PWD $repoName
    
    if (Test-Path $repoPath) {
        # Repository folder exists, navigate into it
        Write-Host "`r$(' ' * 50)`rUsing existing repository..." -NoNewline
        Set-Location $repoPath
        
        # Verify it's actually a git repository
        if (-not (Test-Path ".git")) {
            Write-Host "`r`nERROR: Directory '$repoName' exists but is not a git repository." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
    else {
        # Clone the repository into current directory
        Write-Host "`r$(' ' * 50)`rCloning repository..." -NoNewline
        Invoke-CommandOrExit "git" "clone --branch $activeBranch $repoUrl" "Failed to clone repository." -ShowOutput
        
        # Change to the cloned repository directory
        Set-Location $repoPath
    }
    
    # Verify we're in the correct directory
    if (-not (Test-Path "hkn_student_parser.py")) {
        Write-Host "`r`nERROR: Python script not found in repository. Check repository contents." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Update repository
Invoke-CommandOrExit "git" "fetch" "Failed to fetch from remote."
Invoke-CommandOrExit "git" "checkout $activeBranch" "Failed to checkout branch '$activeBranch'."
Invoke-CommandOrExit "git" "pull origin $activeBranch" "Failed to pull changes from '$activeBranch'."
Write-Host "`r$(' ' * 50)`rRepository ready!" -ForegroundColor Green

# 5. Python Virtual Environment Setup
Write-Host "Setting up Python environment..." -NoNewline
if (-not (Test-Path $venvName)) {
    Invoke-CommandOrExit "python" "-m venv $venvName" "Failed to create virtual environment."
}

# Activate virtual environment
. ".\$venvName\Scripts\Activate.ps1"
Write-Host "`r$(' ' * 50)`rPython environment ready!" -ForegroundColor Green

# 6. Install Dependencies
Write-Host "Installing packages..." -NoNewline
Invoke-CommandOrExit "python" "-m pip install --upgrade pip" "Failed to upgrade pip."
Invoke-CommandOrExit "pip" "install -r requirements.txt" "Failed to install required packages from requirements.txt."
Write-Host "`r$(' ' * 50)`rPackages installed!" -ForegroundColor Green

# 6a Make an executable shortcut for the script
# Create a shortcut to this script
$shortcutPath = Join-Path $PWD "HKN Recruitment Filter.lnk"
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
$Shortcut.WorkingDirectory = $PWD.Path
$Shortcut.Save()
Write-Host "`r$(' ' * 50)`rShortcut created!" -ForegroundColor Green

# 7. Run the Python Script
Write-Host "Running recruitment filter..." -NoNewline
Write-Host ""  # Move to next line for Python output
Invoke-CommandOrExit "python" "hkn_student_parser.py" "Python script execution failed. Check the output above for errors." -ShowOutput
Write-Host "Recruitment filter completed!" -ForegroundColor Green

Write-Host "`nGenerated files: seniors.csv, juniors.csv, sophomores.csv, report.txt" -ForegroundColor Cyan
Write-Host "Check report.txt for detailed summary." -ForegroundColor Cyan

# 8. Deactivate virtual environment
deactivate
Write-Host "`r$(' ' * 50)`rVirtual environment deactivated!" -Foreground
Write-Host "`nAll tasks completed successfully!" -ForegroundColor Green
Read-Host "Press Enter to exit"