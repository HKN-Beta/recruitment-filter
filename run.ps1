# HKN Recruitment Filter - All-in-One Setup and Run Script
# This script automates the entire process: installs dependencies, clones/updates the repo, sets up the Python environment, and runs the application.

# --- CONFIGURATION ---
$repoUrl = "https://github.com/HKN-Beta/recruitment-filter.git"
$activeBranch = "Fa2025"
$venvName = "hknRecruitmentEnv"
# --- END CONFIGURATION ---

Write-Host "Works3"



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

    # 1c. Install missing software in parallel
    $jobs = @()
    
    if (-not (Test-Command "git")) {
        $jobs += Start-Job -ScriptBlock {
            winget install --id Git.Git -e --source winget
            return @{Name = "Git"; ExitCode = $LASTEXITCODE }
        }
    }
    
    if (-not (Test-Command "python")) {
        $jobs += Start-Job -ScriptBlock {
            winget install --id Python.Python.3 -e --source winget
            return @{Name = "Python"; ExitCode = $LASTEXITCODE }
        }
    }
    
    # Wait for all installations to complete
    if ($jobs.Count -gt 0) {
        $results = $jobs | Wait-Job | Receive-Job
        $jobs | Remove-Job
        
        # Check if any installations failed
        foreach ($result in $results) {
            if ($result.ExitCode -ne 0) {
                Write-Host "`r`nERROR: Failed to install $($result.Name)." -ForegroundColor Red
                Read-Host "Press Enter to exit"
                exit 1
            }
        }
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

# Update repository with optimized git operations
# First, handle potential git ownership issues
$currentDir = $PWD.Path
git config --global --add safe.directory $currentDir > $null 2>&1

# Check if we need to fetch/checkout at all
$currentBranch = git rev-parse --abbrev-ref HEAD 2>$null
$needsUpdate = $false

if ($currentBranch -ne $activeBranch) {
    $needsUpdate = $true
}

if ($needsUpdate) {
    # Only run git operations if needed
    try {
        Invoke-CommandOrExit "git" "fetch origin" "Failed to fetch from remote."
        Invoke-CommandOrExit "git" "checkout $activeBranch" "Failed to checkout branch '$activeBranch'."
    }
    catch {
        Write-Host "`r$(' ' * 50)`rWarning: Git operations failed, continuing with existing code..." -ForegroundColor Yellow
    }
}
else {
    # Always try to pull latest changes and check if there was an update
    $beforeCommit = git rev-parse HEAD 2>$null
    git pull origin $activeBranch > $null 2>&1
    $afterCommit = git rev-parse HEAD 2>$null
    
    # If the commit hash changed, there was an update
    if ($beforeCommit -ne $afterCommit) {
        Write-Host "`r`nScript updated! Please rerun the script to use the latest version." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 0
    }
}

Write-Host "`r$(' ' * 50)`rRepository ready!" -ForegroundColor Green

# 5. Install uv and Setup Python Virtual Environment
Write-Host "Setting up Python environment..." -NoNewline

# Check if uv is already installed globally or install it
if (-not (Test-Command "uv")) {
    Write-Host "`r$(' ' * 50)`rInstalling uv globally..." -NoNewline
    Invoke-CommandOrExit "python" "-m pip install uv --disable-pip-version-check --quiet" "Failed to install uv globally."
}

# Use uv to create virtual environment (much faster than python -m venv)
if (-not (Test-Path $venvName)) {
    Write-Host "`r$(' ' * 50)`rCreating virtual environment with uv..." -NoNewline
    Invoke-CommandOrExit "uv" "venv $venvName" "Failed to create virtual environment with uv."
}

# Activate virtual environment
. ".\$venvName\Scripts\Activate.ps1"
Write-Host "`r$(' ' * 50)`rPython environment ready!" -ForegroundColor Green

# 6. Install Dependencies with uv
Write-Host "Installing packages..." -NoNewline

# Read requirements.txt and check if it exists
$requirementsPath = Join-Path $PWD "requirements.txt"
if (-not (Test-Path $requirementsPath)) {
    Write-Host "`r`nERROR: requirements.txt not found." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Use uv to install packages from requirements.txt - much faster than pip!
Write-Host "`r$(' ' * 50)`rInstalling packages with uv..." -NoNewline
Invoke-CommandOrExit "uv" "pip install -r requirements.txt" "Failed to install packages with uv."

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
Invoke-CommandOrExit "python" "hkn_student_parser.py --mp --lm" "Python script execution failed. Check the output above for errors." -ShowOutput
Write-Host "Recruitment filter completed!" -ForegroundColor Green

# 7a. Generate Emails
Write-Host "Generating emails..." -NoNewline
Invoke-CommandOrExit "python" "generateEmails.py" "Failed to generate emails." -ShowOutput
Write-Host "Emails generated successfully!" -ForegroundColor Green

Write-Host "`nGenerated files: seniors.csv, juniors.csv, sophomores.csv, report.log, emails.csv" -ForegroundColor Cyan
Write-Host "Check report.log for detailed summary." -ForegroundColor Cyan

# 8. Deactivate virtual environment
deactivate
Write-Host "`r$(' ' * 50)`rVirtual environment deactivated!" -ForegroundColor Green
Write-Host "`nAll tasks completed successfully!" -ForegroundColor Green
Read-Host "Press Enter to exit"