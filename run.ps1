# HKN Recruitment Filter - PowerShell Run Script
# This script activates virtual environment, installs/updates packages, and runs the recruitment filter

Write-Host "========================================"
Write-Host "HKN Recruitment Filter Setup and Run"
Write-Host "========================================"
Write-Host

# Function to check if a command exists
function Test-Command {
    param($CommandName)
    try {
        Get-Command $CommandName -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}


# Check if Python is installed
Write-Host "Checking Python installation..."
if (-not (Test-Command "python")) {
    Write-Error "ERROR: Python is not installed or not in PATH"
    Write-Host "Please install Python from https://python.org and ensure it's added to PATH"
    Write-Host
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Python detected:"
python --version
Write-Host

# Check if we're in a git repository and pull latest changes
Write-Host "Checking for git repository..."
if (Test-Path ".git") {
    Write-Host "Git repository detected. Checking for updates..."
    
    # Check if git is installed
    if (Test-Command "git") {
        Write-Host "Pulling latest changes from repository..."
        git pull
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Repository updated successfully!" -ForegroundColor Green
        } else {
            Write-Warning "WARNING: Failed to pull latest changes from git repository"
            Write-Host "You may need to resolve conflicts manually or check your internet connection"
            Write-Host "Continuing with current files..."
        }
    } else {
        Write-Warning "WARNING: Git is not installed or not in PATH"
        Write-Host "Skipping repository update. Install Git from https://git-scm.com/ for automatic updates"
    }
} else {
    Write-Host "Not a git repository. Skipping update check."
}
Write-Host

# Check if virtual environment exists, create if needed
if (-not (Test-Path "hknRecruitmentEnv\Scripts\activate.ps1")) {
    Write-Host "Virtual environment not found. Creating one..."
    python -m venv hknRecruitmentEnv
    if ($LASTEXITCODE -ne 0) {
        Write-Error "ERROR: Failed to create virtual environment"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Activate virtual environment
Write-Host "Activating virtual environment..."
& ".\hknRecruitmentEnv\Scripts\Activate.ps1"
if ($LASTEXITCODE -ne 0) {
    Write-Error "ERROR: Failed to activate virtual environment"
    Write-Host "You may need to run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if pip is available
if (-not (Test-Command "pip")) {
    Write-Error "ERROR: pip is not installed or not available in virtual environment"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Installing/updating required Python packages..."
Write-Host "========================================"

# Upgrade pip first
Write-Host "Upgrading pip..."
python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    Write-Error "ERROR: Failed to upgrade pip"
    Write-Host "Please check your internet connection and try again"
    Read-Host "Press Enter to exit"
    exit 1
}

# Install/upgrade requirements
Write-Host "Installing/upgrading packages from requirements.txt..."
pip install --upgrade -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error "ERROR: Failed to install required packages"
    Write-Host "Please check your internet connection and try again"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host
Write-Host "Package installation completed successfully!" -ForegroundColor Green
Write-Host

Write-Host "Running HKN Recruitment Filter..."
Write-Host "========================================"
Write-Host

# Run the main Python script
python hkn_student_parser.py
if ($LASTEXITCODE -ne 0) {
    Write-Error "ERROR: Script execution failed"
    Write-Host "Please check the error messages above"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "Script completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host

Write-Host "Generated files:"
Write-Host "- seniors.csv" -ForegroundColor Cyan
Write-Host "- juniors.csv" -ForegroundColor Cyan
Write-Host "- sophomores.csv" -ForegroundColor Cyan
Write-Host "- report.txt" -ForegroundColor Cyan
Write-Host

Write-Host "Check the report.txt file for detailed statistics and any errors."
Write-Host

# Deactivate virtual environment
deactivate