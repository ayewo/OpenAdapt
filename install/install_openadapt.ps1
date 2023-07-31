# PowerShell script to pull OpenAdapt and install

################################   PARAMETERS   ################################
# Change these if a different version is required

$setupdir = "C:/OpenAdaptSetup"
$openAdaptURL = "https://github.com/OpenAdaptAI/OpenAdapt.git"
$openAdaptPath = "$env:USERPROFILE\OpenAdapt"

$tesseractCmd = "tesseract"
$tesseractInstaller = "tesseract.exe"
$tesseractInstallerURL = "https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.1.20230401.exe"
$tesseractPath = "C:\Program Files\Tesseract-OCR"

$pythonCmd = "python"
$pythonMinVersion = "3.10.0" # Change this if a different Lower version are supported by OpenAdapt
$pythonMaxVersion = "3.10.12" # Change this if a different Higher version are supported by OpenAdapt
$pythonInstaller = "python-3.10.11-amd64.exe"
$pythonInstallerURL = "https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
# $pythonPath = "C:\Program Files\Python310;C:\Program Files\Python310\Scripts"

$gitCmd = "git"
$gitInstaller = "Git-2.40.1-64-bit.exe"
$gitInstallerURL = "https://github.com/git-for-windows/git/releases/download/v2.40.1.windows.1/Git-2.40.1-64-bit.exe"
$gitPath = "C:\Program Files\Git\bin"
$gitUninstaller = "C:\Program Files\Git\unins000.exe"
################################   PARAMETERS   ################################


################################   FUNCTIONS    ################################
# Run a command and ensure it did not fail
function RunAndCheck {
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $Command,

        [Parameter(Mandatory = $true)]
        [string] $Desc,

        [Parameter(Mandatory = $false)]
        [switch] $SkipCleanup = $false
    )

    Invoke-Expression $Command
    if ($LastExitCode) {
        Write-Host "Failed: $Desc - Exit code: $LastExitCode" -ForegroundColor Red
        if (!$SkipCleanup) {
            Cleanup
            exit
        }
    }
    else {
        Write-Host "Success: $Desc" -ForegroundColor Green
    }
}

# Cleanup function to delete the setup directory
function Cleanup {
    $exists = Test-Path -Path $setupdir
    if ($exists) {
        Set-Location $env:USERPROFILE
        Remove-Item -LiteralPath $setupdir -Force -Recurse
    }
}


# Return true if a command/exe is available
function CheckCMDExists {
    Param
    (
        [Parameter(Mandatory = $true)] [string] $command
    )

    $result = Get-Command $command -errorvariable error -erroraction 'silentlycontinue'
    if ($null -eq $result) {
        # Write-Host "$error"
        return $false
    }
    return $true
}


# Return the current user's PATH variable
function GetUserPath {
    $userEnvPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    return $userEnvPath
}


# Return the system's PATH variable
function GetSystemPath {
    $systemEnvPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    return $systemEnvPath
}


# Refresh Path Environment Variable for both Curent User and System
function RefreshPathVariables {
    $env:Path = GetUserPath + ";" + GetSystemPath
}

function AddFolderToPathVariable {
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $FolderPath
    )

    # Add path to the System Path variable
    Write-Host "Adding $FolderPath to the System PATH environment variable."
    $systemEnvPath = GetSystemPath
    $updatedSystemPath = "$systemEnvPath;$FolderPath"
    [System.Environment]::SetEnvironmentVariable("Path", $updatedSystemPath, "Machine")

    # Add path to the User Path variable
    Write-Host "Adding $FolderPath to the User PATH environment variable."
    $userEnvPath = GetUserPath
    $updatedUserPath = "$userEnvPath;$FolderPath"
    [System.Environment]::SetEnvironmentVariable("Path", $updatedUserPath, "User")    
}


# Return true if a command/exe is available
function GetTesseractCMD {
    # Use tesseract alias if it exists
    if (CheckCMDExists $tesseractCmd) {
        return $tesseractCmd
    }

    # Check if tesseractPath exists and delete it if it does
    if (Test-Path -Path $tesseractPath -PathType Container) {
        Write-Host "Found Existing Old TesseractOCR, Deleting existing TesseractOCR folder"
        # Delete the whole folder
        Remove-Item $tesseractPath -Force -Recurse
    }

    # Downlaod Tesseract OCR
    Write-Host "Downloading Tesseract OCR installer"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $tesseractInstallerURL -OutFile $tesseractInstaller
    $exists = Test-Path -Path $tesseractInstaller -PathType Leaf
    if (!$exists) {
        Write-Host "Failed to download Tesseract OCR installer" -ForegroundColor Red
        Cleanup
        exit
    }

    # Install the Tesseract OCR Setup exe (binary file)
    Write-Host "Installing Tesseract OCR..."
    Start-Process -FilePath $tesseractInstaller -Verb runAs -ArgumentList "/S" -Wait
    Remove-Item $tesseractInstaller

    # Check if Tesseract OCR was installed
    if (Test-Path -Path $tesseractPath -PathType Container) {
        Write-Host "TesseractOCR installation successful." -ForegroundColor Green
    }
    else {
        Write-Host "TesseractOCR installation failed." -ForegroundColor Red
        Cleanup
        exit
    }

    AddFolderToPathVariable $tesseractPath
    Write-Host "Added Tesseract OCR to PATH." -ForegroundColor Green
    RefreshPathVariables

    # Make sure tesseract is now available
    if (CheckCMDExists($tesseractCmd)) {
        return $tesseractCmd
    }

    Write-Host "Error after installing Tesseract OCR."
    # Stop OpenAdapt install
    Cleanup
    exit
}


function ComparePythonVersion($version) {
    $v = [version]::new($version)
    $min = [version]::new($pythonMinVersion)
    $max = [version]::new($pythonMaxVersion)

    return $v -ge $min -and $v -le $max
}


# Check and Istall Python and return the python command
function GetPythonCMD {
    # Use python exe if it exists and is within the required version range
    if (CheckCMDExists $pythonCmd) {
        $res = Invoke-Expression "python -V"
        $versionString = $res.Split(' ')[-1]

        if (ComparePythonVersion $versionString) {
            return $pythonCmd
        }
    } 

    # Install required python version
    Write-Host "Downloading python installer..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $pythonInstallerURL -OutFile $pythonInstaller
    $exists = Test-Path -Path $pythonInstaller -PathType Leaf

    if (!$exists) {
        Write-Host "Failed to download python installer" -ForegroundColor Red
        Cleanup
        exit
    }

    Write-Host "Installing python..."
    $proc = Start-Process -FilePath $pythonInstaller -Verb runAs -ArgumentList '/quiet', 'InstallAllUsers=0', 'PrependPath=1', '/log ".\Python310-Install.log"' -PassThru
    $handle = $proc.Handle
    $proc.WaitForExit();

    if ($proc.ExitCode -ne 0) {
        Write-Warning "The python installer exited with a non-zero status code $($proc.ExitCode)."
    } else {
        # Uncomment if you change 'InstallAllUsers=1' above
        # AddFolderToPathVariable $pythonPath
    }

    RefreshPathVariables

    # Make sure python is now available and within the required version range
    if (CheckCMDExists $pythonCmd) {
        $res = Invoke-Expression "python -V"
        $versionString = $res.Split(' ')[-1]

        if (ComparePythonVersion $versionString) {
            Write-Host "Deleting the downloaded python installer."
            Remove-Item $pythonInstaller
            return $pythonCmd
        }
    }

    Write-Host "Error after installing python. Uninstalling, click 'Yes' if prompted for permission"
    Start-Process -FilePath $pythonInstaller -Verb runAs -ArgumentList '/quiet', '/uninstall' -Wait
    Write-Host "Deleting the downloaded python installer."
    Remove-Item $pythonInstaller
    # Stop OpenAdapt install
    Cleanup
    exit
}


# Check and Install Git and return the git command
function GetGitCMD {
    $gitExists = CheckCMDExists $gitCmd
    if (!$gitExists) {
        # Install git
        Write-Host "Downloading git installer..."
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $gitInstallerURL -OutFile $gitInstaller
        $exists = Test-Path -Path $gitInstaller -PathType Leaf
        if (!$exists) {
            Write-Host "Failed to download git installer" -ForegroundColor Red
            exit
        }

        Write-Host "Installing git..."
        $proc = Start-Process -FilePath $gitInstaller -Verb runAs -ArgumentList '/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /COMPONENTS="icons,ext\reg\shellhere,assoc,assoc_sh"' -PassThru
        $handle = $proc.Handle
        $proc.WaitForExit();

        if ($proc.ExitCode -ne 0) {
            Write-Warning "The git installer exited with a non-zero status code $($proc.ExitCode)."
        } else {
            AddFolderToPathVariable $gitPath
        }
        Write-Host "Deleting the downloaded git installer."
        Remove-Item $gitInstaller

        RefreshPathVariables

        # Make sure git is now available
        $gitExists = CheckCMDExists $gitCmd
        if (!$gitExists) {
            Write-Host "Error after installing git. Uninstalling..."
            Start-Process -FilePath $gitUninstaller -Verb runAs -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait
            Cleanup
            exit
        }
    }
    # Return the git command
    return $gitCmd
}
################################   FUNCTIONS    ################################


################################   SCRIPT    ################################

Write-Host "Install Script Started..." -ForegroundColor Yellow

# Create a new directory and run the setup from there
New-Item -ItemType Directory -Path $setupdir -Force
Set-Location -Path $setupdir
Set-ExecutionPolicy RemoteSigned -Scope Process -Force

# Check and Install the required softwares for OpenAdapt
$tesseract = GetTesseractCMD
RunAndCheck "$tesseract --version" "check TesseractOCR"

$python = GetPythonCMD
RunAndCheck "$python --version" "check Python"

$git = GetGitCMD
RunAndCheck "$git --version" "check Git"

# Setup OpenAdapt in the user's home directory
Set-Location -Path $env:USERPROFILE
if (Test-Path -Path $openAdaptPath -PathType Container) { Remove-Item $openAdaptPath -Force -Recurse }

# OpenAdapt Setup
RunAndCheck "git clone -q $openAdaptURL" "clone git repo $openAdaptURL"
Set-Location $openAdaptPath
RunAndCheck "pip install poetry" "Run ``pip install poetry``"
RunAndCheck "poetry install" "Run ``poetry install``"
RunAndCheck "poetry run alembic upgrade head" "Run ``alembic upgrade head``" -SkipCleanup:$true
RunAndCheck "poetry run pytest" "Run ``Pytest``" -SkipCleanup:$true
Write-Host "OpenAdapt installed Successfully!" -ForegroundColor Green
Start-Process powershell -Verb RunAs -ArgumentList "-NoExit", "-Command", "Set-Location -Path '$openAdaptPath'; poetry shell"

################################   SCRIPT    ################################
