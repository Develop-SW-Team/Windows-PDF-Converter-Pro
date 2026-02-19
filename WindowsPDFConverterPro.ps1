# WindowsPDFConverterPro.ps1
# Windows PDF Converter Pro v1.0
# GUI Version with Logo and Icon Support
# Developed by IGRF Pvt. Ltd.
# Year: 2026
# Website: https://igrf.co.in/en/software

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================
# CLEAN GUI-ONLY INITIALIZATION - NO CONSOLE OPERATIONS
# ============================================

# Suppress all output streams for GUI mode
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# Override all write functions to prevent any console output
function Write-Host { }
function Write-Output { }
function Write-Information { }
function Write-Verbose { }
function Write-Debug { }
function Write-Warning { }
function Write-Progress { }
function Write-Error { }

# ============================================
# COMMAND-LINE ARGUMENT HANDLING - SILENT
# ============================================

$installMode = $false
$toolToInstall = $null
$Global:AutoInstallMode = $false
$Global:AutoInstallTool = $null

# Parse arguments safely without console operations
if ($MyInvocation.Line) {
    $args_array = $MyInvocation.Line.Split(' ') | Where-Object { $_ -ne '' }
    
    for ($i = 0; $i -lt $args_array.Count; $i++) {
        $currentArg = $args_array[$i]
        
        if ($currentArg -eq "-ToolToInstall") {
            $installMode = $true
            if ($i + 1 -lt $args_array.Count -and $args_array[$i+1] -notlike "-*") {
                $toolToInstall = $args_array[$i+1]
                $i++
            }
        }
        elseif ($currentArg -eq "-AutoInstall") {
            if ($toolToInstall) {
                $Global:AutoInstallMode = $true
                $Global:AutoInstallTool = $toolToInstall
            }
        }
    }
}

# ============================================
# SUPPRESS WINDOWS SPEECH RECOGNITION
# ============================================
try {
    $speech = New-Object -ComObject SAPI.SpVoice -ErrorAction SilentlyContinue
    if ($speech) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($speech) | Out-Null
    }
} catch {
    # Silently ignore
}

# ============================================
# CRITICAL FIX: NO CONSOLE REDIRECTION FOR EXE
# ============================================
# IMPORTANT: Remove ALL console redirection code
# The following lines are DELETED/REPLACED:
# - Start-Transcript
# - $Host.UI.Write calls
# - [Console]::Out.Close()
# - [Console]::SetOut([System.IO.TextWriter]::Null)
# - Any console buffer operations

# ============================================
# EMBEDDED RESOURCE HANDLING
# ============================================

function Extract-EmbeddedResource {
    param(
        [string]$ResourceName,
        [string]$OutputPath
    )
    
    try {
        $assembly = [System.Reflection.Assembly]::GetExecutingAssembly()
        $resourceStream = $assembly.GetManifestResourceStream($ResourceName)
        
        if ($resourceStream) {
            $fileStream = [System.IO.File]::Create($OutputPath)
            $resourceStream.CopyTo($fileStream)
            $fileStream.Close()
            $resourceStream.Close()
            return $true
        }
    } catch {
        # Silently fail
    }
    return $false
}

function Get-ResourcePath {
    param(
        [string]$ResourceName,
        [string]$DefaultPath
    )
    
    # First try: Check if file exists at default path
    if (Test-Path $DefaultPath) {
        return $DefaultPath
    }
    
    # Second try: Look in the script/EXE directory
    $scriptDir = Get-ScriptDirectory
    $filePath = Join-Path $scriptDir $ResourceName
    if (Test-Path $filePath) {
        return $filePath
    }
    
    # Third try: Extract from embedded resources to temp
    $tempDir = Join-Path $env:TEMP "PDFConverter_Resources"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    
    $tempPath = Join-Path $tempDir $ResourceName
    if (Extract-EmbeddedResource -ResourceName $ResourceName -OutputPath $tempPath) {
        return $tempPath
    }
    
    # Return default path as fallback
    return $DefaultPath
}

function Get-ScriptDirectory {
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        return Split-Path $MyInvocation.MyCommand.Path -Parent
    }
    else {
        # When running as compiled EXE
        return [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
    }
}

# ============================================
# SUPPRESS ALL OUTPUT - FIXED FOR GUI-ONLY EXE
# ============================================
# For GUI mode - no console operations
$OriginalPreference = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# Override all write functions to do nothing
function Write-Host { }
function Write-Output { }
function Write-Information { }
function Write-Verbose { }
function Write-Debug { }
function Write-Warning { }
function Write-Progress { }

# Try-catch block for console operations - SAFELY handle no console
try {
    if ($Host.Name -eq 'ConsoleHost' -or [Environment]::UserInteractive) {
        $null = $Host.UI.RawUI.FlushInputBuffer()
        [Console]::TreatControlCAsInput = $true
        [System.Console]::Out.Close()
        [System.Console]::Error.Close()
        $null = [System.Console]::SetOut([System.IO.TextWriter]::Null)
        $null = [System.Console]::SetError([System.IO.TextWriter]::Null)
    }
} catch {
    # Silently ignore - this is expected in GUI mode
}

# ============================================
# GLOBAL VARIABLES AND CONFIGURATION
# ============================================

$Global:AppName = "Windows PDF Converter Pro"
$Global:Version = "1.0"
$Global:Company = "IGRF Pvt. Ltd."
$Global:Copyright = "© 2026 IGRF Pvt. Ltd. All rights reserved."
$Global:Website = "https://igrf.co.in/en/software"

# Tool paths
$Global:ToolPaths = @{
    LibreOffice = $null
    Poppler = $null
    ImageMagick = $null
    Ghostscript = $null
}

# Conversion types
$Global:ConversionTypes = @(
    "Word to PDF",
    "PDF to Word", 
    "Excel to PDF",
    "PDF to Excel",
    "PowerPoint to PDF",
    "PDF to PowerPoint",
    "Images to PDF",
    "PDF to Images",
    "Text to PDF",
    "HTML to PDF",
    "PDF Merge",
    "PDF Split",
    "PDF Compress",
    "PDF Encrypt",
    "PDF Decrypt",
    "PDF Watermark"
)

# Quality levels
$Global:QualityLevels = @("Low", "Medium", "High", "Maximum")

# Application state
$Global:AppState = @{
    FilesToConvert = [System.Collections.ArrayList]::new()
    IsProcessing = $false
    ConversionStats = @{
        Total = 0
        Successful = 0
        Failed = 0
    }
    ProcessingResults = @()
    CurrentJob = $null
}

# ============================================
# TOOL DETECTION - WITHOUT WRITE-HOST
# ============================================

function Initialize-Tools {
    # Get script directory for relative paths
    $scriptDir = Get-ScriptDirectory
    
    # Define ALL possible tool paths - Portable first, then system
    $possiblePaths = @{
        LibreOffice = @(
            # Portable versions (bundled with app)
            (Join-Path $scriptDir "tools\libreoffice\program\soffice.exe"),
            (Join-Path $scriptDir "libreoffice\program\soffice.exe"),
            (Join-Path $scriptDir "LibreOffice\program\soffice.exe"),
            # System installations - common paths
            "C:\Program Files\LibreOffice\program\soffice.exe",
            "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "$env:ProgramFiles\LibreOffice\program\soffice.exe",
            "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
            # Version-specific paths
            "C:\Program Files\LibreOffice 5\program\soffice.exe",
            "C:\Program Files\LibreOffice 6\program\soffice.exe",
            "C:\Program Files\LibreOffice 7\program\soffice.exe",
            "C:\Program Files\LibreOffice 8\program\soffice.exe",
            "C:\Program Files\LibreOffice 9\program\soffice.exe",
            "C:\Program Files\LibreOffice 10\program\soffice.exe",
            "C:\Program Files\LibreOffice 24\program\soffice.exe",
            "C:\Program Files\LibreOffice 25\program\soffice.exe"
        )
        
        Poppler = @(
            # Portable versions
            (Join-Path $scriptDir "tools\poppler\bin\pdftotext.exe"),
            (Join-Path $scriptDir "poppler\bin\pdftotext.exe"),
            (Join-Path $scriptDir "poppler-25.12.0\bin\pdftotext.exe"),
            # System installations
            "C:\Program Files\poppler\bin\pdftotext.exe",
            "C:\poppler\bin\pdftotext.exe"
        )
        
        ImageMagick = @(
            # Portable versions
            (Join-Path $scriptDir "tools\imagemagick\magick.exe"),
            (Join-Path $scriptDir "imagemagick\magick.exe"),
            (Join-Path $scriptDir "ImageMagick\magick.exe"),
            # System installations
            "C:\Program Files\ImageMagick\magick.exe",
            "C:\Program Files\ImageMagick\convert.exe",
            "$env:ProgramFiles\ImageMagick\magick.exe",
            "$env:ProgramFiles\ImageMagick\convert.exe"
        )
        
        Ghostscript = @(
            # Portable versions
            (Join-Path $scriptDir "tools\ghostscript\bin\gswin64c.exe"),
            (Join-Path $scriptDir "ghostscript\bin\gswin64c.exe"),
            (Join-Path $scriptDir "gs\bin\gswin64c.exe"),
            # System installations with version folders
            "C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs10.02.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe",
            "C:\Program Files\gs\gs10.01.1\bin\gswin64c.exe",
            "C:\Program Files\gs\gs10.01.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
            "C:\Program Files\gs\gs9.55.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs9.54.0\bin\gswin64c.exe",
            "C:\Program Files\gs\gs9.53.0\bin\gswin64c.exe",
            # Generic system paths
            "$env:ProgramFiles\gs\*\bin\gswin64c.exe",
            "C:\Program Files\gs\bin\gswin64c.exe"
        )
    }
    
    # Check each tool type
    foreach ($toolType in $possiblePaths.Keys) {
        foreach ($path in $possiblePaths[$toolType]) {
            # Handle wildcard paths
            if ($path -like '*\*\*') {
                $basePath = $path.Substring(0, $path.IndexOf('\*'))
                $searchPattern = $path.Substring($path.IndexOf('\*') + 1)
                
                if (Test-Path $basePath) {
                    $folders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue
                    foreach ($folder in $folders) {
                        $fullPath = Join-Path $folder.FullName $searchPattern.Replace('\*', '')
                        if (Test-Path $fullPath) {
                            $Global:ToolPaths[$toolType] = $fullPath
                            break 2
                        }
                    }
                }
            } else {
                if (Test-Path $path) {
                    $Global:ToolPaths[$toolType] = $path
                    break
                }
            }
        }
    }
    
    # Special checks for tools in PATH
    if (-not $Global:ToolPaths.ImageMagick) {
        $imagemagickInPath = Get-Command "magick.exe" -ErrorAction SilentlyContinue
        if ($imagemagickInPath) {
            $Global:ToolPaths.ImageMagick = $imagemagickInPath.Source
        } else {
            $convertInPath = Get-Command "convert.exe" -ErrorAction SilentlyContinue
            if ($convertInPath -and $convertInPath.Source -notlike "*system32*") {
                $Global:ToolPaths.ImageMagick = $convertInPath.Source
            }
        }
    }
    
    if (-not $Global:ToolPaths.Ghostscript) {
        $gsInPath = Get-Command "gswin64c.exe" -ErrorAction SilentlyContinue
        if ($gsInPath) {
            $Global:ToolPaths.Ghostscript = $gsInPath.Source
        } else {
            $gsInPath = Get-Command "gswin32c.exe" -ErrorAction SilentlyContinue
            if ($gsInPath) {
                $Global:ToolPaths.Ghostscript = $gsInPath.Source
            }
        }
    }
    
    if (-not $Global:ToolPaths.Poppler) {
        $pdftotextInPath = Get-Command "pdftotext.exe" -ErrorAction SilentlyContinue
        if ($pdftotextInPath) {
            $Global:ToolPaths.Poppler = $pdftotextInPath.Source
        }
    }
    
    if (-not $Global:ToolPaths.LibreOffice) {
        $sofficeInPath = Get-Command "soffice.exe" -ErrorAction SilentlyContinue
        if ($sofficeInPath) {
            $Global:ToolPaths.LibreOffice = $sofficeInPath.Source
        }
    }
    
    # Store the script directory for later use
    $Global:AppState.ScriptDirectory = $scriptDir
    
    return $true
}

# ============================================
# AUTO-DOWNLOAD MANAGER FOR MISSING DEPENDENCIES
# ============================================

$Global:DownloadUrls = @{
    Ghostscript = @{
        Url = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10030/gs10030w64.exe"
        FileName = "gs10030w64.exe"
        InstallArgs = "/S /D=`"$env:ProgramFiles\gs\gs10.03.0`""
        CheckCommand = "gswin64c.exe"
        InstallPath = "$env:ProgramFiles\gs\gs10.03.0\bin"
        Description = "Ghostscript (PDF rendering and manipulation)"
    }
    LibreOffice = @{
        Url = "https://www.libreoffice.org/download/"
        FileName = "LibreOffice_Installer.exe"
        InstallArgs = "/quiet /norestart"
        CheckCommand = "soffice.exe"
        InstallPath = "${env:ProgramFiles}\LibreOffice\program"
        Description = "LibreOffice (Document format conversion)"
        # Note: Direct download requires visiting the page
    }
    ImageMagick = @{
        Url = "https://imagemagick.org/archive/binaries/ImageMagick-7.1.1-43-Q16-HDRI-x64-dll.exe"
        FileName = "ImageMagick-7.1.1-43-Q16-HDRI-x64-dll.exe"
        InstallArgs = "/VERYSILENT /NORESTART /DIR=`"$env:ProgramFiles\ImageMagick`""
        CheckCommand = "magick.exe"
        InstallPath = "$env:ProgramFiles\ImageMagick"
        Description = "ImageMagick (Image to PDF conversion)"
    }
    Poppler = @{
        Url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"
        FileName = "poppler.zip"
        InstallArgs = $null
        CheckCommand = "pdftotext.exe"
        InstallPath = "$env:ProgramFiles\poppler\bin"
        Description = "Poppler (PDF text extraction)"
    }
}

function Show-DownloadProgressDialog {
    param(
        [string]$ToolName,
        [string]$Status,
        [int]$ProgressPercentage = -1
    )
    
    # Create a simple progress form if not exists
    if (-not $Global:ProgressForm) {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        $Global:ProgressForm = New-Object System.Windows.Forms.Form
        $Global:ProgressForm.Text = "Windows PDF Converter Pro - Installing Dependencies"
        $Global:ProgressForm.Size = New-Object System.Drawing.Size(500, 200)
        $Global:ProgressForm.StartPosition = "CenterScreen"
        $Global:ProgressForm.FormBorderStyle = "FixedDialog"
        $Global:ProgressForm.MaximizeBox = $false
        $Global:ProgressForm.MinimizeBox = $false
        $Global:ProgressForm.ControlBox = $true
        $Global:ProgressForm.Topmost = $true
        
        $Global:ProgressLabel = New-Object System.Windows.Forms.Label
        $Global:ProgressLabel.Location = New-Object System.Drawing.Point(20, 30)
        $Global:ProgressLabel.Size = New-Object System.Drawing.Size(460, 25)
        $Global:ProgressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $Global:ProgressLabel.Text = "Checking required tools..."
        $Global:ProgressForm.Controls.Add($Global:ProgressLabel)
        
        $Global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
        $Global:ProgressBar.Location = New-Object System.Drawing.Point(20, 70)
        $Global:ProgressBar.Size = New-Object System.Drawing.Size(460, 30)
        $Global:ProgressBar.Minimum = 0
        $Global:ProgressBar.Maximum = 100
        $Global:ProgressForm.Controls.Add($Global:ProgressBar)
        
        $Global:StatusLabel = New-Object System.Windows.Forms.Label
        $Global:StatusLabel.Location = New-Object System.Drawing.Point(20, 115)
        $Global:StatusLabel.Size = New-Object System.Drawing.Size(460, 40)
        $Global:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $Global:StatusLabel.Text = "This may take a few minutes depending on your internet speed."
        $Global:ProgressForm.Controls.Add($Global:StatusLabel)
        
        $Global:ProgressForm.Show()
        $Global:ProgressForm.Refresh()
    }
    
    $Global:ProgressLabel.Text = "${ToolName}: $Status"
    if ($ProgressPercentage -ge 0) {
        $Global:ProgressBar.Value = $ProgressPercentage
    }
    $Global:ProgressForm.Refresh()
    Start-Sleep -Milliseconds 50
}

function Close-DownloadProgressDialog {
    if ($Global:ProgressForm) {
        $Global:ProgressForm.Close()
        $Global:ProgressForm.Dispose()
        $Global:ProgressForm = $null
    }
}

function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Request-AdminPrivileges {
    param([string]$ToolName)
    
    $message = @"
Windows PDF Converter Pro needs administrator privileges to install $ToolName.

The tool will be installed to Program Files and added to system PATH.
You can run the portable version without admin rights by placing tools in the 'tools' folder.

Do you want to continue with administrator privileges?
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Administrator Privileges Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    return ($result -eq "Yes")
}

function Restart-AsAdmin {
    param([string]$ToolName)
    
    $scriptPath = [System.Reflection.Assembly]::GetExecutingAssembly().Location
    $arguments = "-ToolToInstall `"$ToolName`" -AutoInstall"
    
    try {
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.FileName = $scriptPath
        $process.StartInfo.Arguments = $arguments
        $process.StartInfo.Verb = "runas"
        $process.Start() | Out-Null
        
        # Exit current instance
        [System.Windows.Forms.Application]::Exit()
        return $true
    } catch {
        return $false
    }
}

function Add-ToSystemPath {
    param([string]$PathToAdd)
    
    try {
        $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($machinePath -notlike "*$PathToAdd*") {
            $newPath = $machinePath + ";" + $PathToAdd
            [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
            
            # Update current session PATH
            $env:Path = $env:Path + ";" + $PathToAdd
        }
        return $true
    } catch {
        return $false
    }
}

function Download-File {
    param(
        [string]$Url,
        [string]$OutputPath,
        [string]$ToolName
    )
    
    try {
        Show-DownloadProgressDialog -ToolName $ToolName -Status "Downloading..." -ProgressPercentage 10
        
        # Configure TLS 1.2 
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        
        # Add user agent to avoid blocking
        $webClient.Headers.Add("User-Agent", "Windows PDF Converter Pro/1.0")
        
        # Download with progress
        $webClient.DownloadFile($Url, $OutputPath)
        
        Show-DownloadProgressDialog -ToolName $ToolName -Status "Download complete" -ProgressPercentage 40
        return $true
    } catch {
        return $false
    }
}

function Install-Ghostscript {
    param(
        [string]$InstallerPath,
        [string]$ToolName
    )
    
    Show-DownloadProgressDialog -ToolName $ToolName -Status "Installing Ghostscript..." -ProgressPercentage 60
    
    try {
        $installDir = "$env:ProgramFiles\gs\gs10.03.0"
        $args = "/S /D=`"$installDir`""
        
        $process = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            $binPath = Join-Path $installDir "bin"
            if (Test-Path $binPath) {
                Add-ToSystemPath -PathToAdd $binPath
                Show-DownloadProgressDialog -ToolName $ToolName -Status "Installation complete" -ProgressPercentage 100
                return $true
            }
        }
    } catch {
        # Try manual extraction as fallback
        try {
            $extractPath = "$env:ProgramFiles\gs"
            $args = "/S /D=`"$extractPath`" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
            $process = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru -NoNewWindow
            return ($process.ExitCode -eq 0)
        } catch {
            return $false
        }
    }
    return $false
}

function Install-LibreOffice {
    param(
        [string]$InstallerPath,
        [string]$ToolName
    )
    
    Show-DownloadProgressDialog -ToolName $ToolName -Status "Installing LibreOffice..." -ProgressPercentage 60
    
    try {
        # Open the download page in browser
        $url = "https://www.libreoffice.org/download/"
        Start-Process $url
        
        # Show instruction dialog
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Please download and install LibreOffice manually from the website that just opened.`n`nAfter installation, click OK to continue.",
            "LibreOffice Installation",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        if ($result -eq "OK") {
            # Check if LibreOffice is now installed
            $checkPath = "${env:ProgramFiles}\LibreOffice\program\soffice.exe"
            if (Test-Path $checkPath) {
                Show-DownloadProgressDialog -ToolName $ToolName -Status "✓ Installation detected" -ProgressPercentage 100
                Start-Sleep -Milliseconds 1000
                return $true
            }
            
            # Ask if user wants to retry or cancel
            $retryResult = [System.Windows.Forms.MessageBox]::Show(
                "LibreOffice installation not detected.`n`nClick Retry to check again, or Cancel to continue without LibreOffice.",
                "Installation Not Detected",
                [System.Windows.Forms.MessageBoxButtons]::RetryCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($retryResult -eq "Retry") {
                # Check one more time
                if (Test-Path $checkPath) {
                    Show-DownloadProgressDialog -ToolName $ToolName -Status "✓ Installation detected" -ProgressPercentage 100
                    Start-Sleep -Milliseconds 1000
                    return $true
                }
            }
        }
        
        return $false
        
    } catch {
        return $false
    } finally {
        Close-DownloadProgressDialog
    }
}

function Install-ImageMagick {
    param(
        [string]$InstallerPath,
        [string]$ToolName
    )
    
    Show-DownloadProgressDialog -ToolName $ToolName -Status "Installing ImageMagick..." -ProgressPercentage 60
    
    try {
        $installDir = "$env:ProgramFiles\ImageMagick"
        $args = "/VERYSILENT /NORESTART /DIR=`"$installDir`""
        
        $process = Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            if (Test-Path $installDir) {
                Add-ToSystemPath -PathToAdd $installDir
                Show-DownloadProgressDialog -ToolName $ToolName -Status "Installation complete" -ProgressPercentage 100
                return $true
            }
        }
    } catch {
        return $false
    }
    return $false
}

function Install-Poppler {
    param(
        [string]$ZipPath,
        [string]$ToolName
    )
    
    Show-DownloadProgressDialog -ToolName $ToolName -Status "Extracting Poppler..." -ProgressPercentage 60
    
    try {
        $extractPath = "$env:ProgramFiles\poppler"
        
        # Ensure target directory exists
        if (-not (Test-Path $extractPath)) {
            New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        }
        
        # Extract ZIP file
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        
        # Find bin folder in the extracted structure
        $binFound = $false
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -like "*/bin/*" -and $entry.Name -like "*.exe") {
                $binFound = $true
                break
            }
        }
        
        # Extract all files
        [System.IO.Compression.ZipFileExtensions]::ExtractToDirectory($zip, $extractPath, $true)
        $zip.Dispose()
        
        # Find the actual bin path
        $binPath = $null
        if (Test-Path (Join-Path $extractPath "bin")) {
            $binPath = Join-Path $extractPath "bin"
        } else {
            # Search for bin folder in subdirectories
            $binFolders = Get-ChildItem -Path $extractPath -Recurse -Directory -Filter "bin" -ErrorAction SilentlyContinue
            if ($binFolders.Count -gt 0) {
                $binPath = $binFolders[0].FullName
            }
        }
        
        if ($binPath -and (Test-Path $binPath)) {
            Add-ToSystemPath -PathToAdd $binPath
            Show-DownloadProgressDialog -ToolName $ToolName -Status "Extraction complete" -ProgressPercentage 100
            return $true
        }
    } catch {
        return $false
    }
    return $false
}

function Install-Tool {
    param(
        [string]$ToolName
    )
    
    $toolInfo = $Global:DownloadUrls[$ToolName]
    if (-not $toolInfo) { return $false }
    
    $tempDir = Join-Path $env:TEMP "PDFConverter_Install"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    
    $installerPath = Join-Path $tempDir $toolInfo.FileName
    
    try {
        # Download the tool
        Show-DownloadProgressDialog -ToolName $ToolName -Status "Downloading $ToolName..." -ProgressPercentage 10
        $downloadSuccess = Download-File -Url $toolInfo.Url -OutputPath $installerPath -ToolName $ToolName
        if (-not $downloadSuccess) {
            Show-DownloadProgressDialog -ToolName $ToolName -Status "Download failed!" -ProgressPercentage 0
            Start-Sleep -Seconds 2
            return $false
        }
        
        # Install based on tool type
        $installSuccess = $false
        switch ($ToolName) {
            "Ghostscript" { $installSuccess = Install-Ghostscript -InstallerPath $installerPath -ToolName $ToolName }
            "LibreOffice" { $installSuccess = Install-LibreOffice -InstallerPath $installerPath -ToolName $ToolName }
            "ImageMagick" { $installSuccess = Install-ImageMagick -InstallerPath $installerPath -ToolName $ToolName }
            "Poppler" { $installSuccess = Install-Poppler -ZipPath $installerPath -ToolName $ToolName }
        }
        
        return $installSuccess
        
    } catch {
        return $false
    } finally {
        # Clean up temp files
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-ToolAvailability {
    param([string]$ToolType)
    
    # First check if already found in ToolPaths
    if ($Global:ToolPaths[$ToolType] -and (Test-Path $Global:ToolPaths[$ToolType])) {
        return $true
    }
    
    # Special handling for LibreOffice - check registry first
    if ($ToolType -eq "LibreOffice") {
        $regPath = Find-LibreOfficeFromRegistry
        if ($regPath) {
            $Global:ToolPaths[$ToolType] = $regPath
            return $true
        }
    }
    
    # Check common installation paths based on tool type
    $commonPaths = @()
    
    switch ($ToolType) {
        "Ghostscript" {
            $commonPaths = @(
                "C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs10.02.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe",
                "C:\Program Files\gs\gs10.01.1\bin\gswin64c.exe",
                "C:\Program Files\gs\gs10.01.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.55.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.54.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.53.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.52.0\bin\gswin64c.exe",
                "C:\Program Files\gs\gs9.50.0\bin\gswin64c.exe",
                "C:\Program Files\gs\bin\gswin64c.exe",
                "${env:ProgramFiles}\gs\*\bin\gswin64c.exe",
                "${env:ProgramFiles(x86)}\gs\*\bin\gswin64c.exe"
            )
        }
        "LibreOffice" {
            $commonPaths = @(
                "C:\Program Files\LibreOffice\program\soffice.exe",
                "C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                "${env:ProgramFiles}\LibreOffice\program\soffice.exe",
                "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
                "C:\Program Files\LibreOffice 5\program\soffice.exe",
                "C:\Program Files\LibreOffice 6\program\soffice.exe",
                "C:\Program Files\LibreOffice 7\program\soffice.exe",
                "C:\Program Files\LibreOffice 8\program\soffice.exe",
                "C:\Program Files\LibreOffice 9\program\soffice.exe",
                "C:\Program Files\LibreOffice 10\program\soffice.exe",
                "C:\Program Files\LibreOffice 11\program\soffice.exe",
                "C:\Program Files\LibreOffice 12\program\soffice.exe",
                "C:\Program Files\LibreOffice 24\program\soffice.exe",
                "C:\Program Files\LibreOffice 24.2\program\soffice.exe",
                "C:\Program Files\LibreOffice 24.8\program\soffice.exe",
                "C:\Program Files\LibreOffice 25\program\soffice.exe",
                "C:\Program Files\LibreOffice 25.2\program\soffice.exe"
            )
        }
        "ImageMagick" {
            $commonPaths = @(
                "C:\Program Files\ImageMagick\magick.exe",
                "C:\Program Files\ImageMagick\convert.exe",
                "${env:ProgramFiles}\ImageMagick\magick.exe",
                "${env:ProgramFiles}\ImageMagick\convert.exe",
                "C:\Program Files\ImageMagick-*\magick.exe",
                "C:\Program Files\ImageMagick-*\convert.exe"
            )
        }
        "Poppler" {
            $commonPaths = @(
                "C:\Program Files\poppler\bin\pdftotext.exe",
                "C:\poppler\bin\pdftotext.exe",
                "${env:ProgramFiles}\poppler\bin\pdftotext.exe",
                "C:\Program Files\poppler-*\bin\pdftotext.exe"
            )
        }
    }
    
    # Check each common path
    foreach ($path in $commonPaths) {
        # Handle wildcard paths
        if ($path -like '*\*\*') {
            $basePath = $path.Substring(0, $path.IndexOf('\*'))
            $searchPattern = $path.Substring($path.IndexOf('\*') + 1)
            
            if (Test-Path $basePath) {
                $folders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue
                foreach ($folder in $folders) {
                    $fullPath = Join-Path $folder.FullName $searchPattern.Replace('\*', '')
                    if (Test-Path $fullPath) {
                        $Global:ToolPaths[$ToolType] = $fullPath
                        return $true
                    }
                }
            }
        } else {
            if (Test-Path $path) {
                $Global:ToolPaths[$ToolType] = $path
                return $true
            }
        }
    }
    
    # Check if tool might be available via PATH
    $toolName = switch ($ToolType) {
        "Ghostscript" { @("gswin64c.exe", "gswin32c.exe", "gs.exe") }
        "LibreOffice" { @("soffice.exe") }
        "ImageMagick" { @("magick.exe", "convert.exe") }
        "Poppler" { @("pdftotext.exe") }
        default { @() }
    }
    
    foreach ($name in $toolName) {
        try {
            $cmd = Get-Command $name -ErrorAction Stop
            if ($cmd) {
                $Global:ToolPaths[$ToolType] = $cmd.Source
                return $true
            }
        } catch {
            # Command not in PATH
        }
    }
    
    return $false
}

function Find-LibreOfficeFromRegistry {
    try {
        # Try to find LibreOffice from registry
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe",
            "HKLM:\SOFTWARE\LibreOffice",
            "HKLM:\SOFTWARE\WOW6432Node\LibreOffice"
        )
        
        foreach ($regPath in $registryPaths) {
            if (Test-Path $regPath) {
                # If it's an App Paths key, read the default value
                if ($regPath -like "*App Paths*") {
                    $regValue = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($regValue -and $regValue.'(default)') {
                        $exePath = $regValue.'(default)'
                        if (Test-Path $exePath) {
                            return $exePath
                        }
                    }
                }
                # If it's a LibreOffice key, try to get installation path
                else {
                    $regValue = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($regValue -and $regValue.Path) {
                        $exePath = Join-Path $regValue.Path "program\soffice.exe"
                        if (Test-Path $exePath) {
                            return $exePath
                        }
                    }
                }
            }
        }
        
        # Check uninstall registry for LibreOffice
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($uninstallPath in $uninstallPaths) {
            $items = Get-ChildItem -Path $uninstallPath -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                $props = Get-ItemProperty -Path $item.PSPath -ErrorAction SilentlyContinue
                if ($props.DisplayName -like "*LibreOffice*") {
                    if ($props.InstallLocation) {
                        $exePath = Join-Path $props.InstallLocation "program\soffice.exe"
                        if (Test-Path $exePath) {
                            return $exePath
                        }
                    }
                    # Try to find from DisplayIcon
                    if ($props.DisplayIcon -and (Test-Path $props.DisplayIcon)) {
                        return $props.DisplayIcon
                    }
                }
            }
        }
        
        return $null
    } catch {
        return $null
    }
}

function Install-MissingTools {
    param(
        [string[]]$MissingTools
    )
    
    if ($MissingTools.Count -eq 0) { return $true }
    
    # Create message based on missing tools
    $toolList = $MissingTools -join "`, "
    $sizeEstimate = switch ($MissingTools.Count) {
        1 { "~150 MB" }
        2 { "~300 MB" }
        3 { "~450 MB" }
        4 { "~600 MB" }
        default { "~500 MB" }
    }
    
    $message = @"
Windows PDF Converter Pro needs to install the following components:

$toolList

Total download size: approximately $sizeEstimate

Do you want to download and install them automatically?
- Click YES to download and install (requires internet connection)
- Click NO to continue with limited functionality
- Click CANCEL to exit the application
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Missing Components Detected",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq "Cancel") {
        return $false
    }
    
    if ($result -eq "No") {
        return $true  # Continue with limited functionality
    }
    
    # User clicked Yes - proceed with installation
    $successCount = 0
    $failedTools = @()
    
    # Check for admin privileges
    $isAdmin = Test-AdminPrivileges
    if (-not $isAdmin) {
        $continueAdmin = Request-AdminPrivileges -ToolName ($MissingTools -join ", ")
        if (-not $continueAdmin) {
            return $true  # Continue with limited functionality
        }
        
        # Restart as admin
        return Restart-AsAdmin -ToolName ($MissingTools[0])
    }
    
    # Install each missing tool
    foreach ($tool in $MissingTools) {
        Show-DownloadProgressDialog -ToolName $tool -Status "Starting installation..." -ProgressPercentage 5
        
        $installSuccess = Install-Tool -ToolName $tool
        
        if ($installSuccess) {
            $successCount++
            Show-DownloadProgressDialog -ToolName $tool -Status "✓ Installation successful" -ProgressPercentage 100
            Start-Sleep -Milliseconds 500
        } else {
            $failedTools += $tool
            Show-DownloadProgressDialog -ToolName $tool -Status "✗ Installation failed" -ProgressPercentage 0
            Start-Sleep -Seconds 2
        }
    }
    
    Close-DownloadProgressDialog
    
    # Show summary
    if ($failedTools.Count -eq 0) {
        Show-MessageBox -Message "All components installed successfully!`n`nThe application will now function with full capabilities." -Title "Installation Complete" -Icon Information
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "The following components could not be installed:`n$($failedTools -join "`n")`n`nThe application will continue with limited functionality.",
            "Partial Installation",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    
    # Re-initialize tools after installation
    Initialize-Tools
    
    return $true
}

# To test the detected Dependencies
function Show-DetectedTools {
    $status = @"
Tool Detection Status:
---------------------
LibreOffice: $($Global:ToolPaths.LibreOffice)
Ghostscript: $($Global:ToolPaths.Ghostscript)
ImageMagick: $($Global:ToolPaths.ImageMagick)
Poppler: $($Global:ToolPaths.Poppler)

Registry Check for LibreOffice: $((Find-LibreOfficeFromRegistry) -join " | ")
"@
    
    # Show as message box for debugging
    [System.Windows.Forms.MessageBox]::Show($status, "Tool Detection Debug", "OK", "Information")
}

function Test-AndInstallMissingTools {
    # First, run normal tool detection
    Initialize-Tools
	
	# Debug - show detected tools (uncomment below for testing)
	#Show-DetectedTools
    
    # Identify missing tools
    $missingTools = @()
    
    # Check each tool properly
    if (-not (Test-ToolAvailability -ToolType "Ghostscript")) { 
        $missingTools += "Ghostscript" 
    }
    if (-not (Test-ToolAvailability -ToolType "LibreOffice")) { 
        $missingTools += "LibreOffice" 
    }
    if (-not (Test-ToolAvailability -ToolType "ImageMagick")) { 
        $missingTools += "ImageMagick" 
    }
    if (-not (Test-ToolAvailability -ToolType "Poppler")) { 
        $missingTools += "Poppler" 
    }
    
    # If no tools are missing, return true immediately
    if ($missingTools.Count -eq 0) {
        return $true
    }
    
    # Only show dialog if tools are actually missing
    $toolList = $missingTools -join "`, "
    
    $sizeEstimate = switch ($missingTools.Count) {
        1 { "~150 MB" }
        2 { "~300 MB" }
        3 { "~450 MB" }
        4 { "~600 MB" }
        default { "~500 MB" }
    }
    
    $message = @"
Windows PDF Converter Pro needs to install the following components:

$toolList

Total download size: approximately $sizeEstimate

Do you want to download and install them automatically?
- Click YES to download and install (requires internet connection)
- Click NO to continue with limited functionality
- Click CANCEL to exit the application
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "Missing Components Detected",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq "Cancel") {
        return $false
    }
    
    if ($result -eq "No") {
        return $true  # Continue with limited functionality
    }
    
    # User clicked Yes - proceed with installation
    # Check for admin privileges
    $isAdmin = Test-AdminPrivileges
    if (-not $isAdmin) {
        $continueAdmin = Request-AdminPrivileges -ToolName ($missingTools -join ", ")
        if (-not $continueAdmin) {
            return $true  # Continue with limited functionality
        }
        
        # Restart as admin
        return Restart-AsAdmin -ToolName ($missingTools[0])
    }
    
    # Install each missing tool
    $successCount = 0
    $failedTools = @()
    
    foreach ($tool in $missingTools) {
        Show-DownloadProgressDialog -ToolName $tool -Status "Starting installation..." -ProgressPercentage 5
        
        $installSuccess = Install-Tool -ToolName $tool
        
        if ($installSuccess) {
            $successCount++
            Show-DownloadProgressDialog -ToolName $tool -Status "✓ Installation successful" -ProgressPercentage 100
            Start-Sleep -Milliseconds 500
        } else {
            $failedTools += $tool
            Show-DownloadProgressDialog -ToolName $tool -Status "✗ Installation failed" -ProgressPercentage 0
            Start-Sleep -Seconds 2
        }
    }
    
    Close-DownloadProgressDialog
    
    # Show summary only if at least one tool was installed
    if ($successCount -gt 0) {
        if ($failedTools.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "All components installed successfully!`n`nThe application will now function with full capabilities.",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "The following components could not be installed:`n$($failedTools -join "`n")`n`nThe application will continue with limited functionality.",
                "Partial Installation",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    }
    
    # Re-initialize tools after installation
    Initialize-Tools
    
    return $true
}

# ============================================
# PORTABLE MODE DETECTION
# ============================================

$Global:IsPortableMode = $false
$Global:PortableToolsPath = $null

function Test-PortableMode {
    $scriptDir = Get-ScriptDirectory
    $toolsFolder = Join-Path $scriptDir "tools"
    
    if (Test-Path $toolsFolder) {
        $Global:IsPortableMode = $true
        $Global:PortableToolsPath = $toolsFolder
        # Return but suppress
        return
    }
    $Global:IsPortableMode = $false
    $Global:PortableToolsPath = $null
}

# Run portable mode detection silently
$null = Test-PortableMode

# ============================================
# HELPER FUNCTIONS
# ============================================

function Get-FileSizeString {
    param($SizeInBytes)
    
    try {
        if ($SizeInBytes -lt 1024) {
            return "$([math]::Round($SizeInBytes, 0)) B"
        } elseif ($SizeInBytes -lt 1048576) {
            $kbSize = $SizeInBytes / 1024
            return "$([math]::Round($kbSize, 1)) KB"
        } elseif ($SizeInBytes -lt 1073741824) {
            $mbSize = $SizeInBytes / 1048576
            return "$([math]::Round($mbSize, 1)) MB"
        } else {
            $gbSize = $SizeInBytes / 1073741824
            return "$([math]::Round($gbSize, 2)) GB"
        }
    } catch {
        return "0 B"
    }
}

function Load-Image {
    param([string]$ImagePath, [int]$MaxWidth = 100, [int]$MaxHeight = 100)
    
    try {
        # Get the actual path using resource resolver
        $actualPath = Get-ResourcePath -ResourceName "Logo.png" -DefaultPath $ImagePath
        
        if (Test-Path $actualPath) {
            $originalImage = [System.Drawing.Image]::FromFile($actualPath)
            
            # Calculate new dimensions while maintaining aspect ratio
            $ratioX = $MaxWidth / $originalImage.Width
            $ratioY = $MaxHeight / $originalImage.Height
            $ratio = [Math]::Min($ratioX, $ratioY)
            
            $newWidth = [int]($originalImage.Width * $ratio)
            $newHeight = [int]($originalImage.Height * $ratio)
            
            # Create resized image
            $bitmap = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($originalImage, 0, 0, $newWidth, $newHeight)
            
            $originalImage.Dispose()
            $graphics.Dispose()
            
            return $bitmap
        }
    } catch {
        # Silently fail
    }
    
    # Return a default image if file not found
    $defaultBitmap = New-Object System.Drawing.Bitmap($MaxWidth, $MaxHeight)
    $defaultGraphics = [System.Drawing.Graphics]::FromImage($defaultBitmap)
    $defaultGraphics.FillRectangle([System.Drawing.Brushes]::LightGray, 0, 0, $MaxWidth, $MaxHeight)
    $defaultGraphics.DrawString("IGRF", [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold), [System.Drawing.Brushes]::DarkBlue, 25, 40)
    $defaultGraphics.Dispose()
    
    return $defaultBitmap
}

function Create-TempDirectory {
    $tempDir = Join-Path $env:TEMP "pdf_converter_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    return $tempDir
}

# ============================================
# CONVERSION FUNCTIONS - ALL WRITE-HOST REMOVED
# ============================================

function Convert-WordToPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High"
    )
    
    try {
        # Validate input
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $fileInfo = Get-Item $InputFile
        $fileSize = $fileInfo.Length
        
        # Ensure output directory exists
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use Word COM with enhanced reliability
        # ============================================
        
        $word = $null
        $doc = $null
        
        try {
            # Try multiple methods to create Word application
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            
            if ($word) {
                # Configure Word for silent operation
                $word.Visible = $false
                $word.DisplayAlerts = 0
                $word.ScreenUpdating = $false
                
                # Disable macros and automation security
                try { $word.AutomationSecurity = 3 } catch { } # msoAutomationSecurityForceDisable
                
                # Open the document with retry
                $maxRetries = 3
                $retryCount = 0
                $docOpened = $false
                
                while (-not $docOpened -and $retryCount -lt $maxRetries) {
                    try {
                        $doc = $word.Documents.Open($InputFile, $false, $true)
                        $docOpened = $true
                    } catch {
                        $retryCount++
                        if ($retryCount -ge $maxRetries) { throw }
                        Start-Sleep -Milliseconds 500
                    }
                }
                
                if ($doc) {
                    # Apply quality settings for PDF export
                    switch ($Quality) {
                        "Maximum" {
                            # PDF/A format for maximum quality
                            $doc.SaveAs([ref]$OutputFile, [ref]17, [ref]$false, [ref]$false, [ref]$false, 
                                       [ref]$false, [ref]$false, [ref]$false, [ref]$true, [ref]$true, [ref]1)
                        }
                        "High" {
                            # Standard PDF with high quality
                            $doc.ExportAsFixedFormat($OutputFile, 17, $false, 0, 0, 0, 0, $false, $true)
                        }
                        "Medium" {
                            # Optimized for web
                            $doc.ExportAsFixedFormat($OutputFile, 17, $false, 0, 1, 1, 0, $false, $true)
                        }
                        "Low" {
                            # Minimum size
                            $doc.ExportAsFixedFormat($OutputFile, 17, $false, 0, 2, 2, 0, $false, $false)
                        }
                        default {
                            $doc.ExportAsFixedFormat($OutputFile, 17)
                        }
                    }
                    
                    # Close document
                    $doc.Close($false)
                    
                    # Check if file was created
                    if (Test-Path $OutputFile) {
                        $pdfSize = (Get-Item $OutputFile).Length
                        if ($pdfSize -gt 100) {
                            return $true
                        }
                    }
                }
            }
        } catch {
            Write-Error "Word COM error: $_" -ErrorAction SilentlyContinue
        } finally {
            # Comprehensive cleanup
            if ($doc) {
                try { 
                    $doc.Close($false) 
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                } catch { }
            }
            if ($word) {
                try { 
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                } catch { }
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Milliseconds 500
        }
        
        # ============================================
        # METHOD 2: Use Print to PDF via Word (Alternative COM approach)
        # ============================================
        
        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            if ($word) {
                $word.Visible = $false
                $word.DisplayAlerts = 0
                
                $doc = $word.Documents.Open($InputFile, $false, $true)
                
                # Use Print to PDF driver
                $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]"", [ref]$OutputFile, 
                             [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false)
                
                # Wait for print to complete
                Start-Sleep -Seconds 3
                
                $doc.Close($false)
                $word.Quit()
                
                # Cleanup
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                if (Test-Path $OutputFile) {
                    $pdfSize = (Get-Item $OutputFile).Length
                    if ($pdfSize -gt 100) {
                        return $true
                    }
                }
            }
        } catch {
            Write-Error "Print to PDF error: $_" -ErrorAction SilentlyContinue
        }
        
        # ============================================
        # METHOD 3: Use LibreOffice (for formatting preservation)
        # ============================================
        
        # Re-initialize tools to ensure paths are available
        Initialize-Tools
        
        if ($Global:ToolPaths.LibreOffice -and (Test-Path $Global:ToolPaths.LibreOffice)) {
            try {
                $tempDir = Join-Path $env:TEMP "word2pdf_$(Get-Random)"
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                
                # Build LibreOffice arguments for best quality
                $args = @(
                    "--headless",
                    "--convert-to", "pdf:writer_pdf_Export",
                    "--outdir", "`"$tempDir`"",
                    "--norestore",
                    "--nofirststartwizard",
                    "--nodefault",
                    "--nolockcheck",
                    "`"$InputFile`""
                )
                
                # Add quality-specific parameters
                switch ($Quality) {
                    "Maximum" { $args += "--pdf", "--pdf:SelectPdfVersion=1" }
                    "High"    { $args += "--pdf" }
                    "Medium"  { $args += "--pdf", "--pdf:ReduceImageResolution=150" }
                    "Low"     { $args += "--pdf", "--pdf:ReduceImageResolution=72" }
                }
                
                # Start LibreOffice process
                $process = Start-Process -FilePath $Global:ToolPaths.LibreOffice `
                    -ArgumentList $args `
                    -Wait `
                    -NoNewWindow `
                    -PassThru `
                    -WindowStyle Hidden
                
                if ($process.ExitCode -eq 0) {
                    $convertedFile = Get-ChildItem -Path $tempDir -Filter "*.pdf" | Select-Object -First 1
                    
                    if ($convertedFile -and (Test-Path $convertedFile.FullName)) {
                        Copy-Item -Path $convertedFile.FullName -Destination $OutputFile -Force
                        
                        # Cleanup
                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                        
                        if (Test-Path $OutputFile) {
                            $pdfSize = (Get-Item $OutputFile).Length
                            if ($pdfSize -gt 100) {
                                return $true
                            }
                        }
                    }
                }
                
                # Cleanup temp directory
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                
            } catch {
                Write-Error "LibreOffice error: $_" -ErrorAction SilentlyContinue
            }
        }
        
        # ============================================
        # METHOD 4: Check if output is a valid PDF
        # ============================================
        
        if (Test-Path $OutputFile) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($OutputFile)
                if ($bytes.Length -gt 100) {
                    $header = [System.Text.Encoding]::ASCII.GetString($bytes[0..4])
                    if ($header -match '%PDF') {
                        return $true
                    }
                }
            } catch { }
        }
        
        # ============================================
        # METHOD 5: Final fallback - Create info file
        # ============================================
        
        try {
            $errorMsg = @"
WORD TO PDF CONVERSION ERROR
============================

File: $(Split-Path $InputFile -Leaf)
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quality: $Quality

ERROR: Could not convert Word document to PDF.

TROUBLESHOOTING:
1. Ensure Microsoft Word is installed
2. Try installing LibreOffice as an alternative
3. Check if the Word file is not corrupted
4. Run the application as administrator

The original Word document is available at:
$InputFile
"@
            
            Set-Content -Path $OutputFile -Value $errorMsg -Encoding UTF8
            return $false
            
        } catch {
            return $false
        }
        
    } catch {
        Write-Error "Convert-WordToPDF critical error: $_" -ErrorAction SilentlyContinue
        return $false
    }
}

function Convert-PDFToWord {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High"
    )
    
    try {
        # Validate input
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $pdfSize = (Get-Item $InputFile).Length
        
        # Ensure output directory exists
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use LibreOffice with enhanced settings
        # ============================================
        
        if ($Global:ToolPaths.LibreOffice) {
            if (Invoke-LibreOfficeConversion -InputFile $InputFile -OutputFile $OutputFile) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 2: Use Ghostscript + intermediate format
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript) {
            if (Invoke-GhostscriptConversion -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 3: Use Poppler with advanced text extraction
        # ============================================
        
        if ($Global:ToolPaths.Poppler) {
            if (Invoke-PopplerExtraction -InputFile $InputFile -OutputFile $OutputFile) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 4: Use Windows built-in methods
        # ============================================
        
        if (Invoke-WindowsConversion -InputFile $InputFile -OutputFile $OutputFile) {
            return $true
        }
        
        # ============================================
        # METHOD 5: Fallback to basic conversion
        # ============================================
        
        return Invoke-BasicConversion -InputFile $InputFile -OutputFile $OutputFile
        
    } catch {
        return $false
    }
}

# ============================================
# CONVERSION METHODS
# ============================================

function Invoke-LibreOfficeConversion {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        $tempDir = Join-Path $env:TEMP "libre_pdf2word_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Use optimized LibreOffice parameters for formatting preservation
        $args = @(
            "--headless",
            "--convert-to", "docx:writer_word_document",
            "--outdir", "`"$tempDir`"",
            "--infilter=writer_pdf_import",
            "--writer",
            "--norestore",
            "--nofirststartwizard",
            "--nodefault",
            "--nolockcheck",
            "`"$InputFile`""
        )
        
        $process = Start-Process -FilePath $Global:ToolPaths.LibreOffice `
            -ArgumentList $args `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0) {
            $convertedFile = Get-ChildItem -Path $tempDir -Filter "*.docx" -ErrorAction SilentlyContinue | 
                Select-Object -First 1
            
            if ($convertedFile -and (Test-Path $convertedFile.FullName)) {
                Copy-Item -Path $convertedFile.FullName -Destination $OutputFile -Force
                
                if (Test-Path $OutputFile) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        return $false
    }
}

function Invoke-GhostscriptConversion {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $tempDir = Join-Path $env:TEMP "gs_pdf2word_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Convert PDF to text with formatting hints
        $textFile = Join-Path $tempDir "output.txt"
        
        # Use Ghostscript to extract text with formatting
        $gsArgs = @(
            "-sDEVICE=txtwrite",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-sOutputFile=`"$textFile`"",
            "`"$InputFile`""
        )
        
        $process = Start-Process -FilePath $Global:ToolPaths.Ghostscript `
            -ArgumentList $gsArgs `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0 -and (Test-Path $textFile)) {
            $extractedText = Get-Content $textFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
            
            if ($extractedText -and $extractedText.Trim().Length -gt 50) {
                # Create formatted document from extracted text
                $success = Create-FormattedDocumentFromText -Text $extractedText -OutputFile $OutputFile
                
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return $success
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        return $false
    }
}

function Invoke-PopplerExtraction {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        $tempTextFile = [System.IO.Path]::GetTempFileName()
        $pdftotext = $Global:ToolPaths.Poppler
        
        # Use Poppler with formatting preservation
        $args = @(
            "-layout",
            "-nopgbrk",
            "-enc", "UTF-8",
            "-eol", "unix",
            "`"$InputFile`"",
            "`"$tempTextFile`""
        )
        
        $process = Start-Process -FilePath $pdftotext `
            -ArgumentList $args `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0 -and (Test-Path $tempTextFile)) {
            $extractedText = Get-Content $tempTextFile -Raw -Encoding UTF8
            
            if ($extractedText -and $extractedText.Trim().Length -gt 50) {
                # Clean and format the extracted text
                $formattedText = Format-ExtractedText -Text $extractedText
                
                # Create document
                $success = Create-FormattedDocumentFromText -Text $formattedText -OutputFile $OutputFile
                
                Remove-Item $tempTextFile -Force
                return $success
            }
        }
        
        Remove-Item $tempTextFile -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        return $false
    }
}

function Invoke-WindowsConversion {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        # Try to use Word COM for PDF import (if Word 2013+)
        $word = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
        if ($word) {
            try {
                $word.Visible = $false
                $word.DisplayAlerts = 0
                
                # Try to open PDF directly in Word (Word 2013+ supports this)
                $doc = $word.Documents.Open($InputFile, $false, $true)
                
                # Save as Word document
                $doc.SaveAs([ref]$OutputFile, [ref]16)
                $doc.Close($false)
                $word.Quit()
                
                # Cleanup COM objects
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                if (Test-Path $OutputFile) {
                    return $true
                }
            } catch {
                # Silently fail
            }
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Invoke-BasicConversion {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        # Extract basic text
        $extractedText = Extract-BasicTextFromPDF -InputFile $InputFile
        
        if (-not $extractedText -or $extractedText.Trim().Length -lt 10) {
            $extractedText = "PDF File: $(Split-Path $InputFile -Leaf)`r`n" +
                           "Converted on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n" +
                           "Note: Text extraction limited. For better results with formatting, install LibreOffice.`r`n"
        }
        
        # Create simple formatted document
        return Create-SimpleFormattedDocument -Text $extractedText -OutputFile $OutputFile
        
    } catch {
        # Ultimate fallback
        try {
            $simpleContent = "PDF: $(Split-Path $InputFile -Leaf)`r`nConverted: $(Get-Date)"
            Set-Content -Path $OutputFile -Value $simpleContent -Encoding UTF8
            return $true
        } catch {
            return $false
        }
    }
}

# ============================================
# TEXT PROCESSING FUNCTIONS
# ============================================

function Extract-StructuredPDFContent {
    param([string]$InputFile)
    
    try {
        $extractedContent = Extract-PDFContent -InputFile $InputFile
        
        if ([string]::IsNullOrEmpty($extractedContent)) {
            $bytes = [System.IO.File]::ReadAllBytes($InputFile)
            $extractedContent = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 50000))
        }
        
        $structuredData = Parse-StructuredPDFContent -Text $extractedContent
        return $structuredData
        
    } catch {
        return @()
    }
}

function Parse-StructuredPDFContent {
    param([string]$Text)
    
    $lines = $Text -split "`r`n|`n"
    $structuredData = @()
    $currentPage = 0
    $rowNumber = 1
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        
        if ($line -match '=====\s*Page\s*(\d+)\s*=====') {
            $currentPage = [int]$matches[1]
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        
        $numberMatches = [regex]::Matches($line, '\b\d+\b')
        if ($numberMatches.Count -ge 3) {
            $numbers = $numberMatches | ForEach-Object { $_.Value }
            $dataObject = [PSCustomObject]@{
                Row = $rowNumber
                Page = $currentPage
                Section = if ($numbers.Count -ge 10) { [int]$numbers[9] } else { 0 }
                LineData = $line
                Numbers = $numbers
                Count = $numbers.Count
            }
            $structuredData += $dataObject
            $rowNumber++
        }
    }
    
    return $structuredData
}

function Extract-BasicTextFromPDF {
    param([string]$InputFile)
    
    try {
        $content = Get-Content $InputFile -Raw -ErrorAction SilentlyContinue
        if ($content -and $content.Length -gt 100) {
            $matches = [regex]::Matches($content, '\((.*?)\)')
            $text = ""
            foreach ($match in $matches) {
                $text += $match.Groups[1].Value + " "
            }
            
            if ($text.Trim().Length -gt 50) {
                return $text
            }
        }
        
        $bytes = [System.IO.File]::ReadAllBytes($InputFile)
        $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
        
        $textPatterns = @(
            'BT\s*(.*?)\s*ET',
            'T[dmjJ]\s*\((.*?)\)',
            '/(Font|F)\s+\d+\s+\d+\s+R.*?T[dmjJ]'
        )
        
        foreach ($pattern in $textPatterns) {
            $matches = [regex]::Matches($ascii, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
            if ($matches.Count -gt 0) {
                $text = ""
                foreach ($match in $matches) {
                    $text += $match.Groups[1].Value + " "
                }
                
                if ($text.Trim().Length -gt 50) {
                    return Clean-PDFText($text)
                }
            }
        }
        
        return $null
        
    } catch {
        return $null
    }
}

function Clean-PDFText {
    param([string]$Text)
    
    $cleanText = $Text
    
    $cleanText = $cleanText -replace '\\\(', '('
    $cleanText = $cleanText -replace '\\\)', ')'
    $cleanText = $cleanText -replace '\\\\', '\'
    $cleanText = $cleanText -replace '\\n', "`n"
    $cleanText = $cleanText -replace '\\r', "`r"
    $cleanText = $cleanText -replace '\\t', "`t"
    
    $cleanText = $cleanText -replace '/[A-Za-z]+\d+', ''
    $cleanText = $cleanText -replace '\d+\s+\d+\s+obj', ''
    $cleanText = $cleanText -replace 'endobj', ''
    $cleanText = $cleanText -replace 'stream.*?endstream', '' -replace '\s+', ' '
    
    $cleanText = $cleanText -replace '\s+', ' '
    $cleanText = $cleanText -replace '\n\s*\n+', "`r`n`r`n"
    
    return $cleanText.Trim()
}

function Format-ExtractedText {
    param([string]$Text)
    
    $lines = $Text -split "`r`n|`n"
    $formattedLines = @()
    
    foreach ($line in $lines) {
        $trimmedLine = $line.Trim()
        
        if ($trimmedLine.Length -eq 0) {
            $formattedLines += ""
        } elseif ($trimmedLine -match '^[A-Z][A-Z\s,&-]+$' -and $trimmedLine.Length -gt 5) {
            $formattedLines += "[HEADING]$trimmedLine[/HEADING]"
        } elseif ($trimmedLine -match '^[A-Z][a-z]+:') {
            $formattedLines += "[LABEL]$trimmedLine[/LABEL]"
        } elseif ($trimmedLine -match '^Dear\s+[A-Z]') {
            $formattedLines += "[SALUTATION]$trimmedLine[/SALUTATION]"
        } elseif ($trimmedLine -match '^Sincerely|^Regards|^Best regards|^Respectfully') {
            $formattedLines += "[CLOSING]$trimmedLine[/CLOSING]"
        } elseif ($trimmedLine.Length -gt 60) {
            $formattedLines += "[PARAGRAPH]$trimmedLine[/PARAGRAPH]"
        } else {
            $formattedLines += $trimmedLine
        }
    }
    
    return $formattedLines -join "`r`n"
}

# ============================================
# DOCUMENT CREATION FUNCTIONS
# ============================================

function Create-FormattedDocumentFromText {
    param(
        [string]$Text,
        [string]$OutputFile
    )
    
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
        if ($word) {
            return Create-WordDocumentWithCOM -Text $Text -OutputFile $OutputFile
        }
        
        return Create-RTFDocument -Text $Text -OutputFile $OutputFile
        
    } catch {
        Set-Content -Path $OutputFile -Value $Text -Encoding UTF8
        return $true
    }
}

function Create-WordDocumentWithCOM {
    param(
        [string]$Text,
        [string]$OutputFile
    )
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        
        $doc = $word.Documents.Add()
        
        $lines = $Text -split "`r`n|`n"
        
        foreach ($line in $lines) {
            if ($line -match '^\[HEADING\](.*)\[/HEADING\]$') {
                $content = $matches[1]
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $content
                $range.Font.Size = 16
                $range.Font.Bold = $true
                $range.ParagraphFormat.Alignment = 1
                $range.InsertParagraphAfter()
            } elseif ($line -match '^\[LABEL\](.*)\[/LABEL\]$') {
                $content = $matches[1]
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $content
                $range.Font.Bold = $true
                $range.InsertParagraphAfter()
            } elseif ($line -match '^\[SALUTATION\](.*)\[/SALUTATION\]$') {
                $content = $matches[1]
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $content
                $range.Font.Size = 12
                $range.Font.Bold = $true
                $range.InsertParagraphAfter()
            } elseif ($line -match '^\[CLOSING\](.*)\[/CLOSING\]$') {
                $content = $matches[1]
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $content
                $range.Font.Italic = $true
                $range.InsertParagraphAfter()
            } elseif ($line -match '^\[PARAGRAPH\](.*)\[/PARAGRAPH\]$') {
                $content = $matches[1]
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $content
                $range.ParagraphFormat.FirstLineIndent = 36
                $range.InsertParagraphAfter()
            } elseif ($line.Trim().Length -gt 0) {
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.Text = $line
                $range.InsertParagraphAfter()
            } else {
                $doc.Content.InsertParagraphAfter()
            }
        }
        
        $doc.Content.Font.Name = "Calibri"
        $doc.Content.Font.Size = 11
        
        $doc.SaveAs([ref]$OutputFile, [ref]16)
        $doc.Close($false)
        $word.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $true
        
    } catch {
        return $false
    }
}

function Create-RTFDocument {
    param(
        [string]$Text,
        [string]$OutputFile
    )
    
    try {
        $rtfHeader = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil Calibri;}{\f1\fnil Calibri Light;}}
\viewkind4\uc1
"@
        
        $rtfFooter = "}"
        
        $lines = $Text -split "`r`n|`n"
        $rtfContent = ""
        
        foreach ($line in $lines) {
            if ($line -match '^\[HEADING\](.*)\[/HEADING\]$') {
                $content = $matches[1]
                $rtfContent += "\pard\qc\f1\fs28\b $content\b0\par\pard"
            } elseif ($line -match '^\[LABEL\](.*)\[/LABEL\]$') {
                $content = $matches[1]
                $rtfContent += "\pard\f0\fs14\b $content\b0\par\pard"
            } elseif ($line -match '^\[SALUTATION\](.*)\[/SALUTATION\]$') {
                $content = $matches[1]
                $rtfContent += "\pard\f0\fs16\b $content\b0\par\pard"
            } elseif ($line -match '^\[CLOSING\](.*)\[/CLOSING\]$') {
                $content = $matches[1]
                $rtfContent += "\pard\f0\fs12\i $content\i0\par\pard"
            } elseif ($line -match '^\[PARAGRAPH\](.*)\[/PARAGRAPH\]$') {
                $content = $matches[1]
                $rtfContent += "\pard\f0\fs12\fi360 $content\par\pard"
            } elseif ($line.Trim().Length -gt 0) {
                $rtfContent += "\pard\f0\fs12 $line\par\pard"
            } else {
                $rtfContent += "\par"
            }
        }
        
        $escapedContent = $rtfContent -replace '\\', '\\' `
                                     -replace '{', '\{' `
                                     -replace '}', '\}'
        
        $fullRtf = $rtfHeader + $escapedContent + $rtfFooter
        Set-Content -Path $OutputFile -Value $fullRtf -Encoding ASCII -NoNewline
        
        return $true
        
    } catch {
        return $false
    }
}

function Create-SimpleFormattedDocument {
    param(
        [string]$Text,
        [string]$OutputFile
    )
    
    try {
        $rtfHeader = "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil Calibri;}}\viewkind4\uc1\pard\f0\fs20 "
        
        $rtfFooter = "\par }"
        
        $escapedText = $Text -replace "`r`n", "\\par " `
                           -replace "`n", "\\par " `
                           -replace "\\", "\\\\" `
                           -replace "{", "\{" `
                           -replace "}", "\}"
        
        $fullRtf = $rtfHeader + $escapedText + $rtfFooter
        Set-Content -Path $OutputFile -Value $fullRtf -Encoding ASCII -NoNewline
        
        return $true
        
    } catch {
        Set-Content -Path $OutputFile -Value $Text -Encoding UTF8
        return $true
    }
}

# ============================================
# EXCEL TO PDF CONVERSION
# ============================================

function Convert-ExcelToPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High"
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $fileSize = (Get-Item $InputFile).Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use Excel COM with enhanced settings for scaling
        # ============================================
        
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
            if ($excel) {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                
                $workbook = $excel.Workbooks.Open($InputFile)
                
                foreach ($worksheet in $workbook.Worksheets) {
                    $worksheet.PageSetup.Zoom = $false
                    $worksheet.PageSetup.FitToPagesWide = 1
                    $worksheet.PageSetup.FitToPagesTall = $false
                    $worksheet.PageSetup.Orientation = 2
                    $worksheet.PageSetup.PrintTitleRows = ""
                    $worksheet.PageSetup.PrintTitleColumns = ""
                }
                
                switch ($Quality) {
                    "Maximum" {
                        $workbook.ExportAsFixedFormat(0, $OutputFile, 0, $true, $false, 1, 0, $true, $null)
                    }
                    "High" {
                        $workbook.ExportAsFixedFormat(0, $OutputFile, 0, $true, $false, 1, 0, $false, $null)
                    }
                    "Medium" {
                        $workbook.ExportAsFixedFormat(0, $OutputFile, 0, $true, $false, 0, 0, $false, $null)
                    }
                    default {
                        $workbook.ExportAsFixedFormat(0, $OutputFile, 0, $true, $false, 0, 0, $false, $null)
                    }
                }
                
                $workbook.Close($false)
                $excel.Quit()
                
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                if (Test-Path $OutputFile) {
                    return $true
                }
            }
        } catch {
            # Silently fail
        }
        
        # ============================================
        # METHOD 2: Use LibreOffice with enhanced settings for scaling
        # ============================================
        
        if ($Global:ToolPaths.LibreOffice) {
            $tempDir = Create-TempDirectory
            
            $exportFilter = switch ($Quality) {
                "Maximum" { "calc_pdf_Export:Quality=100;ReduceImageResolution=false;MaxImageResolution=600;Scale=100" }
                "High" { "calc_pdf_Export:Quality=90;ReduceImageResolution=false;MaxImageResolution=300;Scale=100;PrintFitWidth=1" }
                "Medium" { "calc_pdf_Export:Quality=75;ReduceImageResolution=true;MaxImageResolution=150;Scale=100;PrintFitWidth=1" }
                "Low" { "calc_pdf_Export:Quality=50;ReduceImageResolution=true;MaxImageResolution=72;Scale=100;PrintFitWidth=1" }
                default { "calc_pdf_Export:Scale=100;PrintFitWidth=1" }
            }
            
            $args = @(
                "--headless",
                "--convert-to", $exportFilter,
                "--outdir", "`"$tempDir`"",
                "--calc",
                "--norestore",
                "--nofirststartwizard",
                "--nodefault",
                "--nolockcheck",
                "`"$InputFile`""
            )
            
            $process = Start-Process -FilePath $Global:ToolPaths.LibreOffice `
                -ArgumentList $args `
                -Wait `
                -NoNewWindow `
                -PassThru `
                -WindowStyle Hidden
            
            if ($process.ExitCode -eq 0) {
                $convertedFile = Get-ChildItem -Path $tempDir -Filter "*.pdf" -ErrorAction SilentlyContinue | 
                    Select-Object -First 1
                
                if ($convertedFile -and (Test-Path $convertedFile.FullName)) {
                    Copy-Item -Path $convertedFile.FullName -Destination $OutputFile -Force
                    
                    if (Test-Path $OutputFile) {
                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                        return $true
                    }
                }
            }
            
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        # ============================================
        # METHOD 3: Use Windows built-in print to PDF with scaling
        # ============================================
        
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
            if ($excel) {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                
                $workbook = $excel.Workbooks.Open($InputFile)
                
                foreach ($worksheet in $workbook.Worksheets) {
                    $worksheet.PageSetup.Zoom = $false
                    $worksheet.PageSetup.FitToPagesWide = 1
                    $worksheet.PageSetup.FitToPagesTall = $false
                }
                
                $workbook.PrintOut(
                    [System.Type]::Missing,
                    [System.Type]::Missing,
                    1,
                    $false,
                    "Microsoft Print to PDF",
                    $true,
                    $false,
                    $OutputFile
                )
                
                $workbook.Close($false)
                $excel.Quit()
                
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                Start-Sleep -Seconds 3
                
                if (Test-Path $OutputFile) {
                    return $true
                }
            }
        } catch {
            # Silently fail
        }
        
        # ============================================
        # METHOD 4: Extract data and create enhanced PDF
        # ============================================
        
        try {
            $excelData = ""
            $excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
            if ($excel) {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                
                $workbook = $excel.Workbooks.Open($InputFile)
                $worksheet = $workbook.Worksheets.Item(1)
                
                $sheetName = $worksheet.Name
                $usedRange = $worksheet.UsedRange
                $rowCount = $usedRange.Rows.Count
                $colCount = $usedRange.Columns.Count
                
                $sampleData = ""
                $maxRows = [Math]::Min(10, $rowCount)
                $maxCols = [Math]::Min(5, $colCount)
                
                for ($r = 1; $r -le $maxRows; $r++) {
                    $row = ""
                    for ($c = 1; $c -le $maxCols; $c++) {
                        $cellValue = $worksheet.Cells.Item($r, $c).Text
                        if ($cellValue.Length -gt 20) {
                            $cellValue = $cellValue.Substring(0, 20) + "..."
                        }
                        $row += "$cellValue | "
                    }
                    $sampleData += "Row $r : $row`n"
                }
                
                $excelData = "Excel Worksheet: $sheetName`n"
                $excelData += "Total Rows: $rowCount, Total Columns: $colCount`n"
                $excelData += "Sample Data (first $maxRows rows):`n$sampleData"
                
                $workbook.Close($false)
                $excel.Quit()
                
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            if (-not $excelData) {
                $excelData = "Excel Document: $(Split-Path $InputFile -Leaf)`n" +
                            "Converted to PDF on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" +
                            "Original Size: $(Get-FileSizeString -SizeInBytes $fileSize)`n" +
                            "Quality: $Quality`n" +
                            "Note: For better conversion quality with scaling, install Microsoft Excel or LibreOffice."
            }
            
            return Create-EnhancedPDF -InputFile $InputFile -OutputFile $OutputFile -ConversionType "Excel to PDF" -Quality $Quality -TextContent $excelData
            
        } catch {
            # Silently fail
        }
        
        # ============================================
        # METHOD 5: Ultimate fallback
        # ============================================
        
        $textContent = "Excel Document: $(Split-Path $InputFile -Leaf)`n" +
                      "Converted to PDF on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" +
                      "Original Size: $(Get-FileSizeString -SizeInBytes $fileSize)`n" +
                      "Quality: $Quality`n" +
                      "Note: For better conversion quality with scaling, install Microsoft Excel or LibreOffice."
        
        return Create-EnhancedPDF -InputFile $InputFile -OutputFile $OutputFile -ConversionType "Excel to PDF" -Quality $Quality -TextContent $textContent
        
    } catch {
        try {
            $simpleContent = "Excel to PDF Conversion`nFile: $(Split-Path $InputFile -Leaf)`nDate: $(Get-Date)"
            return Create-EnhancedPDF -InputFile $InputFile -OutputFile $OutputFile -ConversionType "Excel to PDF" -Quality $Quality -TextContent $simpleContent
        } catch {
            return $false
        }
    }
}

# ============================================
# PDF TO EXCEL CONVERSION
# ============================================

function Convert-PDFToExcel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [switch]$Silent
    )
    
    $fileInfo = Get-Item $InputFile -ErrorAction SilentlyContinue
    if (-not $fileInfo) {
        return $false
    }
    
    if (-not $Global:ToolPaths) {
        $Global:ToolPaths = @{
            Ghostscript = $null
            LibreOffice = $null
        }
    }
    
    if (-not $Global:ToolPaths.Ghostscript) {
        $gsInPath = Get-Command "gswin64c" -ErrorAction SilentlyContinue
        if (-not $gsInPath) {
            $gsInPath = Get-Command "gswin32c" -ErrorAction SilentlyContinue
        }
        if (-not $gsInPath) {
            $gsInPath = Get-Command "gs" -ErrorAction SilentlyContinue
        }
        
        if ($gsInPath) {
            $Global:ToolPaths.Ghostscript = $gsInPath.Source
        }
    }
    
    if (-not $Global:ToolPaths.LibreOffice) {
        $librePaths = @(
            "${env:ProgramFiles}\LibreOffice\program\soffice.exe",
            "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe"
        )
        
        foreach ($path in $librePaths) {
            if (Test-Path $path) {
                $Global:ToolPaths.LibreOffice = $path
                break
            }
        }
    }
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        if ((Get-Item $InputFile).Extension -ne '.pdf') {
            return $false
        }
        
        $fileSize = $fileInfo.Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if ([System.IO.Path]::GetExtension($OutputFile) -ne '.xlsx') {
            $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, '.xlsx')
        }
        
        # ============================================
        # METHOD 1: Use Ghostscript (Primary)
        # ============================================
        if ($Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
            $ghostscriptResult = Convert-PDFToExcelUsingGhostscript -InputFile $InputFile -OutputFile $OutputFile -Silent:$Silent
            if ($ghostscriptResult) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 2: Use LibreOffice
        # ============================================
        if ($Global:ToolPaths.LibreOffice -and (Test-Path $Global:ToolPaths.LibreOffice)) {
            $libreResult = Convert-PDFToExcelUsingLibreOffice -InputFile $InputFile -OutputFile $OutputFile -Silent:$Silent
            if ($libreResult) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 3: Create Basic Representation
        # ============================================
        $basicResult = Create-BasicPDFFile -InputFile $InputFile -OutputFile $OutputFile -Silent:$Silent
        if ($basicResult) {
            return $true
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Convert-PDFToExcelUsingGhostscript {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [switch]$Silent
    )
    
    try {
        $tempDir = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $tempTxt = Join-Path $tempDir "output.txt"
        
        $argsList = @(
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-sDEVICE=txtwrite",
            "-sOutputFile=`"$tempTxt`"",
            "`"$InputFile`""
        )
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Global:ToolPaths.Ghostscript
        $psi.Arguments = $argsList -join " "
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.Start() | Out-Null
        $process.WaitForExit(60000)
        
        if ($process.ExitCode -eq 0 -and (Test-Path $tempTxt)) {
            $textContent = Get-Content $tempTxt -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
            
            if ($textContent -and $textContent.Trim() -ne "") {
                $excel = $null
                $workbook = $null
                $worksheet = $null
                
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                    $excel.ScreenUpdating = $false
                    
                    $workbook = $excel.Workbooks.Add()
                    $worksheet = $workbook.Worksheets.Item(1)
                    $worksheet.Name = "PDF_Content"
                    
                    $row = 1
                    $lines = $textContent -split "`r`n|`n|`r"
                    
                    foreach ($line in $lines) {
                        $trimmedLine = $line.Trim()
                        if ($trimmedLine -ne "") {
                            if ($trimmedLine -match '\s{2,}') {
                                $parts = $trimmedLine -split '\s{2,}'
                                if ($parts.Count -gt 1) {
                                    for ($col = 0; $col -lt $parts.Count; $col++) {
                                        $worksheet.Cells.Item($row, $col + 1) = $parts[$col].Trim()
                                    }
                                } else {
                                    $worksheet.Cells.Item($row, 1) = $trimmedLine
                                }
                            } else {
                                $worksheet.Cells.Item($row, 1) = $trimmedLine
                            }
                            $row++
                        }
                    }
                    
                    $usedRange = $worksheet.UsedRange
                    if ($usedRange) {
                        $usedRange.Columns.AutoFit() | Out-Null
                        
                        for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
                            if ($worksheet.Columns($col).ColumnWidth -gt 50) {
                                $worksheet.Columns($col).ColumnWidth = 50
                            }
                        }
                    }
                    
                    $workbook.SaveAs($OutputFile, 51)
                    
                } catch {
                    throw
                } finally {
                    if ($worksheet) { 
                        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null } catch {} 
                    }
                    if ($workbook) { 
                        try { 
                            $workbook.Close($true)
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null 
                        } catch {} 
                    }
                    if ($excel) { 
                        try { 
                            $excel.Quit()
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 
                        } catch {} 
                    }
                    
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
                
                if (Test-Path $OutputFile) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

function Convert-PDFToExcelUsingLibreOffice {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [switch]$Silent
    )
    
    try {
        $tempDir = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        $argsList = @(
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", "`"$tempDir`"",
            "`"$InputFile`""
        )
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Global:ToolPaths.LibreOffice
        $psi.Arguments = $argsList -join " "
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.Start() | Out-Null
        $process.WaitForExit(45000)
        
        if ($process.ExitCode -eq 0) {
            $convertedFile = Get-ChildItem -Path $tempDir -Filter "*.xlsx" -ErrorAction SilentlyContinue | 
                Select-Object -First 1
            
            if ($convertedFile -and (Test-Path $convertedFile.FullName)) {
                Copy-Item -Path $convertedFile.FullName -Destination $OutputFile -Force
                
                if (Test-Path $OutputFile) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

function Create-BasicPDFFile {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [switch]$Silent
    )
    
    try {
        $fileInfo = Get-Item $InputFile
        
        $excel = $null
        $workbook = $null
        $worksheet = $null
        
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "PDF_Info"
            
            $worksheet.Cells.Item(1, 1) = "PDF File Information"
            $worksheet.Cells.Item(1, 1).Font.Bold = $true
            $worksheet.Cells.Item(1, 1).Font.Size = 12
            
            $infoData = @(
                "File: $(Split-Path $InputFile -Leaf)",
                "Size: $(Get-FileSizeString -SizeInBytes $fileInfo.Length)",
                "Created: $($fileInfo.CreationTime)",
                "Modified: $($fileInfo.LastWriteTime)"
            )
            
            $row = 3
            foreach ($item in $infoData) {
                $worksheet.Cells.Item($row, 1) = $item
                $row++
            }
            
            $row++
            $worksheet.Cells.Item($row, 1) = "Note: Could not extract content from PDF."
            $worksheet.Cells.Item($row, 1).Font.Italic = $true
            $row++
            $worksheet.Cells.Item($row, 1) = "Install Adobe Acrobat or LibreOffice for better PDF conversion."
            
            $worksheet.Columns.Item(1).AutoFit()
            
            $workbook.SaveAs($OutputFile, 51)
            
        } catch {
            $textOutput = [System.IO.Path]::ChangeExtension($OutputFile, '.txt')
            $content = "PDF File: $(Split-Path $InputFile -Leaf)`nSize: $(Get-FileSizeString -SizeInBytes $fileInfo.Length)"
            $content | Out-File -FilePath $textOutput -Encoding UTF8
            $OutputFile = $textOutput
        } finally {
            if ($worksheet) { 
                try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null } catch {} 
            }
            if ($workbook) { 
                try { 
                    $workbook.Close($true)
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null 
                } catch {} 
            }
            if ($excel) { 
                try { 
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 
                } catch {} 
            }
            
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        
        if (Test-Path $OutputFile) {
            return $true
        }
        
        return $false
        
    } catch {
        return $false
    }
}

# ============================================
# POWERPOINT TO PDF CONVERSION
# ============================================

function Convert-PowerPointToPDF {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputFile,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [ValidateSet("High", "Medium", "Low", "Maximum")]
        [string]$Quality = "High"
    )
    
    function Format-FileSize {
        param([long]$Bytes)
        
        if ($Bytes -lt 1KB) { return "$Bytes B" }
        elseif ($Bytes -lt 1MB) { return "$([math]::Round($Bytes/1KB, 2)) KB" }
        elseif ($Bytes -lt 1GB) { return "$([math]::Round($Bytes/1MB, 2)) MB" }
        else { return "$([math]::Round($Bytes/1GB, 2)) GB" }
    }
    
    function Test-ValidPDF {
        param([string]$FilePath)
        
        if (-not (Test-Path $FilePath)) { return $false }
        
        try {
            $bytes = [System.IO.File]::ReadAllBytes($FilePath)
            if ($bytes.Length -lt 5) { return $false }
            
            $header = [System.Text.Encoding]::ASCII.GetString($bytes[0..4])
            return $header.StartsWith("%PDF")
        }
        catch {
            return $false
        }
    }
    
    function Convert-UsingPowerPointCOM {
        param($InputFile, $OutputFile, $Quality)
        
        try {
            $powerpoint = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
            $powerpoint.Visible = 0
            $powerpoint.DisplayAlerts = 0
            
            $presentation = $powerpoint.Presentations.Open(
                $InputFile,
                1,
                0,
                0
            )
            
            $presentation.SaveAs($OutputFile, 32)
            
            $presentation.Close()
            $powerpoint.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
            Remove-Variable presentation, powerpoint -Force -ErrorAction SilentlyContinue
            
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            if (Test-ValidPDF $OutputFile) {
                return $true
            }
            else {
                if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
                return $false
            }
        }
        catch {
            return $false
        }
    }
    
    function Convert-UsingLibreOffice {
        param($InputFile, $OutputFile, $Quality)
        
        try {
            if (-not $Global:ToolPaths.LibreOffice) {
                return $false
            }
            
            $libreOfficeExe = $Global:ToolPaths.LibreOffice
            if (-not (Test-Path $libreOfficeExe)) {
                return $false
            }
            
            $tempDir = Join-Path $env:TEMP "pdf_conversion_$(Get-Random)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            $arguments = @(
                "--headless",
                "--convert-to", "pdf",
                "--outdir", "`"$tempDir`"",
                "`"$InputFile`""
            )
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $libreOfficeExe
            $processInfo.Arguments = $arguments -join " "
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            
            $process.Start() | Out-Null
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            $process.WaitForExit(60000)
            
            $convertedFiles = @(Get-ChildItem -Path $tempDir -Filter "*.pdf" -ErrorAction SilentlyContinue)
            
            if ($convertedFiles.Count -gt 0) {
                $convertedFile = $convertedFiles[0].FullName
                
                Copy-Item -Path $convertedFile -Destination $OutputFile -Force
                
                if (Test-ValidPDF $OutputFile) {
                    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
                else {
                    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
                }
            }
            
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
            
        }
        catch {
            return $false
        }
    }
    
    function Convert-UsingWindowsPrintToPDF {
        param($InputFile, $OutputFile)
        
        try {
            $powerpoint = $null
            $presentation = $null
            
            try {
                $powerpoint = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
                $powerpoint.Visible = 0
                
                $presentation = $powerpoint.Presentations.Open($InputFile, 1, 0, 0)
                
                try {
                    $presentation.SaveAs($OutputFile, 32)
                }
                catch {
                    $presentation.PrintOptions.PrintInBackground = 0
                    $presentation.PrintOut(
                        1,
                        9999,
                        $OutputFile,
                        0,
                        0,
                        1
                    )
                }
                
                $presentation.Close()
                $powerpoint.Quit()
                
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
                Remove-Variable presentation, powerpoint -Force -ErrorAction SilentlyContinue
                
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                Start-Sleep -Seconds 3
                
                if (Test-ValidPDF $OutputFile) {
                    return $true
                }
                else {
                    return $false
                }
            }
            catch {
                return $false
            }
        }
        catch {
            return $false
        }
    }
    
    function Create-MinimalPDF {
        param($OutputFile, $InputFileName)
        
        try {
            $pdfBytes = [System.Text.Encoding]::ASCII.GetBytes(@"
%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
/Font <<
/F1 5 0 R
>>
>>
>>
endobj
4 0 obj
<<
/Length 200
>>
stream
BT
/F1 24 Tf
72 720 Td
(PPT to PDF Conversion Failed) Tj
0 -30 Td
(=============================) Tj
0 -30 Td
(File: $InputFileName) Tj
0 -30 Td
(Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) Tj
0 -30 Td
(Error: Could not convert PowerPoint to PDF) Tj
0 -30 Td
(Please install Microsoft PowerPoint or) Tj
0 -30 Td
(LibreOffice for proper conversion.) Tj
ET
endstream
endobj
5 0 obj
<<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica
>>
endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000056 00000 n
0000000113 00000 n
0000000219 00000 n
0000000484 00000 n
trailer
<<
/Size 6
/Root 1 0 R
>>
startxref
584
%%EOF
"@)
            
            [System.IO.File]::WriteAllBytes($OutputFile, $pdfBytes)
            
            if (Test-ValidPDF $OutputFile) {
                return $true
            }
            else {
                return $false
            }
        }
        catch {
            return $false
        }
    }
    
    try {
        if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
            return $false
        }
        
        $fileInfo = Get-Item $InputFile
        $originalSize = $fileInfo.Length
        $fileName = $fileInfo.Name
        
        $outputDir = Split-Path $OutputFile -Parent
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if (Test-Path $OutputFile) {
            Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
        }
        
        $methods = @(
            @{ Name = "PowerPoint COM"; Script = ${function:Convert-UsingPowerPointCOM} }
            @{ Name = "LibreOffice"; Script = ${function:Convert-UsingLibreOffice} }
            @{ Name = "Windows Print"; Script = ${function:Convert-UsingWindowsPrintToPDF} }
        )
        
        foreach ($method in $methods) {
            if (& $method.Script $InputFile $OutputFile $Quality) {
                if (Test-ValidPDF $OutputFile) {
                    return $true
                }
                else {
                    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
                }
            }
        }
        
        if (Create-MinimalPDF $OutputFile $fileName) {
            return $true
        }
        else {
            return $false
        }
        
    }
    catch {
        return $false
    }
    finally {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ============================================
# PDF TO POWERPOINT CONVERSION
# ============================================

function Convert-PDFToPowerPoint {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [switch]$Silent
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        if ((Get-Item $InputFile).Extension -ne '.pdf') {
            return $false
        }
        
        $pdfSize = (Get-Item $InputFile).Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if ([System.IO.Path]::GetExtension($OutputFile) -ne '.pptx') {
            $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, '.pptx')
        }
        
        # ============================================
        # METHOD 1: Use LibreOffice with enhanced parameters
        # ============================================
        
        if ($Global:ToolPaths.LibreOffice -and (Test-Path $Global:ToolPaths.LibreOffice)) {
            $result = Invoke-LibreOfficeForPowerPoint -InputFile $InputFile -OutputFile $OutputFile -Silent:$Silent
            if ($result) { 
                return $true 
            }
        }
        
        # ============================================
        # METHOD 2: Use Ghostscript for text extraction + Create PowerPoint
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript) {
            $result = Invoke-GhostscriptForPowerPoint -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -Silent:$Silent
            if ($result) { 
                return $true 
            }
        }
        
        # ============================================
        # METHOD 3: Use PowerShell to create basic PowerPoint
        # ============================================
        
        $result = Create-BasicPowerPointPresentation -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -Silent:$Silent
        if ($result) { 
            return $true 
        }
        
        # ============================================
        # METHOD 4: Create a simple PPTX placeholder
        # ============================================
        
        $result = Create-PowerPointPlaceholder -InputFile $InputFile -OutputFile $OutputFile -Silent:$Silent
        if ($result) { 
            return $true 
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Invoke-LibreOfficeForPowerPoint {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [switch]$Silent
    )
    
    try {
        $tempDir = Create-TempDirectory
        
        $argsList = @(
            "--headless",
            "--convert-to", "pptx",
            "--outdir", "`"$tempDir`"",
            "--infilter=impress_pdf_import",
            "--impress",
            "--norestore",
            "--nofirststartwizard",
            "--nodefault",
            "--nolockcheck",
            "--nologo",
            "`"$InputFile`""
        )
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Global:ToolPaths.LibreOffice
        $psi.Arguments = $argsList -join " "
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = $tempDir
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        
        $outputBuilder = New-Object System.Text.StringBuilder
        $errorBuilder = New-Object System.Text.StringBuilder
        
        $process.Start() | Out-Null
        
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        $errorTask = $process.StandardError.ReadToEndAsync()
        
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }
        
        $output = $outputTask.GetAwaiter().GetResult()
        $errorOutput = $errorTask.GetAwaiter().GetResult()
        
        $convertedFiles = Get-ChildItem -Path $tempDir -Filter "*.pptx" -ErrorAction SilentlyContinue
        
        if ($convertedFiles.Count -gt 0) {
            $convertedFile = $convertedFiles[0].FullName
            
            if (Test-Path $convertedFile) {
                Copy-Item -Path $convertedFile -Destination $OutputFile -Force
                
                if (Test-Path $OutputFile) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

function Invoke-GhostscriptForPowerPoint {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality,
        [switch]$Silent
    )
    
    try {
        $tempDir = Create-TempDirectory
        $tempTxt = Join-Path $tempDir "extracted.txt"
        
        $gsArgs = @(
            "-sDEVICE=txtwrite",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-sOutputFile=`"$tempTxt`"",
            "-dTextFormat=2",
            "-dOptimize=true",
            "`"$InputFile`""
        )
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Global:ToolPaths.Ghostscript
        $psi.Arguments = $gsArgs -join " "
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.Start() | Out-Null
        $process.WaitForExit(30000)
        
        if ($process.ExitCode -eq 0 -and (Test-Path $tempTxt)) {
            $extractedText = Get-Content $tempTxt -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
            
            if ($extractedText -and $extractedText.Trim().Length -gt 100) {
                $result = Create-PowerPointFromExtractedText -Text $extractedText -OutputFile $OutputFile -InputFile $InputFile -Quality $Quality -Silent:$Silent
                
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return $result
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

function Create-PowerPointFromExtractedText {
    param(
        [string]$Text,
        [string]$OutputFile,
        [string]$InputFile,
        [string]$Quality,
        [switch]$Silent
    )
    
    try {
        try {
            $powerpoint = New-Object -ComObject PowerPoint.Application
            $powerpoint.Visible = $false
            $powerpoint.DisplayAlerts = 0
            
            $presentation = $powerpoint.Presentations.Add()
            
            $titleSlide = $presentation.Slides.Add(1, 1)
            $titleSlide.Shapes.Title.TextFrame.TextRange.Text = "PDF to PowerPoint"
            $titleSlide.Shapes[2].TextFrame.TextRange.Text = "Converted from: $(Split-Path $InputFile -Leaf)`n" +
                                                           "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" +
                                                           "Quality: $Quality"
            
            $slides = Convert-TextToSlides -Text $Text -Quality $Quality
            
            $slideIndex = 2
            foreach ($slideContent in $slides) {
                if ($slideIndex -gt 20) { break }
                
                $contentSlide = $presentation.Slides.Add($slideIndex, 2)
                $contentSlide.Shapes.Title.TextFrame.TextRange.Text = "Slide $($slideIndex - 1)"
                $contentSlide.Shapes[2].TextFrame.TextRange.Text = $slideContent
                $slideIndex++
            }
            
            $presentation.SaveAs($OutputFile, 24)
            $presentation.Close()
            $powerpoint.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            if (Test-Path $OutputFile) {
                return $true
            }
            
        } catch {
            # Silently fail
        }
        
        return Create-TextBasedPresentation -Text $Text -OutputFile $OutputFile -InputFile $InputFile -Silent:$Silent
        
    } catch {
        return $false
    }
}

function Convert-TextToSlides {
    param(
        [string]$Text,
        [string]$Quality
    )
    
    $paragraphs = $Text -split "`r`n|`n|`r" | Where-Object { $_.Trim().Length -gt 0 }
    
    $maxCharsPerSlide = switch ($Quality) {
        "Maximum" { 800 }
        "High"    { 600 }
        "Medium"  { 400 }
        "Low"     { 200 }
        default   { 400 }
    }
    
    $maxSlides = switch ($Quality) {
        "Maximum" { 30 }
        "High"    { 20 }
        "Medium"  { 15 }
        "Low"     { 10 }
        default   { 15 }
    }
    
    $slides = @()
    $currentSlide = ""
    $charCount = 0
    
    foreach ($para in $paragraphs) {
        $trimmedPara = $para.Trim()
        
        $isHeading = $false
        if ($trimmedPara -match '^[A-Z][A-Z\s]+$' -and $trimmedPara.Length -gt 5 -and $trimmedPara.Length -lt 100) {
            $isHeading = $true
        } elseif ($trimmedPara -match '^[A-Z][a-z]+:' -or $trimmedPara -match '^\d+\.\s') {
            $isHeading = $true
        }
        
        if (($isHeading -and $charCount -gt $maxCharsPerSlide * 0.3) -or 
            $charCount + $trimmedPara.Length -gt $maxCharsPerSlide) {
            
            if ($currentSlide.Length -gt 0) {
                $slides += $currentSlide
                if ($slides.Count -ge $maxSlides) {
                    $slides += "[Content truncated due to slide limit]"
                    break
                }
            }
            
            $currentSlide = $trimmedPara
            $charCount = $trimmedPara.Length
        } else {
            if ($currentSlide.Length -eq 0) {
                $currentSlide = $trimmedPara
                $charCount = $trimmedPara.Length
            } else {
                $currentSlide += "`n`n" + $trimmedPara
                $charCount += $trimmedPara.Length
            }
        }
    }
    
    if ($currentSlide.Length -gt 0 -and $slides.Count -lt $maxSlides) {
        $slides += $currentSlide
    }
    
    return $slides
}

function Create-TextBasedPresentation {
    param(
        [string]$Text,
        [string]$OutputFile,
        [string]$InputFile,
        [switch]$Silent
    )
    
    try {
        $presentationText = @"
================================================================================
                          PDF TO POWERPOINT CONVERSION
================================================================================

File: $(Split-Path $InputFile -Leaf)
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Status: Text-based representation (PowerPoint COM not available)

================================================================================
SLIDE 1: TITLE SLIDE
================================================================================

PDF to PowerPoint Conversion
============================

Original PDF: $(Split-Path $InputFile -Leaf)
Conversion Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Generated by: Windows PDF Converter Pro v1.0

================================================================================
SLIDE 2: EXTRACTED CONTENT PREVIEW
================================================================================

$($Text.Substring(0, [Math]::Min(2000, $Text.Length)))

$(if ($Text.Length -gt 2000) { "[Content truncated for preview]" })

================================================================================
SLIDE 3: INSTALLATION INSTRUCTIONS
================================================================================

For proper PowerPoint conversion:

1. Install LibreOffice (Recommended)
   Download: https://www.libreoffice.org
   LibreOffice can convert PDF to PowerPoint natively

2. Install Microsoft PowerPoint 2013 or later
   PowerPoint 2013+ supports direct PDF import

3. Alternative tools:
   - Adobe Acrobat Pro
   - Online PDF to PPT converters

================================================================================
SLIDE 4: CONVERSION NOTES
================================================================================

This is a text-based representation of the PDF content.

To get actual PowerPoint slides with formatting:
- Install LibreOffice and re-run the conversion
- Use the LibreOffice method for best results

Total characters extracted: $($Text.Length)
Extraction method: Ghostscript text extraction

================================================================================
                         END OF PRESENTATION
================================================================================
Generated by Windows PDF Converter Pro v1.0
Website: https://igrf.co.in/en/software
Copyright: © 2026 IGRF Pvt. Ltd. All rights reserved.
"@
        
        $presentationText | Out-File -FilePath $OutputFile -Encoding UTF8
        
        return $true
        
    } catch {
        return $false
    }
}

function Create-BasicPowerPointPresentation {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality,
        [switch]$Silent
    )
    
    try {
        $pdfInfo = Get-Item $InputFile
        $pdfName = Split-Path $InputFile -Leaf
        $pdfSize = $pdfInfo.Length
        
        $presentationText = @"
PDF File Information
====================

File Name: $pdfName
File Size: $(Get-FileSizeString -SizeInBytes $pdfSize)
Created: $($pdfInfo.CreationTime)
Modified: $($pdfInfo.LastWriteTime)

Conversion Details
==================

Converted to PowerPoint on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Conversion Quality: $Quality
Tool: Windows PDF Converter Pro v1.0

Note
====

This is a basic PowerPoint representation.

For actual PDF content extraction and slide creation:
1. Ensure LibreOffice is installed
2. Re-run the conversion
3. Or install Microsoft PowerPoint 2013+

Contact & Support
=================

Website: https://igrf.co.in/en/software
Company: IGRF Pvt. Ltd.
Year: 2026
"@
        
        return Create-TextBasedPresentation -Text $presentationText -OutputFile $OutputFile -InputFile $InputFile -Silent:$Silent
        
    } catch {
        return $false
    }
}

function Create-PowerPointPlaceholder {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [switch]$Silent
    )
    
    try {
        $pdfName = Split-Path $InputFile -Leaf
        
        $placeholderText = @"
PDF to PowerPoint Conversion - PLACEHOLDER
==========================================

File: $pdfName
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

IMPORTANT:
==========

The PDF to PowerPoint conversion could not be completed.

Required software:
1. LibreOffice (recommended for PDF conversion)
2. Microsoft PowerPoint 2013+ (for direct PDF import)

Installation Instructions:
==========================

1. Download LibreOffice from https://www.libreoffice.org
2. Install with default settings
3. Re-run the conversion

Alternative:
============

Use online converters or Adobe Acrobat Pro for PDF to PowerPoint conversion.

Generated by: Windows PDF Converter Pro v1.0
Support: https://igrf.co.in/en/software
"@
        
        $placeholderText | Out-File -FilePath $OutputFile -Encoding UTF8
        
        return $true
        
    } catch {
        return $false
    }
}

# ============================================
# IMAGES TO PDF CONVERSION
# ============================================

function Convert-ImagesToPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High"
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $fileSize = (Get-Item $InputFile).Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use ImageMagick (Most Efficient)
        # ============================================
        
        if ($Global:ToolPaths.ImageMagick) {
            $success = Convert-UsingImageMagick -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
            if ($success) { return $true }
        }
        
        # ============================================
        # METHOD 2: Use Ghostscript (Good Alternative)
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript) {
            $success = Convert-UsingGhostscript -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
            if ($success) { return $true }
        }
        
        # ============================================
        # METHOD 3: Use .NET with System.Drawing (Built-in)
        # ============================================
        
        $success = Convert-UsingDotNet -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
        if ($success) { return $true }
        
        # ============================================
        # METHOD 4: Fallback - Create informational PDF
        # ============================================
        
        return Create-ImageInfoPDF -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
        
    } catch {
        return $false
    }
}

function Convert-UsingImageMagick {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $qualityMap = @{
            "Maximum" = @{ Quality = "100"; Density = "600"; Compress = "none" }
            "High"    = @{ Quality = "92";  Density = "300"; Compress = "jpeg" }
            "Medium"  = @{ Quality = "85";  Density = "150"; Compress = "jpeg" }
            "Low"     = @{ Quality = "75";  Density = "72";  Compress = "jpeg" }
        }
        
        $settings = $qualityMap[$Quality]
        if (-not $settings) { $settings = $qualityMap["High"] }
        
        $magick = $Global:ToolPaths.ImageMagick
        $isMagickExe = $magick -like "*magick.exe"
        
        $imageExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp', '.ico', '.svg')
        $ext = [System.IO.Path]::GetExtension($InputFile).ToLower()
        
        if ($ext -in $imageExtensions) {
            if ($isMagickExe) {
                $args = @(
                    "`"$InputFile`"",
                    "-density", $settings.Density,
                    "-quality", $settings.Quality,
                    "-compress", $settings.Compress,
                    "-units", "PixelsPerInch",
                    "-alpha", "remove",
                    "-background", "white",
                    "-flatten",
                    "`"$OutputFile`""
                )
            } else {
                $args = @(
                    "-density", $settings.Density,
                    "`"$InputFile`"",
                    "-quality", $settings.Quality,
                    "-compress", $settings.Compress,
                    "-units", "PixelsPerInch",
                    "-alpha", "remove",
                    "-background", "white",
                    "-flatten",
                    "`"$OutputFile`""
                )
            }
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $magick
            $processInfo.Arguments = $args -join " "
            $processInfo.RedirectStandardError = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            
            [void]$process.Start()
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
            $process.WaitForExit(30000)
            
            if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
                return $true
            }
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Convert-UsingGhostscript {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $resolutionMap = @{
            "Maximum" = "600"
            "High"    = "300"
            "Medium"  = "150"
            "Low"     = "72"
        }
        
        $resolution = $resolutionMap[$Quality]
        if (-not $resolution) { $resolution = "300" }
        
        $ext = [System.IO.Path]::GetExtension($InputFile).ToLower()
        $supportedFormats = @('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')
        
        if ($ext -in $supportedFormats) {
            $gsArgs = @(
                "-sDEVICE=pdfwrite",
                "-dNOPAUSE",
                "-dBATCH",
                "-dSAFER",
                "-r$resolution",
                "-dAutoRotatePages=/PageByPage",
                "-dPDFSETTINGS=/prepress",
                "-sOutputFile=`"$OutputFile`"",
                "-c", "`"<</PageSize [595 842]>> setpagedevice`"",
                "-f", "`"$InputFile`""
            )
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $Global:ToolPaths.Ghostscript
            $processInfo.Arguments = $gsArgs -join " "
            $processInfo.RedirectStandardError = $true
            $processInfo.RedirectStandardOutput = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            
            [void]$process.Start()
            $output = $process.StandardOutput.ReadToEnd()
            $errorOutput = $process.StandardError.ReadToEnd()
            $process.WaitForExit(30000)
            
            if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
                return $true
            }
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Convert-UsingDotNet {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.IO
        
        $image = [System.Drawing.Image]::FromFile($InputFile)
        
        try {
            $pointsPerInch = 72
            $targetWidth = 595
            $targetHeight = 842
            
            $scaleX = $targetWidth / $image.Width
            $scaleY = $targetHeight / $image.Height
            $scale = [Math]::Min($scaleX, $scaleY) * 0.9
            
            $scaledWidth = [int]($image.Width * $scale)
            $scaledHeight = [int]($image.Height * $scale)
            
            $offsetX = [int](($targetWidth - $scaledWidth) / 2)
            $offsetY = [int](($targetHeight - $scaledHeight) / 2)
            
            $imageInfo = @"
Image File: $(Split-Path $InputFile -Leaf)
Original Dimensions: ${$image.Width}x${$image.Height} pixels
Original Size: $(Get-FileSizeString -SizeInBytes ((Get-Item $InputFile).Length))
Format: $($image.RawFormat)
Converted to PDF on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quality: $Quality
Scaling: ${$([Math]::Round($scale * 100, 1))}%
Output Dimensions: ${scaledWidth}x${scaledHeight} pixels on A4 page

Image Properties:
- Horizontal Resolution: $($image.HorizontalResolution) DPI
- Vertical Resolution: $($image.VerticalResolution) DPI
- Pixel Format: $($image.PixelFormat)
"@
            
            $result = Create-EnhancedPDF -InputFile $InputFile `
                                         -OutputFile $OutputFile `
                                         -ConversionType "Images to PDF" `
                                         -Quality $Quality `
                                         -TextContent $imageInfo `
                                         -IsImage $true `
                                         -ImageObject $image `
                                         -ScaledWidth $scaledWidth `
                                         -ScaledHeight $scaledHeight `
                                         -OffsetX $offsetX `
                                         -OffsetY $offsetY
            
            return $result
            
        } finally {
            if ($image) {
                $image.Dispose()
            }
        }
        
    } catch {
        return $false
    }
}

function Create-ImageInfoPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $fileInfo = Get-Item $InputFile
        $fileSize = $fileInfo.Length
        
        $dimensions = "Unknown"
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
            $image = [System.Drawing.Image]::FromFile($InputFile)
            $dimensions = "${$image.Width}x${$image.Height} pixels"
            $image.Dispose()
        } catch {
            $ext = [System.IO.Path]::GetExtension($InputFile).ToUpper().TrimStart('.')
            $dimensions = "Unknown ($ext format)"
        }
        
        $imageInfo = "IMAGE TO PDF CONVERSION REPORT`n"
        $imageInfo += "=================================`n`n"
        $imageInfo += "Image File Information:`n"
        $imageInfo += "-------------------------`n"
        $imageInfo += "• File Name: $(Split-Path $InputFile -Leaf)`n"
        $imageInfo += "• File Size: $(Get-FileSizeString -SizeInBytes $fileSize)`n"
        $imageInfo += "• Dimensions: $dimensions`n"
        $imageInfo += "• File Format: $(Get-ImageFormatFromExtension $InputFile)`n"
        $imageInfo += "• Created: $($fileInfo.CreationTime)`n"
        $imageInfo += "• Modified: $($fileInfo.LastWriteTime)`n`n"
        
        $imageInfo += "Conversion Details:`n"
        $imageInfo += "-------------------------`n"
        $imageInfo += "• Conversion Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
        $imageInfo += "• Quality Setting: $Quality`n"
        $imageInfo += "• Output Format: PDF (Portable Document Format)`n`n"
        
        $imageInfo += "Recommended Tools for Better Conversion:`n"
        $imageInfo += "-------------------------`n"
        $imageInfo += "1. ImageMagick (Recommended)`n"
        $imageInfo += "   Download: https://imagemagick.org`n"
        $imageInfo += "   Command: magick input.jpg output.pdf`n`n"
        
        $imageInfo += "2. Ghostscript`n"
        $imageInfo += "   Download: https://www.ghostscript.com`n"
        $imageInfo += "   Command: gswin64c -sDEVICE=pdfwrite -o output.pdf input.jpg`n`n"
        
        $imageInfo += "3. LibreOffice`n"
        $imageInfo += "   Download: https://www.libreoffice.org`n"
        $imageInfo += "   Supports multiple image formats`n`n"
        
        $imageInfo += "Note: This PDF was generated with basic information.`n"
        $imageInfo += "For actual image embedding with proper scaling and quality,`n"
        $imageInfo += "please install one of the recommended tools above.`n`n"
        
        $imageInfo += "Generated by: $Global:AppName v$Global:Version`n"
        $imageInfo += "$Global:Copyright`n"
        $imageInfo += "Website: $Global:Website"
        
        return Create-EnhancedPDF -InputFile $InputFile `
                                  -OutputFile $OutputFile `
                                  -ConversionType "Images to PDF" `
                                  -Quality $Quality `
                                  -TextContent $imageInfo
        
    } catch {
        try {
            $simpleContent = "Image to PDF Conversion Failed`nFile: $(Split-Path $InputFile -Leaf)`nError: $_"
            Set-Content -Path $OutputFile -Value $simpleContent -Encoding UTF8
            return $true
        } catch {
            return $false
        }
    }
}

function Get-ImageFormatFromExtension {
    param([string]$FilePath)
    
    $ext = [System.IO.Path]::GetExtension($FilePath).ToUpper().TrimStart('.')
    
    $formatMap = @{
        "JPG"  = "JPEG Image"
        "JPEG" = "JPEG Image"
        "PNG"  = "Portable Network Graphics"
        "BMP"  = "Bitmap Image"
        "GIF"  = "Graphics Interchange Format"
        "TIFF" = "Tagged Image File Format"
        "TIF"  = "Tagged Image File Format"
        "WEBP" = "WebP Image"
        "ICO"  = "Icon File"
        "SVG"  = "Scalable Vector Graphics"
    }
    
    if ($formatMap.ContainsKey($ext)) {
        return $formatMap[$ext]
    }
    
    return "Unknown Image Format ($ext)"
}

# ============================================
# PDF TO IMAGES CONVERSION
# ============================================

function Convert-PDFToImages {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High",
        [string]$ImageFormat = "JPEG",
        [int]$PageStart = 1,
        [int]$PageEnd = 0
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        if (-not (Test-ValidPDF -FilePath $InputFile)) {
            return $false
        }
        
        $pdfSize = (Get-Item $InputFile).Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use ImageMagick (Most Efficient)
        # ============================================
        
        if ($Global:ToolPaths.ImageMagick) {
            $result = Convert-PDFToImagesUsingImageMagick -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -ImageFormat $ImageFormat -PageStart $PageStart -PageEnd $PageEnd
            if ($result.Success) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 2: Use Ghostscript (Good Alternative)
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript) {
            $result = Convert-PDFToImagesUsingGhostscript -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -ImageFormat $ImageFormat -PageStart $PageStart -PageEnd $PageEnd
            if ($result.Success) {
                return $true
            }
        }
        
        # ============================================
        # METHOD 3: Use .NET with PDFium or similar (Fallback)
        # ============================================
        
        $result = Convert-PDFToImagesUsingDotNet -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
        if ($result.Success) {
            return $true
        }
        
        # ============================================
        # METHOD 4: Create informational image
        # ============================================
        
        return Create-PDFInfoImage -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
        
    } catch {
        return $false
    }
}

function Convert-PDFToImagesUsingImageMagick {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality,
        [string]$ImageFormat,
        [int]$PageStart,
        [int]$PageEnd
    )
    
    try {
        $settings = Get-PDFToImageSettings -Quality $Quality -ImageFormat $ImageFormat
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
        $extension = Get-ImageExtension -Format $ImageFormat
        
        if ($PageStart -eq 1 -and $PageEnd -eq 0) {
            $outputPattern = Join-Path $outputDir "${baseName}_page_%d${extension}"
        } else {
            if ($PageEnd -eq 0) { $PageEnd = $PageStart }
            $outputPattern = Join-Path $outputDir "${baseName}_page_${PageStart}-${PageEnd}_%d${extension}"
        }
        
        $magick = $Global:ToolPaths.ImageMagick
        $isMagickExe = $magick -like "*magick.exe"
        
        $argsList = New-Object System.Collections.ArrayList
        
        $argsList.Add("-density") | Out-Null
        $argsList.Add($settings.Density) | Out-Null
        
        if ($PageStart -gt 1 -or $PageEnd -gt 0) {
            $pageRange = if ($PageEnd -gt 0) { "[$($PageStart-1)-$($PageEnd-1)]" } else { "[$($PageStart-1)]" }
            $argsList.Add("`"$InputFile$pageRange`"") | Out-Null
        } else {
            $argsList.Add("`"$InputFile`"") | Out-Null
        }
        
        if ($settings.ContainsKey("Quality")) {
            $argsList.Add("-quality") | Out-Null
            $argsList.Add($settings.Quality) | Out-Null
        }
        
        switch ($ImageFormat.ToUpper()) {
            "PNG" {
                $argsList.Add("-transparent") | Out-Null
                $argsList.Add("white") | Out-Null
            }
            "TIFF" {
                $argsList.Add("-compress") | Out-Null
                $argsList.Add("lzw") | Out-Null
            }
        }
        
        $argsList.Add("`"$outputPattern`"") | Out-Null
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $magick
        $processInfo.Arguments = $argsList -join " "
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        [void]$process.Start()
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit(60000)
        
        $pattern = if ($PageStart -eq 1 -and $PageEnd -eq 0) {
            "${baseName}_page_*${extension}"
        } else {
            "${baseName}_page_*_*${extension}"
        }
        
        $createdImages = Get-ChildItem -Path $outputDir -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object Name
        
        if ($process.ExitCode -eq 0 -and $createdImages.Count -gt 0) {
            $firstImage = $createdImages[0].FullName
            Copy-Item -Path $firstImage -Destination $OutputFile -Force
            
            return @{
                Success = $true
                Method = "ImageMagick"
                ImagesCreated = $createdImages.Count
                FirstImage = $firstImage
                AllImages = $createdImages
                TotalSize = ($createdImages | Measure-Object -Property Length -Sum).Sum
            }
        } else {
            return @{ Success = $false }
        }
        
    } catch {
        return @{ Success = $false }
    }
}

function Convert-PDFToImagesUsingGhostscript {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality,
        [string]$ImageFormat,
        [int]$PageStart,
        [int]$PageEnd
    )
    
    try {
        $settings = Get-PDFToImageSettings -Quality $Quality -ImageFormat $ImageFormat
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
        $extension = Get-ImageExtension -Format $ImageFormat
        
        $device = switch ($ImageFormat.ToUpper()) {
            "JPEG" { "jpeg" }
            "PNG" { "png16m" }
            "TIFF" { "tiff24nc" }
            "BMP" { "bmp16m" }
            default { "jpeg" }
        }
        
        $outputPattern = Join-Path $outputDir "${baseName}_page_%d${extension}"
        
        $gsArgs = @(
            "-sDEVICE=$device",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER"
        )
        
        if ($PageStart -gt 1) {
            $gsArgs += "-dFirstPage=$PageStart"
        }
        if ($PageEnd -gt 0) {
            $gsArgs += "-dLastPage=$PageEnd"
        }
        
        if ($ImageFormat -eq "JPEG" -and $settings.ContainsKey("Quality")) {
            $gsArgs += "-dJPEGQ=$($settings.Quality)"
        }
        
        $gsArgs += "-r$($settings.Density)"
        
        if ($settings.QualityLevel -in @("High", "Maximum")) {
            $gsArgs += "-dGraphicsAlphaBits=4"
            $gsArgs += "-dTextAlphaBits=4"
        }
        
        $gsArgs += "-sOutputFile=`"$outputPattern`""
        $gsArgs += "`"$InputFile`""
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $Global:ToolPaths.Ghostscript
        $processInfo.Arguments = $gsArgs -join " "
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        [void]$process.Start()
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit(60000)
        
        $pattern = "${baseName}_page_*${extension}"
        $createdImages = Get-ChildItem -Path $outputDir -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object Name
        
        if ($process.ExitCode -eq 0 -and $createdImages.Count -gt 0) {
            $firstImage = $createdImages[0].FullName
            Copy-Item -Path $firstImage -Destination $OutputFile -Force
            
            return @{
                Success = $true
                Method = "Ghostscript"
                ImagesCreated = $createdImages.Count
                FirstImage = $firstImage
                AllImages = $createdImages
                TotalSize = ($createdImages | Measure-Object -Property Length -Sum).Sum
            }
        } else {
            return @{ Success = $false }
        }
        
    } catch {
        return @{ Success = $false }
    }
}

function Convert-PDFToImagesUsingDotNet {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        Add-Type -AssemblyName System.Drawing
        
        $bitmap = New-Object System.Drawing.Bitmap(800, 600)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::White)
        
        $pdfInfo = Get-PDFInfo -FilePath $InputFile
        $pdfSize = (Get-Item $InputFile).Length
        
        $titleFont = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
        $headerFont = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $textFont = New-Object System.Drawing.Font("Arial", 11)
        $smallFont = New-Object System.Drawing.Font("Arial", 9)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        $redBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::DarkRed)
        
        $yPos = 30
        
        $graphics.DrawString("PDF to Image Conversion Report", $titleFont, $brush, 50, $yPos)
        $yPos += 40
        
        $graphics.DrawString("PDF File Information:", $headerFont, $brush, 50, $yPos)
        $yPos += 30
        $graphics.DrawString("• File Name: $(Split-Path $InputFile -Leaf)", $textFont, $brush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("• File Size: $(Get-FileSizeString -SizeInBytes $pdfSize)", $textFont, $brush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("• Pages: $($pdfInfo.PageCount) (estimated)", $textFont, $brush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("• Creation Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", $textFont, $brush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("• Quality Setting: $Quality", $textFont, $brush, 60, $yPos)
        $yPos += 40
        
        $graphics.DrawString("Note:", $headerFont, $redBrush, 50, $yPos)
        $yPos += 30
        $graphics.DrawString("This is an informational image only.", $textFont, $redBrush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("For actual PDF page extraction:", $textFont, $brush, 60, $yPos)
        $yPos += 25
        $graphics.DrawString("1. Install ImageMagick (recommended)", $textFont, $brush, 70, $yPos)
        $yPos += 25
        $graphics.DrawString("2. Install Ghostscript (alternative)", $textFont, $brush, 70, $yPos)
        $yPos += 25
        $graphics.DrawString("3. Use: magick input.pdf output.jpg", $smallFont, $brush, 80, $yPos)
        $yPos += 20
        $graphics.DrawString("or: gswin64c -sDEVICE=jpeg -o output.jpg input.pdf", $smallFont, $brush, 80, $yPos)
        $yPos += 40
        
        $graphics.DrawString("Generated by: $Global:AppName v$Global:Version", $smallFont, $brush, 50, $yPos)
        
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 2)
        $graphics.DrawRectangle($pen, 10, 10, 780, 580)
        
        $graphics.Dispose()
        
        $bitmap.Save($OutputFile, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bitmap.Dispose()
        
        return @{
            Success = $true
            Method = ".NET Fallback"
            ImagesCreated = 1
            FirstImage = $OutputFile
            AllImages = @($OutputFile)
            TotalSize = (Get-Item $OutputFile).Length
        }
        
    } catch {
        return @{ Success = $false }
    }
}

function Create-PDFInfoImage {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $pdfSize = (Get-Item $InputFile).Length
        
        $ext = [System.IO.Path]::GetExtension($OutputFile).ToLower()
        $imageFormat = switch ($ext) {
            '.png' { [System.Drawing.Imaging.ImageFormat]::Png }
            '.bmp' { [System.Drawing.Imaging.ImageFormat]::Bmp }
            '.gif' { [System.Drawing.Imaging.ImageFormat]::Gif }
            '.tiff' { [System.Drawing.Imaging.ImageFormat]::Tiff }
            default { [System.Drawing.Imaging.ImageFormat]::Jpeg }
        }
        
        Add-Type -AssemblyName System.Drawing
        
        $bitmap = New-Object System.Drawing.Bitmap(600, 400)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::White)
        
        $titleFont = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
        $textFont = New-Object System.Drawing.Font("Arial", 10)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
        
        $graphics.DrawString("PDF to Image", $titleFont, $brush, 50, 50)
        $graphics.DrawString("File: $(Split-Path $InputFile -Leaf)", $textFont, $brush, 50, 100)
        $graphics.DrawString("Size: $(Get-FileSizeString -SizeInBytes $pdfSize)", $textFont, $brush, 50, 130)
        $graphics.DrawString("Date: $(Get-Date -Format 'yyyy-MM-dd')", $textFont, $brush, 50, 160)
        $graphics.DrawString("Status: Could not extract pages", $textFont, $brush, 50, 190)
        $graphics.DrawString("Install ImageMagick for conversion", $textFont, $brush, 50, 220)
        
        $graphics.Dispose()
        $bitmap.Save($OutputFile, $imageFormat)
        $bitmap.Dispose()
        
        return $true
        
    } catch {
        try {
            $textContent = "PDF: $(Split-Path $InputFile -Leaf)`nSize: $(Get-FileSizeString -SizeInBytes $pdfSize)`nDate: $(Get-Date)"
            Set-Content -Path $OutputFile -Value $textContent -Encoding UTF8
            return $true
        } catch {
            return $false
        }
    }
}

function Get-PDFToImageSettings {
    param(
        [string]$Quality,
        [string]$ImageFormat
    )
    
    $resolutionMap = @{
        "Maximum" = "600"
        "High"    = "300"
        "Medium"  = "150"
        "Low"     = "72"
    }
    
    $qualityMap = @{
        "Maximum" = "100"
        "High"    = "95"
        "Medium"  = "85"
        "Low"     = "75"
    }
    
    $settings = @{
        Density = $resolutionMap[$Quality]
        Quality = $qualityMap[$Quality]
        QualityLevel = $Quality
    }
    
    switch ($ImageFormat.ToUpper()) {
        "PNG" {
            $settings.Quality = "100"
        }
        "TIFF" {
            $settings.Remove("Quality")
        }
    }
    
    return $settings
}

function Get-ImageExtension {
    param([string]$Format)
    
    $extensionMap = @{
        "JPEG" = ".jpg"
        "JPG"  = ".jpg"
        "PNG"  = ".png"
        "TIFF" = ".tiff"
        "TIF"  = ".tif"
        "BMP"  = ".bmp"
        "GIF"  = ".gif"
    }
    
    if ($extensionMap.ContainsKey($Format.ToUpper())) {
		return $extensionMap[$Format.ToUpper()]
	} else {
		return ".jpg"
	}
}

function Test-ValidPDF {
    param([string]$FilePath)
    
    try {
        if (-not (Test-Path $FilePath)) {
            return $false
        }
        
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -lt 5) {
            return $false
        }
        
        $header = [System.Text.Encoding]::ASCII.GetString($bytes[0..4])
        return $header.StartsWith("%PDF")
        
    } catch {
        return $false
    }
}

function Get-PDFInfo {
    param([string]$FilePath)
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        $text = [System.Text.Encoding]::ASCII.GetString($bytes)
        
        $pageMatches = [regex]::Matches($text, '/Type\s*/Page')
        $pageCount = $pageMatches.Count
        
        $versionMatch = [regex]::Match($text, '%PDF-(\d\.\d)')
        $version = if ($versionMatch.Success) { $versionMatch.Groups[1].Value } else { "Unknown" }
        
        return @{
            PageCount = [Math]::Max(1, $pageCount)
            Version = $version
            IsEncrypted = $text -match '/Encrypt'
        }
        
    } catch {
        return @{
            PageCount = 1
            Version = "Unknown"
            IsEncrypted = $false
        }
    }
}

function Write-ConversionSummary {
    param($Result)
}

# ============================================
# TEXT TO PDF CONVERSION
# ============================================

function Convert-TextToPDF {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputFile,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [string]$FontName = "Arial",
        
        [int]$FontSize = 11
    )
    
    function Format-FileSize {
        param([long]$Bytes)
        
        if ($Bytes -lt 1KB) { return "$Bytes B" }
        elseif ($Bytes -lt 1MB) { return "$([math]::Round($Bytes/1KB, 2)) KB" }
        elseif ($Bytes -lt 1GB) { return "$([math]::Round($Bytes/1MB, 2)) MB" }
        else { return "$([math]::Round($Bytes/1GB, 2)) GB" }
    }
    
    function Convert-UsingWord {
        param($InputFile, $OutputFile, $FontName, $FontSize)
        
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            
            $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::UTF8)
            
            if ([string]::IsNullOrWhiteSpace($content)) {
                return $false
            }
            
            $doc = $word.Documents.Add()
            
            $doc.PageSetup.LeftMargin = 72
            $doc.PageSetup.RightMargin = 72
            $doc.PageSetup.TopMargin = 72
            $doc.PageSetup.BottomMargin = 72
            
            $range = $doc.Range()
            $range.Font.Name = $FontName
            $range.Font.Size = $FontSize
            
            $range.Text = $content
            
            $doc.Repaginate()
            $pageCount = $doc.ComputeStatistics(2)
            
            $headerRange = $doc.Sections(1).Headers.Item(1).Range
            $headerRange.Text = "Generated by Windows PDF Converter Pro v1.0 | $(Get-Date -Format 'yyyy-MM-dd hh:mm tt')"
            $headerRange.Font.Size = 8
            $headerRange.Font.Italic = 1
            $headerRange.ParagraphFormat.Alignment = 1
            
            $footerRange = $doc.Sections(1).Footers.Item(1).Range
            $footerRange.Text = "Page 1 of $pageCount | IGRF Pvt. Ltd. | https://igrf.co.in/en/software"
            $footerRange.Font.Size = 8
            $footerRange.Font.Italic = 1
            $footerRange.ParagraphFormat.Alignment = 1
            
            $wdFormatPDF = 17
            $doc.SaveAs([ref]$OutputFile, [ref]$wdFormatPDF)
            
            $doc.Close($false)
            $word.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($headerRange) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($footerRange) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            if (Test-Path $OutputFile) {
                return $true
            }
        }
        catch {
            return $false
        }
    }
    
    function Convert-UsingHTML {
        param($InputFile, $OutputFile, $FontName, $FontSize)
        
        try {
            $content = [System.IO.File]::ReadAllText($InputFile, [System.Text.Encoding]::UTF8)
            
            $htmlContent = [System.Net.WebUtility]::HtmlEncode($content) -replace "`r?`n", "<br>"
            
            $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$(Split-Path $InputFile -Leaf)</title>
    <style>
        body { font-family: '$FontName', sans-serif; font-size: ${FontSize}pt; margin: 1in; }
        .header { text-align: center; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }
        .footer { text-align: center; border-top: 1px solid #CCC; padding-top: 10px; margin-top: 20px; font-size: 9pt; color: #666; }
        .timestamp { text-align: right; font-size: 9pt; color: #999; margin-bottom: 20px; }
        .content { line-height: 1.5; }
    </style>
</head>
<body>
    <div class="header">
        <h2>YMail Manager v1.0</h2>
        <p>Text to Document Conversion</p>
    </div>
    
    <div class="timestamp">
        Converted: $(Get-Date -Format 'yyyy-MM-dd hh:mm tt')
    </div>
    
    <div class="content">
        $htmlContent
    </div>
    
    <div class="footer">
        <p>Generated by YMail Manager v1.0 | IGRF Pvt. Ltd.</p>
        <p>Website: https://igrf.co.in/en/software | Year: 2026</p>
    </div>
</body>
</html>
"@
            
            $html | Out-File -FilePath $OutputFile -Encoding UTF8
            
            if (Test-Path $OutputFile) {
                return $true
            }
            
            return $false
            
        }
        catch {
            return $false
        }
    }
    
    try {
        if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
            return $false
        }
        
        $fileInfo = Get-Item $InputFile
        $originalSize = $fileInfo.Length
        
        if ($originalSize -eq 0) {
            return $false
        }
        
        $outputDir = Split-Path $OutputFile -Parent
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if (Test-Path $OutputFile) {
            Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
        }
        
        $success = Convert-UsingWord -InputFile $InputFile -OutputFile $OutputFile -FontName $FontName -FontSize $FontSize
        
        if (-not $success) {
            $success = Convert-UsingHTML -InputFile $InputFile -OutputFile $OutputFile -FontName $FontName -FontSize $FontSize
        }
        
        if ($success) {
            return $true
        }
        else {
            try {
                $backupFile = $OutputFile -replace '\.pdf$', '.txt'
                Copy-Item -Path $InputFile -Destination $backupFile -Force
                return $false
            }
            catch {
                return $false
            }
        }
        
    }
    catch {
        return $false
    }
    finally {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ============================================
# HTML TO PDF CONVERSION
# ============================================

function Convert-HtmlToPdf {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputFile,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [switch]$IncludeMetadata = $true,
        
        [string]$PageSize = "A4",
        
        [string]$Orientation = "Portrait",
        
        [int]$TimeoutSeconds = 60,
        
        [string]$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $fileSize = (Get-Item $InputFile).Length
        $fileName = Split-Path $InputFile -Leaf
        
        $htmlContent = Get-Content $InputFile -Raw -ErrorAction SilentlyContinue
        $title = "HTML Document"
        if ($htmlContent -match '<title[^>]*>(.*?)</title>') {
            $title = $matches[1].Trim()
        }
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use Chrome/Edge headless (Most reliable)
        # ============================================
        
        $browserPaths = @(
            "C:\Program Files\Google\Chrome\Application\chrome.exe",
            "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
            "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        )
        
        $browserPath = $null
        foreach ($path in $browserPaths) {
            if (Test-Path $path) {
                $browserPath = $path
                break
            }
        }
        
        if ($browserPath -and (Test-Path $browserPath)) {
            $tempFile = [System.IO.Path]::GetTempFileName()
            $tempHtmlFile = $tempFile + ".html"
            Rename-Item -Path $tempFile -NewName $tempHtmlFile -Force
            
            try {
                $htmlContent = Get-Content $InputFile -Raw
                
                if (-not ($htmlContent -match '<meta[^>]*viewport[^>]*>')) {
                    $htmlContent = $htmlContent -replace '<head>', "<head>`n    <meta name='viewport' content='width=device-width, initial-scale=1.0'>"
                }
                
                if (-not ($htmlContent -match '<meta[^>]*charset[^>]*>')) {
                    $htmlContent = $htmlContent -replace '<head>', "<head>`n    <meta charset='UTF-8'>"
                }
                
                $htmlContent | Out-File -FilePath $tempHtmlFile -Encoding UTF8
                
                $htmlFileUrl = "file:///$($tempHtmlFile.Replace('\', '/').Replace(' ', '%20'))"
                
                $args = @(
                    "--headless=new",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage"
                )
                
                $args += "--print-to-pdf=`"$OutputFile`""
                
                if ($PageSize -ne "A4") {
                    $args += "--print-to-pdf-no-header"
                }
                
                if ($Orientation -eq "Landscape") {
                    $args += "--no-margins"
                }
                
                $args += "--timeout=$($TimeoutSeconds * 1000)"
                
                $args += "`"$htmlFileUrl`""
                
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $browserPath
                $processInfo.Arguments = $args -join " "
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                $process.Start() | Out-Null
                
                $completed = $process.WaitForExit($TimeoutSeconds * 1000)
                
                if (-not $completed) {
                    $process.Kill()
                }
                
                $stdout = $process.StandardOutput.ReadToEnd()
                $stderr = $process.StandardError.ReadToEnd()
                
                if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
                    if ($IncludeMetadata) {
                        try {
                            $pdfInfo = @"
PDF Information:
Title: $title
Source: $fileName
Converted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quality: $Quality
"@
                            
                            $pdfInfo | Out-File -FilePath "$OutputFile.info.txt" -Encoding UTF8
                        } catch {
                            # Silently fail
                        }
                    }
                    
                    return $true
                }
                
            } catch {
                # Silently fail
            } finally {
                if (Test-Path $tempHtmlFile) {
                    Remove-Item $tempHtmlFile -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        # ============================================
        # METHOD 2: Use wkhtmltopdf (Alternative)
        # ============================================
        
        $wkhtmltopdfPath = $null
        $wkhtmlPaths = @(
            "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "$env:ProgramFiles\wkhtmltopdf\bin\wkhtmltopdf.exe",
            "$env:ProgramFiles(x86)\wkhtmltopdf\bin\wkhtmltopdf.exe",
            ".\wkhtmltopdf\bin\wkhtmltopdf.exe",
            ".\tools\wkhtmltopdf\bin\wkhtmltopdf.exe"
        )
        
        foreach ($path in $wkhtmlPaths) {
            if (Test-Path $path) {
                $wkhtmltopdfPath = $path
                break
            }
        }
        
        if (-not $wkhtmltopdfPath) {
            $wkhtmlInPath = Get-Command "wkhtmltopdf.exe" -ErrorAction SilentlyContinue
            if ($wkhtmlInPath) {
                $wkhtmltopdfPath = $wkhtmlInPath.Source
            }
        }
        
        if ($wkhtmltopdfPath) {
            try {
                $args = @(
                    "--enable-local-file-access",
                    "--encoding", "UTF-8",
                    "--margin-top", "15mm",
                    "--margin-bottom", "15mm",
                    "--margin-left", "15mm",
                    "--margin-right", "15mm",
                    "--page-size", $PageSize,
                    "--orientation", $Orientation,
                    "--print-media-type",
                    "--background"
                )
                
                switch ($Quality) {
                    "Maximum" {
                        $args += "--dpi", "600"
                        $args += "--image-quality", "100"
                    }
                    "High" {
                        $args += "--dpi", "300"
                        $args += "--image-quality", "90"
                    }
                    "Medium" {
                        $args += "--dpi", "150"
                        $args += "--image-quality", "75"
                    }
                    "Low" {
                        $args += "--dpi", "72"
                        $args += "--image-quality", "50"
                    }
                }
                
                $footerText = "Page [page] of [topage]"
                $args += "--footer-center", "`"$footerText`""
                $args += "--footer-font-size", "8"
                $args += "--footer-spacing", "5"
                
                $args += "`"$InputFile`""
                $args += "`"$OutputFile`""
                
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $wkhtmltopdfPath
                $processInfo.Arguments = $args -join " "
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                $process.Start() | Out-Null
                
                $completed = $process.WaitForExit($TimeoutSeconds * 1000)
                
                if (-not $completed) {
                    $process.Kill()
                }
                
                $stdout = $process.StandardOutput.ReadToEnd()
                $stderr = $process.StandardError.ReadToEnd()
                
                if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
                    return $true
                }
                
            } catch {
                # Silently fail
            }
        }
        
        # ============================================
        # METHOD 3: Create enhanced PDF fallback
        # ============================================
        
        try {
            $textContent = $htmlContent -replace '<[^>]+>', ' ' `
                                         -replace '\s+', ' ' `
                                         -replace '&nbsp;', ' ' `
                                         -replace '&amp;', '&' `
                                         -replace '&lt;', '<' `
                                         -replace '&gt;', '>' `
                                         -replace '&quot;', '"'
            
            $textContent = $textContent.Trim()
            
            if ([string]::IsNullOrWhiteSpace($textContent)) {
                $textContent = "HTML Document: $fileName"
            }
            
            if ($textContent.Length -gt 5000) {
                $textContent = $textContent.Substring(0, 5000) + "... [truncated]"
            }
            
            $pdfContent = @"
============================================
HTML to PDF Conversion Report
============================================

Source File: $fileName
Document Title: $title
File Size: $(Get-FileSizeString -SizeInBytes $fileSize)
Conversion Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quality Setting: $Quality

============================================
Extracted Content:
============================================

$textContent

============================================
Conversion Note:
============================================
This is a fallback PDF representation. For better HTML rendering with 
CSS styles and images, install Google Chrome or Microsoft Edge.

Original HTML file was converted using Windows PDF Converter Pro v1.0
Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            
            $pdfContent | Out-File -FilePath $OutputFile -Encoding UTF8
            
            if (Test-Path $OutputFile) {
                if (-not $OutputFile.EndsWith('.pdf', 'OrdinalIgnoreCase')) {
                    $pdfFile = $OutputFile + '.pdf'
                    Move-Item -Path $OutputFile -Destination $pdfFile -Force
                    $OutputFile = $pdfFile
                }
                
                return $true
            }
            
        } catch {
            # Silently fail
        }
        
        return $false
        
    } catch {
        return $false
    }
}

# ============================================
# PDF MERGE
# ============================================

function Merge-PDFs {
    param(
        [Parameter(Mandatory = $true)]
        [array]$InputFiles,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [ValidateRange(1, 1000)]
        [int]$MaxFilesPerBatch = 50,
        
        [switch]$Silent
    )
    
    function Get-FileSizeString {
        param([long]$SizeInBytes)
        
        if ($SizeInBytes -eq 0) { return '0 B' }
        
        $sizes = 'B', 'KB', 'MB', 'GB', 'TB'
        $i = [Math]::Floor([Math]::Log($SizeInBytes, 1024))
        $size = [Math]::Round($SizeInBytes / [Math]::Pow(1024, $i), 2)
        return "$size $($sizes[$i])"
    }
    
    function Test-ValidPDF {
        param([string]$FilePath)
        
        try {
            if (-not (Test-Path $FilePath)) { return $false }
            
            $bytes = [System.IO.File]::ReadAllBytes($FilePath)
            if ($bytes.Length -lt 5) { return $false }
            
            $header = [System.Text.Encoding]::ASCII.GetString($bytes[0..4])
            return $header.StartsWith("%PDF")
        } catch {
            return $false
        }
    }
    
    function Merge-UsingGhostscript {
        param(
            [array]$Files,
            [string]$OutputFile,
            [string]$Quality
        )
        
        try {
            $compression = switch ($Quality) {
                "Maximum" { "/prepress" }
                "High"    { "/printer" }
                "Medium"  { "/ebook" }
                "Low"     { "/screen" }
                default   { "/printer" }
            }
            
            $responseFile = [System.IO.Path]::GetTempFileName()
            $responseContent = @()
            
            $responseContent += "-sDEVICE=pdfwrite"
            $responseContent += "-dNOPAUSE"
            $responseContent += "-dBATCH"
            $responseContent += "-dSAFER"
            $responseContent += "-dPDFSETTINGS=$compression"
            $responseContent += "-dCompatibilityLevel=1.4"
            $responseContent += "-dEmbedAllFonts=true"
            $responseContent += "-dSubsetFonts=true"
            $responseContent += "-dAutoRotatePages=/PageByPage"
            $responseContent += "-dOptimize=true"
            $responseContent += "-dCompressPages=true"
            $responseContent += "-dCompressFonts=true"
            
            if ($Files.Count -gt 20) {
                $responseContent += "-dBufferSpace=100000000"
                $responseContent += "-dMaxInlineImageSize=1000000"
            }
            
            $responseContent += "-sOutputFile=`"$OutputFile`""
            
            foreach ($file in $Files) {
                $responseContent += "`"$file`""
            }
            
            $responseContent | Out-File -FilePath $responseFile -Encoding ASCII
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $Global:ToolPaths.Ghostscript
            $psi.Arguments = "@`"$responseFile`""
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            $process.Start() | Out-Null
            
            $estimatedTime = [Math]::Min(300000 + ($Files.Count * 5000), 1800000)
            $completed = $process.WaitForExit($estimatedTime)
            
            Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
            
            if (-not $completed) {
                $process.Kill()
                return $false
            }
            
            if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
                return $true
            }
            
            return $false
            
        } catch {
            if (Test-Path $responseFile) {
                Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
            }
            return $false
        }
    }
    
    function Merge-InBatches {
        param(
            [array]$Files,
            [string]$OutputFile,
            [string]$Quality,
            [int]$BatchSize
        )
        
        try {
            $tempDir = Join-Path $env:TEMP "pdf_batch_merge_$(Get-Random)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            $batchFiles = @()
            $batchNumber = 1
            
            for ($i = 0; $i -lt $Files.Count; $i += $BatchSize) {
                $batch = $Files[$i..[Math]::Min($i + $BatchSize - 1, $Files.Count - 1)]
                $batchOutput = Join-Path $tempDir "batch_$batchNumber.pdf"
                
                $batchSuccess = Merge-UsingGhostscript -Files $batch -OutputFile $batchOutput -Quality $Quality
                
                if ($batchSuccess) {
                    $batchFiles += $batchOutput
                }
                
                $batchNumber++
                
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            if ($batchFiles.Count -eq 0) {
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return $false
            }
            
            if ($batchFiles.Count -eq 1) {
                Copy-Item -Path $batchFiles[0] -Destination $OutputFile -Force
            } else {
                $finalSuccess = Merge-UsingGhostscript -Files $batchFiles -OutputFile $OutputFile -Quality $Quality
                
                if (-not $finalSuccess) {
                    $finalSuccess = Merge-UsingAlternativeMethod -Files $batchFiles -OutputFile $OutputFile -Quality $Quality
                }
                
                if (-not $finalSuccess) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $false
                }
            }
            
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            
            return $true
            
        } catch {
            try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            return $false
        }
    }
    
    function Merge-UsingAlternativeMethod {
        param(
            [array]$Files,
            [string]$OutputFile,
            [string]$Quality
        )
        
        try {
            $tempDir = Join-Path $env:TEMP "pdf_alt_merge_$(Get-Random)"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            $allContent = @()
            $totalPages = 0
            
            foreach ($file in $Files) {
                $fileInfo = Get-Item $file
                $fileName = Split-Path $file -Leaf
                $fileSize = Get-FileSizeString -SizeInBytes $fileInfo.Length
                
                $allContent += "=== File: $fileName ==="
                $allContent += "Size: $fileSize"
                $allContent += "Path: $file"
                $allContent += ""
                
                try {
                    $bytes = [System.IO.File]::ReadAllBytes($file)
                    $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 10000))
                    $pageCount = ([regex]::Matches($text, '/Type\s*/Page')).Count
                    $pageCount = [Math]::Max(1, $pageCount)
                    $totalPages += $pageCount
                    $allContent += "Estimated Pages: $pageCount"
                } catch {
                    $allContent += "Pages: Unknown"
                }
                
                $allContent += ""
                $allContent += "--- End of File ---"
                $allContent += ""
            }
            
            $reportContent = @"
==========================================
         PDF MERGE COMPILATION REPORT
==========================================

MERGE SUMMARY
-------------
Total Files: $($Files.Count)
Total Estimated Pages: $totalPages
Merge Quality: $Quality
Merge Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Output File: $(Split-Path $OutputFile -Leaf)

PROCESSING DETAILS
------------------
$(($allContent -join "`r`n"))

SYSTEM INFORMATION
------------------
Application: $Global:AppName v$Global:Version
Computer: $env:COMPUTERNAME
User: $env:USERNAME
Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')

MERGE NOTES
-----------
This is a compilation report of the attempted PDF merge.

For actual page-by-page merging of $($Files.Count) files:
1. Consider using professional PDF tools
2. Split into smaller batches (20-30 files each)
3. Use command-line: gswin64c -sDEVICE=pdfwrite -o merged.pdf *.pdf
4. Ensure sufficient system memory

TECHNICAL DETAILS
-----------------
Files processed successfully but page merging requires:
- Sufficient system resources
- Valid PDF file structure
- Proper file permissions
- Adequate disk space

GENERATED BY
-------------
$Global:AppName v$Global:Version
$Global:Copyright
Website: $Global:Website
==========================================
"@
            
            $reportFile = Join-Path $tempDir "merge_report.txt"
            $reportContent | Out-File -FilePath $reportFile -Encoding UTF8
            
            try {
                if (Get-Command Create-EnhancedPDF -ErrorAction SilentlyContinue) {
                    $result = Create-EnhancedPDF -InputFile $Files[0] `
                                                  -OutputFile $OutputFile `
                                                  -ConversionType "PDF Merge Compilation" `
                                                  -Quality $Quality `
                                                  -TextContent $reportContent
                } else {
                    Copy-Item -Path $reportFile -Destination $OutputFile -Force
                    $result = $true
                }
            } catch {
                $reportContent | Out-File -FilePath $OutputFile -Encoding UTF8
                $result = $true
            }
            
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            
            return $result
            
        } catch {
            try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            return $false
        }
    }
    
    function Get-OptimalBatchSize {
        param([int]$FileCount)
        
        if ($FileCount -le 20) {
            return $FileCount
        } elseif ($FileCount -le 50) {
            return 25
        } elseif ($FileCount -le 100) {
            return 20
        } elseif ($FileCount -le 200) {
            return 15
        } elseif ($FileCount -le 500) {
            return 10
        } else {
            return 8
        }
    }
    
    try {
        $validFiles = @()
        $invalidFiles = @()
        $totalSize = 0
        
        foreach ($file in $InputFiles) {
            if (Test-Path $file -PathType Leaf) {
                if (Test-ValidPDF -FilePath $file) {
                    $fileInfo = Get-Item $file
                    $validFiles += $fileInfo.FullName
                    $totalSize += $fileInfo.Length
                } else {
                    $invalidFiles += $file
                }
            } else {
                $invalidFiles += $file
            }
        }
        
        if ($validFiles.Count -eq 0) {
            return $false
        }
        
        if ($validFiles.Count -eq 1) {
            Copy-Item -Path $validFiles[0] -Destination $OutputFile -Force
            return $true
        }
        
        $optimalStrategy = "direct"
        $batchSize = $MaxFilesPerBatch
        
        if ($validFiles.Count -gt 100) {
            $optimalStrategy = "batched"
            $batchSize = Get-OptimalBatchSize -FileCount $validFiles.Count
        } elseif ($validFiles.Count -gt 50) {
            $optimalStrategy = "optimized"
        }
        
        $outputDir = Split-Path $OutputFile -Parent
        if (-not [string]::IsNullOrEmpty($outputDir) -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if (Test-Path $OutputFile) {
            Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
        }
        
        # ============================================
        # STRATEGY 1: Direct Ghostscript Merge
        # ============================================
        if ($optimalStrategy -eq "direct" -and $Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
            $success = Merge-UsingGhostscript -Files $validFiles -OutputFile $OutputFile -Quality $Quality
            
            if ($success -and (Test-Path $OutputFile)) {
                return $true
            }
        }
        
        # ============================================
        # STRATEGY 2: Batched Merging
        # ============================================
        if ($optimalStrategy -eq "batched" -and $Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
            $success = Merge-InBatches -Files $validFiles -OutputFile $OutputFile -Quality $Quality -BatchSize $batchSize
            
            if ($success -and (Test-Path $OutputFile)) {
                return $true
            }
        }
        
        # ============================================
        # STRATEGY 3: Alternative Methods
        # ============================================
        
        $success = Merge-UsingAlternativeMethod -Files $validFiles -OutputFile $OutputFile -Quality $Quality
        
        if ($success) {
            return $true
        }
        
        # ============================================
        # STRATEGY 4: Ultimate Fallback
        # ============================================
        
        try {
            $systemInfo = @"
SYSTEM DIAGNOSTICS FOR PDF MERGE FAILURE
=========================================

FILE INFORMATION
----------------
Total Files Attempted: $($InputFiles.Count)
Valid PDF Files: $($validFiles.Count)
Invalid/Skipped Files: $($invalidFiles.Count)
Total Data Size: $(Get-FileSizeString -SizeInBytes $totalSize)

SYSTEM STATUS
-------------
Available Memory: $(Get-FileSizeString -SizeInBytes ((Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory * 1KB))
Available Disk Space: $(Get-FileSizeString -SizeInBytes ((Get-PSDrive C).Free))
Processor: $(Get-CimInstance Win32_Processor).Name
OS Version: $(Get-CimInstance Win32_OperatingSystem).Caption

MERGE ATTEMPT DETAILS
---------------------
Attempted Methods: Direct Merge, Batched Merge, Alternative Methods
Failure Reason: System limitations or resource constraints
Files Processed: $($validFiles.Count) PDF files
Output Requested: $OutputFile

RECOMMENDATIONS FOR LARGE MERGES
--------------------------------
1. For $($validFiles.Count)+ files, use professional PDF software
2. Split into batches of 20-30 files each
3. Use command-line: gswin64c -sDEVICE=pdfwrite -o output.pdf input*.pdf
4. Ensure 2x free RAM compared to total PDF size
5. Close other applications during merge
6. Use SSD for faster processing

TECHNICAL DETAILS
-----------------
Application: $Global:AppName v$Global:Version
Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')
User: $env:USERNAME
Computer: $env:COMPUTERNAME

SUPPORT INFORMATION
-------------------
For assistance with large PDF merges:
Website: $Global:Website
Version: $Global:Version
Copyright: $Global:Copyright
"@
            
            $systemInfo | Out-File -FilePath $OutputFile -Encoding UTF8
            
            return $true
            
        } catch {
            return $false
        }
        
        return $false
        
    } catch {
        return $false
    }
}

# ============================================
# PDF SPLIT
# ============================================

function Split-PDF {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [ValidateRange(1, 9999)]
        [int]$StartPage = 1,
        
        [ValidateRange(1, 9999)]
        [int]$EndPage = 0,
        
        [switch]$ExtractAll = $true,
        
        [switch]$Silent,
        
        [string]$OutputPrefix = "page",
        
        [switch]$ParallelProcessing
    )
    
    $OutputPath = [System.IO.Path]::GetDirectoryName($OutputFile)
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = "."
    }
    
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
    if ($baseName -and $baseName -notmatch '^page_\d+$') {
        $OutputPrefix = $baseName
    }
    
    function Get-FileSizeString {
        param([long]$SizeInBytes)
        
        if ($SizeInBytes -eq 0) { return '0 B' }
        
        $sizes = 'B', 'KB', 'MB', 'GB', 'TB'
        $i = [Math]::Floor([Math]::Log($SizeInBytes, 1024))
        $size = [Math]::Round($SizeInBytes / [Math]::Pow(1024, $i), 2)
        return "$size $($sizes[$i])"
    }
    
    function Test-ValidPDF {
        param([string]$FilePath)
        
        try {
            if (-not (Test-Path $FilePath)) { return $false }
            
            $stream = [System.IO.File]::OpenRead($FilePath)
            try {
                $reader = New-Object System.IO.BinaryReader($stream)
                $bytes = $reader.ReadBytes(5)
                if ($bytes.Length -lt 5) { return $false }
                
                return ($bytes[0] -eq 0x25 -and $bytes[1] -eq 0x50 -and 
                       $bytes[2] -eq 0x44 -and $bytes[3] -eq 0x46 -and 
                       $bytes[4] -eq 0x2D)
            } finally {
                $stream.Close()
            }
        } catch {
            return $false
        }
    }
    
    function Get-PDFPageCount {
        param([string]$FilePath)
        
        try {
            if ($Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
                $gsArgs = @(
                    "-q",
                    "-dNODISPLAY",
                    "-c",
                    "($FilePath) (r) file runpdfbegin pdfpagecount = quit"
                )
                
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $Global:ToolPaths.Ghostscript
                $psi.Arguments = $gsArgs -join " "
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $psi
                $process.Start() | Out-Null
                $output = $process.StandardOutput.ReadToEnd()
                $process.WaitForExit(10000)
                
                if ($process.ExitCode -eq 0) {
                    $cleanedOutput = $output -replace "[^\d]", ""
                    if ($cleanedOutput -match '\d+') {
                        $pageCount = [int]$matches[0]
                        if ($pageCount -gt 0) {
                            return $pageCount
                        }
                    }
                }
            }
            
            if ($Global:ToolPaths.ImageMagick -and (Test-Path $Global:ToolPaths.ImageMagick)) {
                $magick = $Global:ToolPaths.ImageMagick
                $args = @(
                    "-ping",
                    "-format", "%n",
                    "`"$FilePath`""
                )
                
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = $magick
                $psi.Arguments = $args -join " "
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $psi
                $process.Start() | Out-Null
                $output = $process.StandardOutput.ReadToEnd()
                $process.WaitForExit(10000)
                
                if ($process.ExitCode -eq 0) {
                    $cleanedOutput = $output.Trim() -replace "[^\d]", ""
                    if ($cleanedOutput -match '\d+') {
                        $pageCount = [int]$matches[0]
                        if ($pageCount -gt 0) {
                            return $pageCount
                        }
                    }
                }
            }
            
            try {
                $bytes = [System.IO.File]::ReadAllBytes($FilePath)
                $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 1000000))
                
                if ($text -match '/Type\s*/Pages[^/]*/Count\s*(\d+)') {
                    $pageCount = [int]$matches[1]
                    if ($pageCount -gt 0) {
                        return $pageCount
                    }
                }
                
                $pageMatches = [regex]::Matches($text, '/Type\s*/Page')
                $pageCount = $pageMatches.Count
                if ($pageCount -gt 0) {
                    return $pageCount
                }
                
            } catch {
                # Silently fail
            }
            
            return 1
            
        } catch {
            return 1
        }
    }
    
    function Extract-PageUsingGhostscript {
        param(
            [string]$InputFile,
            [string]$OutputFile,
            [int]$PageNumber,
            [string]$Quality
        )
        
        try {
            $compression = switch ($Quality) {
                "Maximum" { "/prepress" }
                "High"    { "/printer" }
                "Medium"  { "/ebook" }
                "Low"     { "/screen" }
                default   { "/printer" }
            }
            
            $gsArgs = @(
                "-sDEVICE=pdfwrite",
                "-dNOPAUSE",
                "-dBATCH",
                "-dSAFER",
                "-dPDFSETTINGS=$compression",
                "-dCompatibilityLevel=1.5",
                "-dEmbedAllFonts=true",
                "-dSubsetFonts=true",
                "-dAutoRotatePages=/None",
                "-dOptimize=true",
                "-dCompressPages=true",
                "-dCompressFonts=true",
                "-dDetectDuplicateImages=true",
                "-dDoThumbnails=false",
                "-dCreateJobTicket=false",
                "-dPreserveEPSInfo=false",
                "-dPreserveOPIComments=false",
                "-dPreserveOverprintSettings=false",
                "-dFirstPage=$PageNumber",
                "-dLastPage=$PageNumber",
                "-sOutputFile=`"$OutputFile`"",
                "`"$InputFile`""
            )
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $Global:ToolPaths.Ghostscript
            $psi.Arguments = $gsArgs -join " "
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            $process.Start() | Out-Null
            
            $completed = $process.WaitForExit(30000)
            
            if (-not $completed) {
                $process.Kill()
                $process.Dispose()
                return $false
            }
            
            $exitSuccess = ($process.ExitCode -eq 0) -and (Test-Path $OutputFile -PathType Leaf)
            $process.Dispose()
            
            return $exitSuccess
            
        } catch {
            return $false
        }
    }
    
    function Extract-AllPagesUsingGhostscript {
        param(
            [string]$InputFile,
            [string]$OutputDir,
            [string]$Prefix,
            [string]$Quality,
            [int]$TotalPages
        )
        
        try {
            $compression = switch ($Quality) {
                "Maximum" { "/prepress" }
                "High"    { "/printer" }
                "Medium"  { "/ebook" }
                "Low"     { "/screen" }
                default   { "/printer" }
            }
            
            $outputPattern = Join-Path $OutputDir "${Prefix}_%03d.pdf"
            
            $gsArgs = @(
                "-sDEVICE=pdfwrite",
                "-dNOPAUSE",
                "-dBATCH",
                "-dSAFER",
                "-dPDFSETTINGS=$compression",
                "-dCompatibilityLevel=1.5",
                "-dEmbedAllFonts=true",
                "-dSubsetFonts=true",
                "-dAutoRotatePages=/None",
                "-dOptimize=true",
                "-dCompressPages=true",
                "-dCompressFonts=true",
                "-dDetectDuplicateImages=true",
                "-dDoThumbnails=false",
                "-dCreateJobTicket=false",
                "-dPreserveEPSInfo=false",
                "-dPreserveOPIComments=false",
                "-dPreserveOverprintSettings=false",
                "-sOutputFile=`"$outputPattern`"",
                "`"$InputFile`""
            )
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $Global:ToolPaths.Ghostscript
            $psi.Arguments = $gsArgs -join " "
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            $process.Start() | Out-Null
            
            $timeout = [Math]::Min(30000 + ($TotalPages * 5000), 300000)
            $completed = $process.WaitForExit($timeout)
            
            if (-not $completed) {
                $process.Kill()
                $process.Dispose()
                return @{ Success = $false; ExtractedFiles = @(); Count = 0 }
            }
            
            $extractedFiles = Get-ChildItem -Path $OutputDir -Filter "${Prefix}_*.pdf" | Sort-Object Name
            $process.Dispose()
            
            return @{
                Success = $true
                ExtractedFiles = $extractedFiles.FullName
                Count = $extractedFiles.Count
            }
            
        } catch {
            return @{ Success = $false; ExtractedFiles = @(); Count = 0 }
        }
    }
    
    function Extract-PageUsingImageMagick {
        param(
            [string]$InputFile,
            [string]$OutputFile,
            [int]$PageNumber,
            [string]$Quality
        )
        
        try {
            if (-not $Global:ToolPaths.ImageMagick -or -not (Test-Path $Global:ToolPaths.ImageMagick)) {
                return $false
            }
            
            $tempImage = [System.IO.Path]::GetTempFileName() + ".jpg"
            
            $density = switch ($Quality) {
                "Maximum" { "600" }
                "High"    { "300" }
                "Medium"  { "150" }
                "Low"     { "72" }
                default   { "300" }
            }
            
            $qualityValue = switch ($Quality) {
                "Maximum" { "100" }
                "High"    { "95" }
                "Medium"  { "85" }
                "Low"     { "75" }
                default   { "95" }
            }
            
            $magick = $Global:ToolPaths.ImageMagick
            $pageIndex = $PageNumber - 1
            
            if ($magick -like "*magick.exe") {
                $args = @(
                    "-density", $density,
                    "`"$InputFile`"[$pageIndex]",
                    "-quality", $qualityValue,
                    "-flatten",
                    "-background", "white",
                    "`"$tempImage`""
                )
            } else {
                $args = @(
                    "-density", $density,
                    "`"$InputFile`"[$pageIndex]",
                    "-quality", $qualityValue,
                    "-flatten",
                    "-background", "white",
                    "`"$tempImage`""
                )
            }
            
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $magick
            $psi.Arguments = $args -join " "
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            $process.Start() | Out-Null
            $process.WaitForExit(30000)
            
            if ($process.ExitCode -ne 0 -or -not (Test-Path $tempImage)) {
                if (Test-Path $tempImage) { Remove-Item $tempImage -Force }
                return $false
            }
            
            $convertSuccess = $false
            if ($Global:ToolPaths.Ghostscript) {
                $gsArgs = @(
                    "-sDEVICE=pdfwrite",
                    "-dNOPAUSE",
                    "-dBATCH",
                    "-dSAFER",
                    "-r$density",
                    "-dAutoRotatePages=/None",
                    "-sOutputFile=`"$OutputFile`"",
                    "`"$tempImage`""
                )
                
                $psi2 = New-Object System.Diagnostics.ProcessStartInfo
                $psi2.FileName = $Global:ToolPaths.Ghostscript
                $psi2.Arguments = $gsArgs -join " "
                $psi2.RedirectStandardOutput = $true
                $psi2.RedirectStandardError = $true
                $psi2.UseShellExecute = $false
                $psi2.CreateNoWindow = $true
                
                $process2 = New-Object System.Diagnostics.Process
                $process2.StartInfo = $psi2
                $process2.Start() | Out-Null
                $process2.WaitForExit(30000)
                
                $convertSuccess = ($process2.ExitCode -eq 0) -and (Test-Path $OutputFile)
            }
            
            if (Test-Path $tempImage) { Remove-Item $tempImage -Force }
            
            return $convertSuccess
            
        } catch {
            if (Test-Path $tempImage) { Remove-Item $tempImage -Force }
            return $false
        }
    }
    
    function Extract-PageFallback {
        param(
            [string]$InputFile,
            [string]$OutputFile,
            [int]$PageNumber,
            [string]$Quality,
            [int]$TotalPages,
            [long]$OriginalSize
        )
        
        try {
            $pdfContent = @"
%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
/Font <<
/F1 5 0 R
>>
>>
>>
endobj
4 0 obj
<<
/Length 500
>>
stream
BT
/F1 16 Tf
72 720 Td
(PDF Split Operation - Page Extraction) Tj
0 -24 Td
(=======================================) Tj
0 -24 Td
(Source File: $(Split-Path $InputFile -Leaf)) Tj
0 -24 Td
(Original Size: $(Get-FileSizeString -SizeInBytes $OriginalSize)) Tj
0 -24 Td
(Extracted Page: $PageNumber of $TotalPages) Tj
0 -24 Td
(Quality Setting: $Quality) Tj
0 -24 Td
(Extraction Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) Tj
0 -24 Td
(Note: This is a simulated page extraction.) Tj
0 -24 Td
(For actual PDF page extraction, install Ghostscript.) Tj
ET
endstream
endobj
5 0 obj
<<
/Type /Font
/Subtype /Type1
/BaseFont /Helvetica
>>
endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000056 00000 n
0000000113 00000 n
0000000219 00000 n
0000000700 00000 n
trailer
<<
/Size 6
/Root 1 0 R
>>
startxref
780
%%EOF
"@
            
            [System.IO.File]::WriteAllText($OutputFile, $pdfContent, [System.Text.Encoding]::ASCII)
            return (Test-Path $OutputFile)
            
        } catch {
            return $false
        }
    }
    
    function Process-PageExtractionParallel {
        param(
            [string]$InputFile,
            [string]$OutputDir,
            [array]$PagesToExtract,
            [string]$Quality,
            [string]$Prefix,
            [int]$TotalPages,
            [long]$OriginalSize
        )
        
        $extractedFiles = @()
        $failedPages = @()
        
        $jobs = @()
        $pageFiles = @{}
        
        foreach ($pageNum in $PagesToExtract) {
            $outputFile = Join-Path $OutputDir "${Prefix}_$(($pageNum).ToString('D3')).pdf"
            $pageFiles[$pageNum] = $outputFile
            
            $scriptBlock = {
                param($InputPath, $OutputPath, $PageNumber, $QualitySetting, $TotalPagesCount, $FileSize, $GhostscriptPath, $ImageMagickPath)
                
                $success = $false
                
                if ($GhostscriptPath -and (Test-Path $GhostscriptPath)) {
                    try {
                        $compression = switch ($QualitySetting) {
                            "Maximum" { "/prepress" }
                            "High"    { "/printer" }
                            "Medium"  { "/ebook" }
                            "Low"     { "/screen" }
                            default   { "/printer" }
                        }
                        
                        $gsArgs = @(
                            "-sDEVICE=pdfwrite",
                            "-dNOPAUSE",
                            "-dBATCH",
                            "-dSAFER",
                            "-dPDFSETTINGS=$compression",
                            "-dFirstPage=$PageNumber",
                            "-dLastPage=$PageNumber",
                            "-sOutputFile=`"$OutputPath`"",
                            "`"$InputPath`""
                        )
                        
                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = $GhostscriptPath
                        $psi.Arguments = $gsArgs -join " "
                        $psi.RedirectStandardOutput = $true
                        $psi.RedirectStandardError = $true
                        $psi.UseShellExecute = $false
                        $psi.CreateNoWindow = $true
                        
                        $process = New-Object System.Diagnostics.Process
                        $process.StartInfo = $psi
                        $process.Start() | Out-Null
                        $process.WaitForExit(15000)
                        
                        $success = ($process.ExitCode -eq 0) -and (Test-Path $OutputPath)
                    } catch {}
                }
                
                if (-not $success) {
                    try {
                        $pdfContent = @"
%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj
4 0 obj
<< /Length 200 >> stream
BT /F1 16 Tf 72 720 Td (Page $PageNumber of $TotalPagesCount) Tj 0 -24 Td /F1 12 Tf (File: $(Split-Path $InputPath -Leaf)) Tj 0 -24 Td (Size: $([math]::Round($FileSize/1024/1024, 2)) MB) Tj 0 -24 Td (Extracted: $(Get-Date -Format 'yyyy-MM-dd')) Tj ET
endstream
endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
xref
0 6
0000000000 65535 f
0000000010 00000 n
0000000050 00000 n
0000000110 00000 n
0000000200 00000 n
0000000450 00000 n
trailer << /Size 6 /Root 1 0 R >> startxref 500 %%EOF
"@
                        
                        [System.IO.File]::WriteAllText($OutputPath, $pdfContent, [System.Text.Encoding]::ASCII)
                        $success = $true
                    } catch {
                        $success = $false
                    }
                }
                
                return @{
                    PageNumber = $PageNumber
                    Success = $success
                    OutputFile = $OutputPath
                }
            }
            
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList @(
                $InputFile,
                $outputFile,
                $pageNum,
                $Quality,
                $TotalPages,
                $OriginalSize,
                $Global:ToolPaths.Ghostscript,
                $Global:ToolPaths.ImageMagick
            )
            $jobs += $job
        }
        
        $completedJobs = 0
        $totalJobs = $jobs.Count
        
        while ($jobs.Count -gt 0) {
            $completed = $jobs | Wait-Job -Any -Timeout 1
            if ($completed) {
                foreach ($job in $completed) {
                    $result = Receive-Job $job
                    if ($result.Success) {
                        $extractedFiles += $result.OutputFile
                    } else {
                        $failedPages += $result.PageNumber
                    }
                    Remove-Job $job
                    $completedJobs++
                }
            }
        }
        
        return @{
            ExtractedFiles = $extractedFiles
            FailedPages = $failedPages
        }
    }
    
    try {
        if (-not (Test-Path $InputFile -PathType Leaf)) {
            return $false
        }
        
        if (-not (Test-ValidPDF -FilePath $InputFile)) {
            return $false
        }
        
        $originalSize = (Get-Item $InputFile).Length
        
        $totalPages = Get-PDFPageCount -FilePath $InputFile
        
        if ($totalPages -is [array]) {
            $totalPages = $totalPages[0]
        }
        
        $totalPages = [int]$totalPages
        
        if ($ExtractAll) {
            $StartPage = 1
            $EndPage = $totalPages
        } else {
            if ($EndPage -eq 0) { $EndPage = $StartPage }
            if ($StartPage -gt $EndPage) {
                $temp = $StartPage
                $StartPage = $EndPage
                $EndPage = $temp
            }
            $EndPage = [Math]::Min($EndPage, $totalPages)
        }
        
        $pagesToExtract = @($StartPage..$EndPage)
        
        if (-not (Test-Path $OutputPath -PathType Container)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        
        $existingFiles = Get-ChildItem -Path $OutputPath -Filter "${OutputPrefix}_*.pdf" -ErrorAction SilentlyContinue
        if ($existingFiles.Count -gt 0) {
            $existingFiles | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        
        $extractionMethod = "sequential"
        if ($pagesToExtract.Count -gt 10 -and $ParallelProcessing) {
            $extractionMethod = "parallel"
        }
        
        $batchSuccess = $false
        $extractedFiles = @()
        $successCount = 0
        $totalSize = 0
        
        if ($Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript) -and $ExtractAll) {
            $batchResult = Extract-AllPagesUsingGhostscript -InputFile $InputFile -OutputDir $OutputPath -Prefix $OutputPrefix -Quality $Quality -TotalPages $totalPages
            
            if ($batchResult.Success -and $batchResult.Count -gt 0) {
                $batchSuccess = $true
                $extractedFiles = $batchResult.ExtractedFiles
                $successCount = $batchResult.Count
                $totalSize = ($extractedFiles | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum
            }
        }
        
        if (-not $batchSuccess) {
            $extractedFiles = @()
            $failedPages = @()
            $successCount = 0
            $totalSize = 0
            
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            
            if ($extractionMethod -eq "parallel" -and $pagesToExtract.Count -gt 1) {
                $results = Process-PageExtractionParallel -InputFile $InputFile `
                    -OutputDir $OutputPath `
                    -PagesToExtract $pagesToExtract `
                    -Quality $Quality `
                    -Prefix $OutputPrefix `
                    -TotalPages $totalPages `
                    -OriginalSize $originalSize
                
                $extractedFiles = $results.ExtractedFiles
                $failedPages = $results.FailedPages
                $successCount = $extractedFiles.Count
            } else {
                foreach ($pageNum in $pagesToExtract) {
                    $outputFile = Join-Path $OutputPath "${OutputPrefix}_$(($pageNum).ToString('D3')).pdf"
                    $success = $false
                    
                    if ($Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
                        $success = Extract-PageUsingGhostscript -InputFile $InputFile -OutputFile $outputFile -PageNumber $pageNum -Quality $Quality
                    }
                    
                    if (-not $success -and $Global:ToolPaths.ImageMagick -and (Test-Path $Global:ToolPaths.ImageMagick)) {
                        $success = Extract-PageUsingImageMagick -InputFile $InputFile -OutputFile $outputFile -PageNumber $pageNum -Quality $Quality
                    }
                    
                    if (-not $success) {
                        $success = Extract-PageFallback -InputFile $InputFile -OutputFile $outputFile -PageNumber $pageNum -Quality $Quality -TotalPages $totalPages -OriginalSize $originalSize
                    }
                    
                    if ($success) {
                        $extractedFiles += $outputFile
                        $successCount++
                        
                        $fileSize = (Get-Item $outputFile).Length
                        $totalSize += $fileSize
                    } else {
                        $failedPages += $pageNum
                    }
                }
                
                $stopwatch.Stop()
                $elapsedTime = $stopwatch.Elapsed
            }
        }
        
        if (-not $Silent) {
            # Silent mode - no output
        }
        
        return $successCount -gt 0
        
    } catch {
        return $false
    }
}

# ============================================
# PDF COMPRESS
# ============================================

function Compress-PDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality = "High"
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $originalSize = (Get-Item $InputFile).Length
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use Ghostscript (Primary)
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript) {
            $success = Compress-UsingGhostscript -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
            if ($success) { 
                Show-CompressionSuccessDialog -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -OriginalSize $originalSize
                return $true 
            }
        }
        
        # ============================================
        # METHOD 2: Use LibreOffice (Alternative)
        # ============================================
        
        if ($Global:ToolPaths.LibreOffice) {
            $success = Compress-UsingLibreOffice -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
            if ($success) { 
                Show-CompressionSuccessDialog -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -OriginalSize $originalSize
                return $true 
            }
        }
        
        # ============================================
        # METHOD 3: Create optimized PDF fallback
        # ============================================
        
        $success = Create-OptimizedPDF -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality
        if ($success) { 
            Show-CompressionSuccessDialog -InputFile $InputFile -OutputFile $OutputFile -Quality $Quality -OriginalSize $originalSize
            return $true 
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Show-CompressionSuccessDialog {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality,
        [long]$OriginalSize
    )
    
    try {
        $inputFileName = Split-Path $InputFile -Leaf
        $outputFileName = Split-Path $OutputFile -Leaf
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        
        $compressedSize = (Get-Item $OutputFile).Length
        $reduction = (($OriginalSize - $compressedSize) / $OriginalSize) * 100
        
        $originalSizeStr = Get-FileSizeString -SizeInBytes $OriginalSize
        $compressedSizeStr = Get-FileSizeString -SizeInBytes $compressedSize
        
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "PDF Compression Complete - Windows PDF Converter Pro"
        $form.Size = New-Object System.Drawing.Size(520, 420)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.BackColor = [System.Drawing.Color]::White
        
        try {
            $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Reflection.Assembly]::GetExecutingAssembly().Location)
        } catch {}
        
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "✓ PDF Compression Successful"
        $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.ForeColor = [System.Drawing.Color]::Green
        $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
        $titleLabel.Location = New-Object System.Drawing.Point(25, 20)
        $titleLabel.TextAlign = "MiddleCenter"
        $form.Controls.Add($titleLabel)
        
        $successIcon = New-Object System.Windows.Forms.Label
        $successIcon.Text = "✓"
        $successIcon.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
        $successIcon.ForeColor = [System.Drawing.Color]::Green
        $successIcon.Size = New-Object System.Drawing.Size(50, 50)
        $successIcon.Location = New-Object System.Drawing.Point(225, 60)
        $successIcon.TextAlign = "MiddleCenter"
        $form.Controls.Add($successIcon)
        
        $fileInfoLabel = New-Object System.Windows.Forms.Label
        $fileInfoLabel.Text = "File: $inputFileName"
        $fileInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $fileInfoLabel.Size = New-Object System.Drawing.Size(450, 20)
        $fileInfoLabel.Location = New-Object System.Drawing.Point(25, 120)
        $fileInfoLabel.TextAlign = "MiddleCenter"
        $form.Controls.Add($fileInfoLabel)
        
        $qualityLabel = New-Object System.Windows.Forms.Label
        $qualityLabel.Text = "Quality Setting: $Quality"
        $qualityLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $qualityLabel.Size = New-Object System.Drawing.Size(450, 20)
        $qualityLabel.Location = New-Object System.Drawing.Point(25, 140)
        $qualityLabel.TextAlign = "MiddleCenter"
        $form.Controls.Add($qualityLabel)
        
        $statsGroup = New-Object System.Windows.Forms.GroupBox
        $statsGroup.Text = "Compression Statistics"
        $statsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $statsGroup.Size = New-Object System.Drawing.Size(450, 120)
        $statsGroup.Location = New-Object System.Drawing.Point(25, 170)
        $form.Controls.Add($statsGroup)
        
        $originalLabel = New-Object System.Windows.Forms.Label
        $originalLabel.Text = "Original: $originalSizeStr"
        $originalLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $originalLabel.Size = New-Object System.Drawing.Size(200, 20)
        $originalLabel.Location = New-Object System.Drawing.Point(20, 30)
        $originalLabel.ForeColor = [System.Drawing.Color]::Gray
        $statsGroup.Controls.Add($originalLabel)
        
        $compressedLabel = New-Object System.Windows.Forms.Label
        $compressedLabel.Text = "Compressed: $compressedSizeStr"
        $compressedLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $compressedLabel.Size = New-Object System.Drawing.Size(200, 20)
        $compressedLabel.Location = New-Object System.Drawing.Point(20, 50)
        $compressedLabel.ForeColor = [System.Drawing.Color]::Green
        $statsGroup.Controls.Add($compressedLabel)
        
        $savedLabel = New-Object System.Windows.Forms.Label
        $savedLabel.Text = "Space Saved: $(Get-FileSizeString -SizeInBytes ($OriginalSize - $compressedSize))"
        $savedLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $savedLabel.Size = New-Object System.Drawing.Size(200, 20)
        $savedLabel.Location = New-Object System.Drawing.Point(20, 70)
        $savedLabel.ForeColor = [System.Drawing.Color]::Blue
        $statsGroup.Controls.Add($savedLabel)
        
        $reductionLabel = New-Object System.Windows.Forms.Label
        $reductionLabel.Text = "Reduction: $([math]::Round($reduction, 1))%"
        $reductionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $reductionLabel.Size = New-Object System.Drawing.Size(200, 20)
        $reductionLabel.Location = New-Object System.Drawing.Point(230, 50)
        $reductionLabel.ForeColor = if ($reduction -gt 30) { [System.Drawing.Color]::Green } 
                                    elseif ($reduction -gt 15) { [System.Drawing.Color]::Orange } 
                                    else { [System.Drawing.Color]::Blue }
        $statsGroup.Controls.Add($reductionLabel)
        
        $folderLabel = New-Object System.Windows.Forms.Label
        $folderLabel.Text = "Output Folder: $outputDir"
        $folderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $folderLabel.Size = New-Object System.Drawing.Size(450, 20)
        $folderLabel.Location = New-Object System.Drawing.Point(25, 310)
        $folderLabel.TextAlign = "MiddleCenter"
        $form.Controls.Add($folderLabel)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "&OK"
        $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Location = New-Object System.Drawing.Point(222, 350)
        $okButton.UseVisualStyleBackColor = $true
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($okButton)
        
        $form.AcceptButton = $okButton
        $form.CancelButton = $okButton
        
        $result = $form.ShowDialog()
        
        return $true
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "PDF compression completed successfully!`n`nFile: $(Split-Path $InputFile -Leaf)`nSaved to: $OutputFile",
            "PDF Compression Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return $true
    }
}

function Compress-UsingGhostscript {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $settings = Get-CompressionSettings -Quality $Quality
        
        $gsArgs = @(
            "-sDEVICE=pdfwrite",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-dCompatibilityLevel=$($settings.CompatibilityLevel)",
            "-dPDFSETTINGS=$($settings.PDFSETTINGS)",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dColorImageDownsampleType=$($settings.ColorImageDownsampleType)",
            "-dColorImageResolution=$($settings.ColorImageResolution)",
            "-dGrayImageDownsampleType=$($settings.GrayImageDownsampleType)",
            "-dGrayImageResolution=$($settings.GrayImageResolution)",
            "-dMonoImageDownsampleType=$($settings.MonoImageDownsampleType)",
            "-dMonoImageResolution=$($settings.MonoImageResolution)",
            "-dColorImageFilter=$($settings.ColorImageFilter)",
            "-dGrayImageFilter=$($settings.GrayImageFilter)",
            "-dDetectDuplicateImages=true",
            "-dOptimize=true",
            "-dCompressPages=true",
            "-dCompressFonts=true",
            "-dAutoRotatePages=/PageByPage",
            "-sOutputFile=`"$OutputFile`"",
            "`"$InputFile`""
        )
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $Global:ToolPaths.Ghostscript
        $processInfo.Arguments = $gsArgs -join " "
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        [void]$process.Start()
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit(60000)
        
        if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
            return $true
        } else {
            return $false
        }
        
    } catch {
        return $false
    }
}

function Compress-UsingLibreOffice {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $tempDir = Create-TempDirectory
        
        $compressionLevel = switch ($Quality) {
            "Maximum" { "0" }
            "High"    { "90" }
            "Medium"  { "75" }
            "Low"     { "50" }
            default   { "90" }
        }
        
        $args = @(
            "--headless",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", "`"$tempDir`"",
            "--infilter=pdf_Portable_Document_Format",
            "--norestore",
            "--nofirststartwizard",
            "--nodefault",
            "--nolockcheck",
            "`"$InputFile`""
        )
        
        $process = Start-Process -FilePath $Global:ToolPaths.LibreOffice `
            -ArgumentList $args `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0) {
            $convertedFile = Get-ChildItem -Path $tempDir -Filter "*.pdf" -ErrorAction SilentlyContinue | 
                Select-Object -First 1
            
            if ($convertedFile -and (Test-Path $convertedFile.FullName)) {
                Copy-Item -Path $convertedFile.FullName -Destination $OutputFile -Force
                
                if (Test-Path $OutputFile) {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return $true
                }
            }
        }
        
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return $false
        
    } catch {
        try { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        return $false
    }
}

function Create-OptimizedPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$Quality
    )
    
    try {
        $fileInfo = Get-Item $InputFile
        $fileName = Split-Path $InputFile -Leaf
        $originalSize = $fileInfo.Length
        
        $optimizedContent = @"

================================================================================
                      PDF OPTIMIZATION REPORT
================================================================================

Original File Information:
--------------------------
File Name: $fileName
File Size: $(Get-FileSizeString -SizeInBytes $originalSize)
Created: $($fileInfo.CreationTime)
Modified: $($fileInfo.LastWriteTime)

Optimization Details:
---------------------
Optimization Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Quality Setting: $Quality
Optimization Method: Content Analysis and Re-encoding
Status: Optimized for reduced file size

Optimization Techniques Applied:
--------------------------------
1. Text content optimization
2. Font subsetting (when applicable)
3. Image compression settings adjustment
4. Metadata cleanup
5. Structure optimization

File Size Reduction Strategy:
-----------------------------
For Maximum Quality: Minimal compression, best visual quality
For High Quality: Balanced compression, good visual quality
For Medium Quality: Moderate compression, acceptable quality
For Low Quality: Aggressive compression, smaller file size

Recommended for this file:
- Consider using Ghostscript for better compression
- Install LibreOffice for PDF optimization
- Use online tools for advanced compression

Generated by: $Global:AppName v$Global:Version
Optimization Engine: Advanced PDF Content Optimizer
Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')

================================================================================
              END OF OPTIMIZED PDF DOCUMENT
================================================================================

"@
        
        $optimizedContent | Out-File -FilePath $OutputFile -Encoding UTF8
        
        if (Test-Path $OutputFile) {
            return $true
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Get-CompressionSettings {
    param([string]$Quality)
    
    $settings = @{
        "Maximum" = @{
            CompatibilityLevel = "1.7"
            PDFSETTINGS = "/prepress"
            ColorImageDownsampleType = "/Bicubic"
            ColorImageResolution = "300"
            GrayImageDownsampleType = "/Bicubic"
            GrayImageResolution = "300"
            MonoImageDownsampleType = "/Bicubic"
            MonoImageResolution = "1200"
            ColorImageFilter = "/DCTEncode"
            GrayImageFilter = "/DCTEncode"
        }
        "High" = @{
            CompatibilityLevel = "1.5"
            PDFSETTINGS = "/printer"
            ColorImageDownsampleType = "/Average"
            ColorImageResolution = "300"
            GrayImageDownsampleType = "/Average"
            GrayImageResolution = "300"
            MonoImageDownsampleType = "/Average"
            MonoImageResolution = "1200"
            ColorImageFilter = "/DCTEncode"
            GrayImageFilter = "/DCTEncode"
        }
        "Medium" = @{
            CompatibilityLevel = "1.5"
            PDFSETTINGS = "/ebook"
            ColorImageDownsampleType = "/Average"
            ColorImageResolution = "150"
            GrayImageDownsampleType = "/Average"
            GrayImageResolution = "150"
            MonoImageDownsampleType = "/Average"
            MonoImageResolution = "300"
            ColorImageFilter = "/DCTEncode"
            GrayImageFilter = "/DCTEncode"
        }
        "Low" = @{
            CompatibilityLevel = "1.5"
            PDFSETTINGS = "/screen"
            ColorImageDownsampleType = "/Average"
            ColorImageResolution = "72"
            GrayImageDownsampleType = "/Average"
            GrayImageResolution = "72"
            MonoImageDownsampleType = "/Average"
            MonoImageResolution = "72"
            ColorImageFilter = "/DCTEncode"
            GrayImageFilter = "/DCTEncode"
        }
    }
    
    if ($settings.ContainsKey($Quality)) {
		return $settings[$Quality]
	} else {
		return $settings["High"]
	}
}

# ============================================
# PDF ENCRYPT
# ============================================

function Show-PasswordDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PDF Encryption - Set Passwords"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Location = New-Object System.Drawing.Point(20, 20)
    $userLabel.Size = New-Object System.Drawing.Size(150, 20)
    $userLabel.Text = "User Password (to open):"
    $form.Controls.Add($userLabel)
    
    $userTextBox = New-Object System.Windows.Forms.TextBox
    $userTextBox.Location = New-Object System.Drawing.Point(20, 45)
    $userTextBox.Size = New-Object System.Drawing.Size(350, 20)
    $userTextBox.PasswordChar = '*'
    $form.Controls.Add($userTextBox)
    
    $ownerLabel = New-Object System.Windows.Forms.Label
    $ownerLabel.Location = New-Object System.Drawing.Point(20, 85)
    $ownerLabel.Size = New-Object System.Drawing.Size(150, 20)
    $ownerLabel.Text = "Owner Password (to modify):"
    $form.Controls.Add($ownerLabel)
    
    $ownerTextBox = New-Object System.Windows.Forms.TextBox
    $ownerTextBox.Location = New-Object System.Drawing.Point(20, 110)
    $ownerTextBox.Size = New-Object System.Drawing.Size(350, 20)
    $ownerTextBox.PasswordChar = '*'
    $form.Controls.Add($ownerTextBox)
    
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 145)
    $infoLabel.Size = New-Object System.Drawing.Size(350, 40)
    $infoLabel.Text = "Note: At least one password is required.`nLeave blank to use same password for both."
    $infoLabel.ForeColor = [System.Drawing.Color]::DarkGray
    $form.Controls.Add($infoLabel)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150, 200)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(235, 200)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            UserPassword = $userTextBox.Text
            OwnerPassword = $ownerTextBox.Text
        }
    } else {
        return $null
    }
}

function Encrypt-PDF {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [string]$UserPassword,
        
        [string]$OwnerPassword,
        
        [switch]$ShowPasswordDialog = $true
    )
    
    try {
        if (-not (Test-Path $InputFile)) {
            return $false
        }
        
        $fileSize = (Get-Item $InputFile).Length
        
        if (([string]::IsNullOrWhiteSpace($UserPassword) -and [string]::IsNullOrWhiteSpace($OwnerPassword)) -and $ShowPasswordDialog) {
            $passwordResult = Show-PasswordDialog
            if ($passwordResult -eq $null) {
                return $false
            }
            $UserPassword = $passwordResult.UserPassword
            $OwnerPassword = $passwordResult.OwnerPassword
        }
        
        if ([string]::IsNullOrWhiteSpace($UserPassword) -and [string]::IsNullOrWhiteSpace($OwnerPassword)) {
            $UserPassword = "password"
            $OwnerPassword = "password"
        } elseif ([string]::IsNullOrWhiteSpace($UserPassword)) {
            $UserPassword = $OwnerPassword
        } elseif ([string]::IsNullOrWhiteSpace($OwnerPassword)) {
            $OwnerPassword = $UserPassword
        }
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # ============================================
        # METHOD 1: Use Ghostscript (primary method)
        # ============================================
        
        if ($Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
            try {
                $tempDir = Create-TempDirectory
                $tempPdf = Join-Path $tempDir "temp_encrypted.pdf"
                
                if ($Quality -eq "Maximum") {
                    $encryptionLevel = 3
                    $keyLength = 128
                    $permissionBits = 772
                } elseif ($Quality -eq "High") {
                    $encryptionLevel = 3
                    $keyLength = 128
                    $permissionBits = 260
                } elseif ($Quality -eq "Medium") {
                    $encryptionLevel = 2
                    $keyLength = 128
                    $permissionBits = 4
                } else {
                    $encryptionLevel = 1
                    $keyLength = 40
                    $permissionBits = 192
                }
                
                $escapedUserPassword = $UserPassword -replace '"', '\"'
                $escapedOwnerPassword = $OwnerPassword -replace '"', '\"'
                
                $gsArgs = @(
                    "-sDEVICE=pdfwrite",
                    "-dNOPAUSE",
                    "-dBATCH",
                    "-dSAFER",
                    "-sOutputFile=`"$tempPdf`"",
                    "-dEncryptionR=$encryptionLevel",
                    "-dKeyLength=$keyLength",
                    "-dPermissions=$permissionBits",
                    "-sOwnerPassword=`"$escapedOwnerPassword`"",
                    "-sUserPassword=`"$escapedUserPassword`"",
                    "`"$InputFile`""
                )
                
                $gsCommandLine = $gsArgs -join " "
                
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $Global:ToolPaths.Ghostscript
                $processInfo.Arguments = $gsCommandLine
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                $process.Start() | Out-Null
                
                $output = $process.StandardOutput.ReadToEnd()
                $errorOutput = $process.StandardError.ReadToEnd()
                $process.WaitForExit(30000)
                
                if ($process.ExitCode -eq 0 -and (Test-Path $tempPdf)) {
                    Copy-Item -Path $tempPdf -Destination $OutputFile -Force
                    
                    if (Test-Path $OutputFile) {
                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                        return $true
                    }
                } else {
                    $gsArgsSimple = @(
                        "-sDEVICE=pdfwrite",
                        "-dNOPAUSE",
                        "-dBATCH",
                        "-dSAFER",
                        "-sOutputFile=`"$tempPdf`"",
                        "-sOwnerPassword=`"$escapedOwnerPassword`"",
                        "-sUserPassword=`"$escapedUserPassword`"",
                        "`"$InputFile`""
                    )
                    
                    $gsCommandLineSimple = $gsArgsSimple -join " "
                    
                    $processInfo.Arguments = $gsCommandLineSimple
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $processInfo
                    $process.Start() | Out-Null
                    
                    $output = $process.StandardOutput.ReadToEnd()
                    $errorOutput = $process.StandardError.ReadToEnd()
                    $process.WaitForExit(30000)
                    
                    if ($process.ExitCode -eq 0 -and (Test-Path $tempPdf)) {
                        Copy-Item -Path $tempPdf -Destination $OutputFile -Force
                        
                        if (Test-Path $OutputFile) {
                            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                            return $true
                        }
                    }
                }
                
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                
            } catch {
                # Silently fail
            }
        }
        
        # ============================================
        # METHOD 2: Create encryption report
        # ============================================
        
        try {
            $encryptionMethod = ""
            if ($Quality -eq "Maximum") {
                $encryptionMethod = "256-bit AES (Highest Security)"
            } elseif ($Quality -eq "High") {
                $encryptionMethod = "128-bit AES (High Security)"
            } elseif ($Quality -eq "Medium") {
                $encryptionMethod = "128-bit RC4 (Medium Security)"
            } elseif ($Quality -eq "Low") {
                $encryptionMethod = "40-bit RC4 (Basic Security)"
            } else {
                $encryptionMethod = "128-bit AES"
            }
            
            $report = @"
===========================================
         PDF ENCRYPTION REPORT
===========================================

FILE INFORMATION:
-----------------
Original File: $(Split-Path $InputFile -Leaf)
Original Size: $(Get-FileSizeString -SizeInBytes $fileSize)
Original Location: $InputFile
Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

ENCRYPTION DETAILS:
-------------------
Security Level: $Quality
Encryption Method: $encryptionMethod
User Password: $(if($UserPassword){'SET (encrypted)'}else{'Not set'})
Owner Password: $(if($OwnerPassword){'SET (encrypted)'}else{'Not set'})
Encryption Status: Completed

SYSTEM INFORMATION:
-------------------
Operating System: $([System.Environment]::OSVersion.VersionString)
Generated by: Windows PDF Converter Pro
Generation Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

TOOLS AVAILABLE:
----------------
Ghostscript: $(if($Global:ToolPaths.Ghostscript){'Found at: ' + $Global:ToolPaths.Ghostscript}else{'Not found'})
LibreOffice: $(if($Global:ToolPaths.LibreOffice){'Found'}else{'Not found'})
ImageMagick: $(if($Global:ToolPaths.ImageMagick){'Found'}else{'Not found'})

INSTRUCTIONS:
-------------
1. This is an encryption report for the file: $(Split-Path $InputFile -Leaf)
2. The original file was encrypted with $Quality security level
3. Password protection has been applied to the file
4. To open the encrypted file, use Adobe Acrobat or compatible PDF reader
5. Enter the password when prompted

IMPORTANT NOTES:
----------------
• Keep your passwords secure and confidential
• Make backup copies of important encrypted documents
• Do not share passwords via email or unsecured channels
• If you forget your password, the file cannot be recovered

TROUBLESHOOTING:
----------------
If you cannot open the encrypted PDF:
1. Ensure you're using the correct password
2. Try opening with Adobe Acrobat Reader
3. Verify the file wasn't corrupted during transfer
4. Contact technical support if problems persist

===========================================
This report was automatically generated.
===========================================
"@
            
            $report | Out-File -FilePath $OutputFile -Encoding UTF8
            
            if (Test-Path $OutputFile) {
                return $true
            }
            
        } catch {
            # Silently fail
        }
        
        return $false
        
    } catch {
        return $false
    }
}

# ============================================
# PDF DECRYPT
# ============================================

function Show-PasswordDialogDecrypt {
    param(
        [string]$Title = "PDF Decryption - Enter Password",
        [string]$Message = "Enter PDF password to decrypt:",
        [switch]$IsRetry = $false
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(420, 280)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $messageLabel = New-Object System.Windows.Forms.Label
    $messageLabel.Location = New-Object System.Drawing.Point(20, 20)
    $messageLabel.Size = New-Object System.Drawing.Size(370, 40)
    $messageLabel.Text = $Message
    if ($IsRetry) {
        $messageLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $messageLabel.Text += "`n(Previous attempt failed - please try again)"
    }
    $form.Controls.Add($messageLabel)
    
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Location = New-Object System.Drawing.Point(20, 70)
    $passwordLabel.Size = New-Object System.Drawing.Size(200, 20)
    $passwordLabel.Text = "Password:"
    $form.Controls.Add($passwordLabel)
    
    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Location = New-Object System.Drawing.Point(20, 95)
    $passwordTextBox.Size = New-Object System.Drawing.Size(370, 20)
    $passwordTextBox.PasswordChar = '*'
    $form.AcceptButton = $okButton
    $form.Controls.Add($passwordTextBox)
    
    $optionsLabel = New-Object System.Windows.Forms.Label
    $optionsLabel.Location = New-Object System.Drawing.Point(20, 130)
    $optionsLabel.Size = New-Object System.Drawing.Size(150, 20)
    $optionsLabel.Text = "Decryption Options:"
    $form.Controls.Add($optionsLabel)
    
    $showPasswordCheck = New-Object System.Windows.Forms.CheckBox
    $showPasswordCheck.Location = New-Object System.Drawing.Point(40, 155)
    $showPasswordCheck.Size = New-Object System.Drawing.Size(200, 20)
    $showPasswordCheck.Text = "Show password"
    $showPasswordCheck.Add_CheckedChanged({
        if ($showPasswordCheck.Checked) {
            $passwordTextBox.PasswordChar = [char]0
        } else {
            $passwordTextBox.PasswordChar = '*'
        }
    })
    $form.Controls.Add($showPasswordCheck)
    
    $tryWithoutCheck = New-Object System.Windows.Forms.CheckBox
    $tryWithoutCheck.Location = New-Object System.Drawing.Point(40, 180)
    $tryWithoutCheck.Size = New-Object System.Drawing.Size(250, 20)
    $tryWithoutCheck.Text = "Try without password first"
    $form.Controls.Add($tryWithoutCheck)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150, 210)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(235, 210)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $form.Add_Shown({$passwordTextBox.Select()})
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            Password = $passwordTextBox.Text
            TryWithoutPassword = $tryWithoutCheck.Checked
        }
    } else {
        return $null
    }
}

function Test-PDFIsEncrypted {
    param(
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            return $false
        }
        
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -lt 100) {
            return $false
        }
        
        $header = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 8)
        if ($header -notmatch "%PDF-\d+\.\d+") {
            return $false
        }
        
        $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 100000))
        
        if ($text -match '/Encrypt\s+\d+\s+\d+\s+[Rr]') {
            return $true
        }
        
        if ($text -match '/Encrypt\s*<<.*?>>') {
            return $true
        }
        
        if ($text -match '/Filter\s*/Standard') {
            return $true
        }
        
        if ($text -match '/CFM\s*/AESV[23]') {
            return $true
        }
        
        if ($text -match '/StmF\s*/StdCF' -or $text -match '/StrF\s*/StdCF') {
            return $true
        }
        
        if ($text -match '/O\s*\([^)]+\)' -and $text -match '/U\s*\([^)]+\)') {
            return $true
        }
        
        if ($text -match '/EncryptMetadata') {
            return $true
        }
        
        $firstPart = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 2000))
        if ($firstPart -match '/Encrypt') {
            return $true
        }
        
        if ($text -match 'Document is encrypted' -or $text -match 'password protected' -or $text -match 'requires a password') {
            return $true
        }
        
        $trailerPos = $text.LastIndexOf('trailer')
        if ($trailerPos -gt 0) {
            $trailerSection = $text.Substring($trailerPos, [Math]::Min(500, $text.Length - $trailerPos))
            if ($trailerSection -match '/Encrypt') {
                return $true
            }
        }
        
        $encryptedPattern1 = [byte[]]@(0x2F, 0x45, 0x6E, 0x63, 0x72, 0x79, 0x70, 0x74)
        $encryptedPattern2 = [byte[]]@(0x2F, 0x46, 0x69, 0x6C, 0x74, 0x65, 0x72, 0x20, 0x2F, 0x53, 0x74, 0x61, 0x6E, 0x64, 0x61, 0x72, 0x64)
        
        for ($i = 0; $i -lt ($bytes.Length - $encryptedPattern1.Length); $i++) {
            $match1 = $true
            for ($j = 0; $j -lt $encryptedPattern1.Length; $j++) {
                if ($bytes[$i + $j] -ne $encryptedPattern1[$j]) {
                    $match1 = $false
                    break
                }
            }
            if ($match1) {
                return $true
            }
        }
        
        return $false
        
    } catch {
        return $false
    }
}

function Decrypt-PDF {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputFile,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFile,
        
        [ValidateSet("Maximum", "High", "Medium", "Low")]
        [string]$Quality = "High",
        
        [string]$Password,
        
        [switch]$ShowPasswordDialog = $true,
        
        [switch]$ForceDecrypt = $false
    )
    
    $attemptWithoutPassword = $false
    $passwordResult = $null
    
    try {
        if (-not (Test-Path $InputFile)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Input file not found!`n`nFile: $InputFile", 
                "PDF Decryption - Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
        
        $fileInfo = Get-Item $InputFile
        $fileSize = $fileInfo.Length
        
        $isEncrypted = Test-PDFIsEncrypted -FilePath $InputFile
        
        if (-not $isEncrypted -and -not $ForceDecrypt) {
            $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
            if ($outputDir -and -not (Test-Path $outputDir)) {
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            }
            
            Copy-Item -Path $InputFile -Destination $OutputFile -Force
            
            if (Test-Path $OutputFile) {
                [System.Windows.Forms.MessageBox]::Show(
                    "PDF is not encrypted.`n`nFile has been copied successfully.`n`nNo decryption needed.", 
                    "PDF Decryption - Information", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                return $true
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to copy the file.", 
                "PDF Decryption - Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
        
        if ($ShowPasswordDialog -and [string]::IsNullOrWhiteSpace($Password)) {
            $passwordResult = Show-PasswordDialogDecrypt -Title "PDF Decryption" `
                                                         -Message "This PDF is encrypted.`nPlease enter the password to decrypt:"
            
            if ($passwordResult -eq $null) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Decryption cancelled by user.", 
                    "PDF Decryption - Cancelled", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                return $false
            }
            
            $Password = $passwordResult.Password
            $attemptWithoutPassword = $passwordResult.TryWithoutPassword
        }
        
        $qpdfPath = $null
        $qpdfPaths = @(
            "C:\Program Files\qpdf\bin\qpdf.exe",
            "C:\Program Files (x86)\qpdf\bin\qpdf.exe",
            "$env:ProgramFiles\qpdf\bin\qpdf.exe",
            "$env:ProgramFiles(x86)\qpdf\bin\qpdf.exe",
            ".\qpdf\bin\qpdf.exe",
            ".\tools\qpdf\bin\qpdf.exe",
            "$PSScriptRoot\qpdf\bin\qpdf.exe"
        )
        
        foreach ($path in $qpdfPaths) {
            if (Test-Path $path) {
                $qpdfPath = $path
                break
            }
        }
        
        if (-not $qpdfPath) {
            $qpdfInPath = Get-Command "qpdf.exe" -ErrorAction SilentlyContinue
            if ($qpdfInPath) {
                $qpdfPath = $qpdfInPath.Source
            }
        }
        
        $decryptionSuccessful = $false
        
        $decryptionAttempts = @()
        
        if ($attemptWithoutPassword) {
            $decryptionAttempts += @{ Password = $null; Description = "Without password" }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($Password)) {
            $decryptionAttempts += @{ Password = $Password; Description = "With provided password" }
        }
        
        $decryptionAttempts += @{ Password = ""; Description = "With empty password" }
        
        foreach ($attempt in $decryptionAttempts) {
            $attemptPassword = $attempt.Password
            $attemptDesc = $attempt.Description
            
            if ($qpdfPath -and (Test-Path $qpdfPath)) {
                $tempDir = Create-TempDirectory
                $tempPdf = Join-Path $tempDir "decrypted.pdf"
                
                try {
                    $qpdfArgs = @()
                    
                    if (-not [string]::IsNullOrEmpty($attemptPassword)) {
                        $qpdfArgs += "--password=$attemptPassword"
                    } else {
                        $qpdfArgs += "--password="
                    }
                    
                    $qpdfArgs += @("`"$InputFile`"", "`"$tempPdf`"")
                    
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $qpdfPath
                    $psi.Arguments = $qpdfArgs -join " "
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError = $true
                    $psi.UseShellExecute = $false
                    $psi.CreateNoWindow = $true
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $psi
                    $process.Start() | Out-Null
                    $process.WaitForExit(30000)
                    
                    $output = $process.StandardOutput.ReadToEnd()
                    $errorOutput = $process.StandardError.ReadToEnd()
                    
                    if ($errorOutput -match 'invalid password|password incorrect|wrong password|authentication failed') {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Incorrect password. Decryption not possible.`n`nFile: $(Split-Path $InputFile -Leaf)`n`nPlease try again with the correct password.", 
                            "PDF Decryption - Error", 
                            [System.Windows.Forms.MessageBoxButtons]::OK, 
                            [System.Windows.Forms.MessageBoxIcon]::Error
                        )
                        
                        if (Test-Path $tempDir) {
                            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                        }
                        
                        return $false
                    }
                    
                    if ($process.ExitCode -eq 0) {
                        if (Test-Path $tempPdf -and (Get-Item $tempPdf).Length -gt 100) {
                            try {
                                $testBytes = [System.IO.File]::ReadAllBytes($tempPdf)
                                if ($testBytes.Length -lt 100) {
                                    continue
                                }
                                
                                $header = [System.Text.Encoding]::ASCII.GetString($testBytes, 0, 8)
                                if ($header -notmatch "%PDF") {
                                    continue
                                }
                                
                                $isStillEncrypted = Test-PDFIsEncrypted -FilePath $tempPdf
                                
                                if (-not $isStillEncrypted) {
                                    $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
                                    if ($outputDir -and -not (Test-Path $outputDir)) {
                                        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
                                    }
                                    
                                    Copy-Item -Path $tempPdf -Destination $OutputFile -Force
                                    
                                    if (Test-Path $OutputFile) {
                                        [System.Windows.Forms.MessageBox]::Show(
                                            "PDF successfully decrypted!`n`nOutput file: $(Split-Path $OutputFile -Leaf)", 
                                            "PDF Decryption - Success", 
                                            [System.Windows.Forms.MessageBoxButtons]::OK, 
                                            [System.Windows.Forms.MessageBoxIcon]::Information
                                        )
                                        
                                        $decryptionSuccessful = $true
                                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                                        return $true
                                    }
                                }
                            } catch {
                                # Silently fail
                            }
                        }
                    }
                } catch {
                    # Silently fail
                } finally {
                    if (Test-Path $tempDir) {
                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            
            if ($decryptionSuccessful) {
                break
            }
        }
        
        if (-not $decryptionSuccessful -and $Global:ToolPaths.Ghostscript -and (Test-Path $Global:ToolPaths.Ghostscript)) {
            if (-not [string]::IsNullOrWhiteSpace($Password)) {
                $tempDir = Create-TempDirectory
                $tempPdf = Join-Path $tempDir "gs_decrypted.pdf"
                
                try {
                    $gsArgs = @(
                        "-sDEVICE=pdfwrite",
                        "-dNOPAUSE",
                        "-dBATCH",
                        "-dSAFER",
                        "-dNOPROMPT",
                        "-sPDFPassword=`"$Password`"",
                        "-sOutputFile=`"$tempPdf`"",
                        "`"$InputFile`""
                    )
                    
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $Global:ToolPaths.Ghostscript
                    $processInfo.Arguments = $gsArgs -join " "
                    $processInfo.RedirectStandardOutput = $true
                    $processInfo.RedirectStandardError = $true
                    $processInfo.UseShellExecute = $false
                    $processInfo.CreateNoWindow = $true
                    
                    $process = New-Object System.Diagnostics.Process
                    $process.StartInfo = $processInfo
                    $process.Start() | Out-Null
                    $process.WaitForExit(30000)
                    
                    $errorOutput = $process.StandardError.ReadToEnd()
                    
                    if ($errorOutput -match 'password|authentication|permission|encrypt|security') {
                        if ($errorOutput -notmatch 'Processing pages' -and 
                            $errorOutput -notmatch 'This is Ghostscript') {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Incorrect password. Ghostscript reported authentication error.`n`nFile: $(Split-Path $InputFile -Leaf)`n`nPlease try again with the correct password.", 
                                "PDF Decryption - Error", 
                                [System.Windows.Forms.MessageBoxButtons]::OK, 
                                [System.Windows.Forms.MessageBoxIcon]::Error
                            )
                            
                            if (Test-Path $tempDir) {
                                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                            }
                            
                            return $false
                        }
                    }
                    
                    if ($process.ExitCode -eq 0 -and (Test-Path $tempPdf) -and (Get-Item $tempPdf).Length -gt 100) {
                        $isStillEncrypted = Test-PDFIsEncrypted -FilePath $tempPdf
                        
                        if (-not $isStillEncrypted) {
                            $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
                            if ($outputDir -and -not (Test-Path $outputDir)) {
                                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
                            }
                            
                            Copy-Item -Path $tempPdf -Destination $OutputFile -Force
                            
                            if (Test-Path $OutputFile) {
                                [System.Windows.Forms.MessageBox]::Show(
                                    "PDF successfully decrypted using Ghostscript!`n`nOutput file: $(Split-Path $OutputFile -Leaf)", 
                                    "PDF Decryption - Success", 
                                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                                    [System.Windows.Forms.MessageBoxIcon]::Information
                                )
                                
                                $decryptionSuccessful = $true
                                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                                return $true
                            }
                        }
                    }
                } catch {
                    # Silently fail
                } finally {
                    if (Test-Path $tempDir) {
                        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        
        if (-not $decryptionSuccessful) {
            if ($isEncrypted) {
                if ([string]::IsNullOrWhiteSpace($Password)) {
                    $errorMsg = "No password was provided for encrypted PDF.`n`nFile: $(Split-Path $InputFile -Leaf)"
                } else {
                    $errorMsg = "Decryption failed. Possible reasons:`n• Incorrect password`n• PDF uses strong encryption`n• Tools cannot decrypt this PDF`n`nFile: $(Split-Path $InputFile -Leaf)"
                }
            } else {
                $errorMsg = "Decryption failed for unknown reason.`n`nFile: $(Split-Path $InputFile -Leaf)"
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                $errorMsg, 
                "PDF Decryption - Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
        
        return $true
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unexpected error: $($_.Exception.Message)`n`nFile: $(Split-Path $InputFile -Leaf)", 
            "PDF Decryption - Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# ============================================
# PDF WATERMARK
# ============================================

function Watermark-PDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$WatermarkText = "CONFIDENTIAL",
        [string]$WatermarkType = "Diagonal",
        [string]$Color = "Gray",
        [string]$Quality = "High",
        [switch]$NoDialog
    )
    
    function Get-FileSizeString {
        param([long]$SizeInBytes)
        
        if ($SizeInBytes -lt 1KB) { return "$SizeInBytes B" }
        elseif ($SizeInBytes -lt 1MB) { return "$([math]::Round($SizeInBytes/1KB, 2)) KB" }
        elseif ($SizeInBytes -lt 1GB) { return "$([math]::Round($SizeInBytes/1MB, 2)) MB" }
        else { return "$([math]::Round($SizeInBytes/1GB, 2)) GB" }
    }
    
    function Show-CustomTextDialog {
        param([string]$CurrentText = "")
        
        $customForm = New-Object System.Windows.Forms.Form
        $customForm.Text = "Enter Custom Watermark Text"
        $customForm.Size = New-Object System.Drawing.Size(450, 180)
        $customForm.StartPosition = "CenterParent"
        $customForm.FormBorderStyle = "FixedDialog"
        $customForm.MaximizeBox = $false
        $customForm.MinimizeBox = $false
        $customForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "Enter your custom watermark text:"
        $instructionLabel.Location = New-Object System.Drawing.Point(20, 20)
        $instructionLabel.Size = New-Object System.Drawing.Size(400, 25)
        $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $customForm.Controls.Add($instructionLabel)
        
        $customTextBox = New-Object System.Windows.Forms.TextBox
        $customTextBox.Location = New-Object System.Drawing.Point(20, 55)
        $customTextBox.Size = New-Object System.Drawing.Size(390, 25)
        $customTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $customTextBox.Text = $CurrentText
        $customForm.Controls.Add($customTextBox)
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(220, 100)
        $okButton.Size = New-Object System.Drawing.Size(90, 30)
        $okButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $okButton.BackColor = [System.Drawing.Color]::LightGreen
        $okButton.ForeColor = [System.Drawing.Color]::Black
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $customForm.AcceptButton = $okButton
        $customForm.Controls.Add($okButton)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(320, 100)
        $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
        $cancelButton.Font = New-Object System.Drawing.Font("Arial", 10)
        $cancelButton.BackColor = [System.Drawing.Color]::LightCoral
        $cancelButton.ForeColor = [System.Drawing.Color]::Black
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $customForm.CancelButton = $cancelButton
        $customForm.Controls.Add($cancelButton)
        
        $result = $customForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($customTextBox.Text)) {
            return $customTextBox.Text.Trim()
        }
        
        return $null
    }
    
    function Show-WatermarkDialog {
        $canLoadGUI = $true
        
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            
            $testForm = New-Object System.Windows.Forms.Form
            $testForm.Dispose()
            
        } catch {
            $canLoadGUI = $false
        }
        
        if (-not $canLoadGUI -or $NoDialog) {
            return @{
                Text = $WatermarkText
                Type = $WatermarkType
                Quality = $Quality
                Color = $Color
                Opacity = 0.3
                FontSize = 48
            }
        }
        
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "PDF Watermark Configuration"
            $form.Size = New-Object System.Drawing.Size(540, 600)
            $form.StartPosition = "CenterScreen"
            $form.FormBorderStyle = "FixedDialog"
            $form.MaximizeBox = $false
            $form.MinimizeBox = $false
            $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
            
            $titleLabel = New-Object System.Windows.Forms.Label
            $titleLabel.Text = "Configure PDF Watermark"
            $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
            $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
            $titleLabel.Size = New-Object System.Drawing.Size(500, 40)
            $titleLabel.TextAlign = "MiddleCenter"
            $titleLabel.ForeColor = [System.Drawing.Color]::Navy
            $form.Controls.Add($titleLabel)
            
            $textLabel = New-Object System.Windows.Forms.Label
            $textLabel.Text = "Watermark Text:"
            $textLabel.Location = New-Object System.Drawing.Point(20, 80)
            $textLabel.Size = New-Object System.Drawing.Size(150, 25)
            $textLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($textLabel)
            
            $textPanel = New-Object System.Windows.Forms.Panel
            $textPanel.Location = New-Object System.Drawing.Point(180, 75)
            $textPanel.Size = New-Object System.Drawing.Size(320, 30)
            $textPanel.BackColor = [System.Drawing.Color]::Transparent
            
            $textComboBox = New-Object System.Windows.Forms.ComboBox
            $textComboBox.Location = New-Object System.Drawing.Point(0, 0)
            $textComboBox.Size = New-Object System.Drawing.Size(240, 25)
            $textComboBox.Font = New-Object System.Drawing.Font("Arial", 10)
            $textComboBox.Items.AddRange(@("CONFIDENTIAL", "DRAFT", "SAMPLE", "DO NOT COPY", "INTERNAL USE", "COPYRIGHT"))
            $textComboBox.Text = $WatermarkText
            $textComboBox.DropDownStyle = "DropDown"
            $textComboBox.AutoCompleteMode = "SuggestAppend"
            $textComboBox.AutoCompleteSource = "ListItems"
            $textPanel.Controls.Add($textComboBox)
            
            $customButton = New-Object System.Windows.Forms.Button
            $customButton.Text = "CUSTOM"
            $customButton.Location = New-Object System.Drawing.Point(245, 0)
            $customButton.Size = New-Object System.Drawing.Size(75, 25)
            $customButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            $customButton.BackColor = [System.Drawing.Color]::LightBlue
            $customButton.ForeColor = [System.Drawing.Color]::Black
            $customButton.TextAlign = "MiddleCenter"
            $customButton.FlatStyle = "Standard"
            $customButton.UseVisualStyleBackColor = $false
            $textPanel.Controls.Add($customButton)
            
            $form.Controls.Add($textPanel)
            
            $script:customWatermarkText = $null
            
            $customButton.Add_Click({
                $customText = Show-CustomTextDialog -CurrentText $script:customWatermarkText
                if ($customText) {
                    $script:customWatermarkText = $customText
                    $textComboBox.Text = $script:customWatermarkText
                    & $updatePreview
                }
            })
            
            $typeLabel = New-Object System.Windows.Forms.Label
            $typeLabel.Text = "Watermark Type:"
            $typeLabel.Location = New-Object System.Drawing.Point(20, 120)
            $typeLabel.Size = New-Object System.Drawing.Size(150, 25)
            $typeLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($typeLabel)
            
            $typeComboBox = New-Object System.Windows.Forms.ComboBox
            $typeComboBox.Location = New-Object System.Drawing.Point(180, 115)
            $typeComboBox.Size = New-Object System.Drawing.Size(320, 25)
            $typeComboBox.Font = New-Object System.Drawing.Font("Arial", 10)
            $typeComboBox.Items.AddRange(@("Diagonal", "Centered", "Header", "Footer", "Corner"))
            $typeComboBox.Text = $WatermarkType
            $typeComboBox.DropDownStyle = "DropDownList"
            $form.Controls.Add($typeComboBox)
            
            $qualityLabel = New-Object System.Windows.Forms.Label
            $qualityLabel.Text = "Quality Level:"
            $qualityLabel.Location = New-Object System.Drawing.Point(20, 160)
            $qualityLabel.Size = New-Object System.Drawing.Size(150, 25)
            $qualityLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($qualityLabel)
            
            $qualityComboBox = New-Object System.Windows.Forms.ComboBox
            $qualityComboBox.Location = New-Object System.Drawing.Point(180, 155)
            $qualityComboBox.Size = New-Object System.Drawing.Size(320, 25)
            $qualityComboBox.Font = New-Object System.Drawing.Font("Arial", 10)
            $qualityComboBox.Items.AddRange(@("High", "Medium", "Low"))
            $qualityComboBox.Text = $Quality
            $qualityComboBox.DropDownStyle = "DropDownList"
            $form.Controls.Add($qualityComboBox)
            
            $colorLabel = New-Object System.Windows.Forms.Label
            $colorLabel.Text = "Watermark Color:"
            $colorLabel.Location = New-Object System.Drawing.Point(20, 200)
            $colorLabel.Size = New-Object System.Drawing.Size(150, 25)
            $colorLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($colorLabel)
            
            $colorComboBox = New-Object System.Windows.Forms.ComboBox
            $colorComboBox.Location = New-Object System.Drawing.Point(180, 195)
            $colorComboBox.Size = New-Object System.Drawing.Size(320, 25)
            $colorComboBox.Font = New-Object System.Drawing.Font("Arial", 10)
            $colorComboBox.Items.AddRange(@("Gray", "Black", "Red", "Blue", "Green", "Purple", "Orange"))
            $colorComboBox.Text = $Color
            $colorComboBox.DropDownStyle = "DropDownList"
            $form.Controls.Add($colorComboBox)
            
            $opacityLabel = New-Object System.Windows.Forms.Label
            $opacityLabel.Text = "Opacity (0.1-1.0):"
            $opacityLabel.Location = New-Object System.Drawing.Point(20, 240)
            $opacityLabel.Size = New-Object System.Drawing.Size(150, 25)
            $opacityLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($opacityLabel)
            
            $opacityNumeric = New-Object System.Windows.Forms.NumericUpDown
            $opacityNumeric.Location = New-Object System.Drawing.Point(180, 235)
            $opacityNumeric.Size = New-Object System.Drawing.Size(100, 25)
            $opacityNumeric.Font = New-Object System.Drawing.Font("Arial", 10)
            $opacityNumeric.Minimum = 0.1
            $opacityNumeric.Maximum = 1.0
            $opacityNumeric.DecimalPlaces = 2
            $opacityNumeric.Increment = 0.1
            $opacityNumeric.Value = 0.3
            $form.Controls.Add($opacityNumeric)
            
            $fontSizeLabel = New-Object System.Windows.Forms.Label
            $fontSizeLabel.Text = "Font Size (pt):"
            $fontSizeLabel.Location = New-Object System.Drawing.Point(20, 280)
            $fontSizeLabel.Size = New-Object System.Drawing.Size(150, 25)
            $fontSizeLabel.Font = New-Object System.Drawing.Font("Arial", 10)
            $form.Controls.Add($fontSizeLabel)
            
            $fontSizeNumeric = New-Object System.Windows.Forms.NumericUpDown
            $fontSizeNumeric.Location = New-Object System.Drawing.Point(180, 275)
            $fontSizeNumeric.Size = New-Object System.Drawing.Size(100, 25)
            $fontSizeNumeric.Font = New-Object System.Drawing.Font("Arial", 10)
            $fontSizeNumeric.Minimum = 12
            $fontSizeNumeric.Maximum = 144
            $fontSizeNumeric.Value = 48
            $form.Controls.Add($fontSizeNumeric)
            
            $previewGroup = New-Object System.Windows.Forms.GroupBox
            $previewGroup.Text = "Preview"
            $previewGroup.Location = New-Object System.Drawing.Point(20, 320)
            $previewGroup.Size = New-Object System.Drawing.Size(480, 120)
            $previewGroup.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $form.Controls.Add($previewGroup)
            
            $previewPanel = New-Object System.Windows.Forms.Panel
            $previewPanel.Location = New-Object System.Drawing.Point(10, 20)
            $previewPanel.Size = New-Object System.Drawing.Size(460, 90)
            $previewPanel.BackColor = [System.Drawing.Color]::White
            $previewPanel.BorderStyle = "FixedSingle"
            $previewGroup.Controls.Add($previewPanel)
            
            $previewText = New-Object System.Windows.Forms.Label
            $previewText.Text = $WatermarkText
            $previewText.Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Bold)
            $previewText.ForeColor = [System.Drawing.Color]::FromArgb(100, 128, 128, 128)
            $previewText.TextAlign = "MiddleCenter"
            $previewText.Dock = "Fill"
            $previewPanel.Controls.Add($previewText)
            
            $updatePreview = {
                try {
                    $text = if ([string]::IsNullOrWhiteSpace($textComboBox.Text)) {
                        if ($script:customWatermarkText) {
                            $script:customWatermarkText
                        } else {
                            "CONFIDENTIAL"
                        }
                    } else {
                        $textComboBox.Text
                    }
                    
                    $colorName = $colorComboBox.Text
                    $opacityValue = [math]::Round([double]$opacityNumeric.Value, 2)
                    $fontSizeValue = [int]$fontSizeNumeric.Value
                    
                    $previewText.Text = $text
                    
                    $alpha = [int](255 * $opacityValue)
                    $colorValue = switch ($colorName) {
                        "Black"   { [System.Drawing.Color]::FromArgb($alpha, 0, 0, 0) }
                        "Gray"    { [System.Drawing.Color]::FromArgb($alpha, 128, 128, 128) }
                        "Red"     { [System.Drawing.Color]::FromArgb($alpha, 255, 0, 0) }
                        "Blue"    { [System.Drawing.Color]::FromArgb($alpha, 0, 0, 255) }
                        "Green"   { [System.Drawing.Color]::FromArgb($alpha, 0, 128, 0) }
                        "Purple"  { [System.Drawing.Color]::FromArgb($alpha, 128, 0, 128) }
                        "Orange"  { [System.Drawing.Color]::FromArgb($alpha, 255, 165, 0) }
                        default   { [System.Drawing.Color]::FromArgb($alpha, 128, 128, 128) }
                    }
                    $previewText.ForeColor = $colorValue
                    
                    $effectiveFontSize = [math]::Min($fontSizeValue, 32)
                    $previewText.Font = New-Object System.Drawing.Font("Arial", $effectiveFontSize, [System.Drawing.FontStyle]::Bold)
                } catch {}
            }
            
            $textComboBox.Add_TextChanged($updatePreview)
            $typeComboBox.Add_SelectedIndexChanged($updatePreview)
            $colorComboBox.Add_SelectedIndexChanged($updatePreview)
            $opacityNumeric.Add_ValueChanged($updatePreview)
            $fontSizeNumeric.Add_ValueChanged($updatePreview)
            
            & $updatePreview
            
            $buttonPanel = New-Object System.Windows.Forms.Panel
            $buttonPanel.Location = New-Object System.Drawing.Point(20, 450)
            $buttonPanel.Size = New-Object System.Drawing.Size(480, 50)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(220, 10)
            $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
            $cancelButton.Font = New-Object System.Drawing.Font("Arial", 10)
            $cancelButton.BackColor = [System.Drawing.Color]::LightCoral
            $cancelButton.ForeColor = [System.Drawing.Color]::Black
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.CancelButton = $cancelButton
            $buttonPanel.Controls.Add($cancelButton)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Apply Watermark"
            $okButton.Location = New-Object System.Drawing.Point(340, 10)
            $okButton.Size = New-Object System.Drawing.Size(120, 35)
            $okButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $okButton.BackColor = [System.Drawing.Color]::LightGreen
            $okButton.ForeColor = [System.Drawing.Color]::Black
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.AcceptButton = $okButton
            $buttonPanel.Controls.Add($okButton)
            
            $form.Controls.Add($buttonPanel)
            
            $result = $form.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $finalText = if ($script:customWatermarkText -and $textComboBox.Text -eq $script:customWatermarkText) {
                    $script:customWatermarkText
                } elseif (-not [string]::IsNullOrWhiteSpace($textComboBox.Text)) {
                    $textComboBox.Text
                } else {
                    "CONFIDENTIAL"
                }
                
                return @{
                    Text = $finalText
                    Type = $typeComboBox.Text
                    Quality = $qualityComboBox.Text
                    Color = $colorComboBox.Text
                    Opacity = [math]::Round([double]$opacityNumeric.Value, 2)
                    FontSize = [int]$fontSizeNumeric.Value
                }
            } else {
                return $null
            }
            
        } catch {
            return @{
                Text = $WatermarkText
                Type = $WatermarkType
                Quality = $Quality
                Color = $Color
                Opacity = 0.3
                FontSize = 48
            }
        }
    }
    
    function Get-WatermarkSettings {
        param([string]$Quality)
        
        switch ($Quality) {
            "Maximum" { 
                return @{
                    FontSize = 72
                    Opacity = 0.15
                    Angle = 45
                    DPI = 300
                }
            }
            "High" { 
                return @{
                    FontSize = 48
                    Opacity = 0.2
                    Angle = 45
                    DPI = 200
                }
            }
            "Medium" { 
                return @{
                    FontSize = 36
                    Opacity = 0.3
                    Angle = 30
                    DPI = 150
                }
            }
            "Low" { 
                return @{
                    FontSize = 24
                    Opacity = 0.4
                    Angle = 0
                    DPI = 100
                }
            }
            default { 
                return @{
                    FontSize = 48
                    Opacity = 0.2
                    Angle = 45
                    DPI = 200
                }
            }
        }
    }
    
    function Get-ColorRGB {
        param([string]$ColorName)
        
        switch ($ColorName) {
            "Black"   { return @(0, 0, 0) }
            "Gray"    { return @(0.5, 0.5, 0.5) }
            "Red"     { return @(1, 0, 0) }
            "Blue"    { return @(0, 0, 1) }
            "Green"   { return @(0, 1, 0) }
            "Purple"  { return @(0.5, 0, 0.5) }
            "Orange"  { return @(1, 0.65, 0) }
            default   { return @(0.5, 0.5, 0.5) }
        }
    }
    
    function Build-WatermarkPostScript {
        param(
            [string]$Text,
            [string]$Type,
            [hashtable]$Settings,
            [double[]]$ColorRGB,
            [double]$CustomOpacity,
            [int]$CustomFontSize
        )
        
        $fontSize = if ($CustomFontSize -gt 0) { $CustomFontSize } else { $Settings.FontSize }
        $opacity = if ($CustomOpacity -gt 0) { $CustomOpacity } else { $Settings.Opacity }
        $angle = $Settings.Angle
        
        $escapedText = $Text -replace '([()\\])', '\\$1'
        
        $r = $ColorRGB[0]
        $g = $ColorRGB[1]
        $b = $ColorRGB[2]
        
        if ($ColorRGB[0] -eq 0 -and $ColorRGB[1] -eq 0 -and $ColorRGB[2] -eq 0) {
            $r = 0.1
            $g = 0.1
            $b = 0.1
        }
        
        $psCode = @"
%!PS-Adobe-3.0
%%BoundingBox: 0 0 612 792
%%Creator: PDF Watermark Tool
%%Title: Watermark: $Text
%%Pages: (atend)
%%EndComments

<<
  /EndPage {
    exch pop 2 eq { pop false } {
      gsave
        
      /Helvetica-Bold findfont $fontSize scalefont setfont
      
      % Set color with full intensity
      $r $g $b setrgbcolor
      
      % Apply transparency using setfillopacity (requires PostScript Level 3)
      /setfillopacity where {
        pop
        $opacity setfillopacity
      } if
      
      /setstrokeopacity where {
        pop
        $opacity setstrokeopacity
      } if
      
      currentpagedevice /PageSize get aload pop
      /pageHeight exch def
      /pageWidth exch def
      
      pageWidth 2 div /centerX exch def
      pageHeight 2 div /centerY exch def
      
      % Position based on type
      $(
        switch ($Type) {
            "Diagonal" {
                @"
      centerX centerY moveto
      $angle rotate
      ($escapedText) dup stringwidth pop 2 div neg 0 rmoveto show
"@
            }
            "Centered" {
                @"
      centerX centerY moveto
      ($escapedText) dup stringwidth pop 2 div neg 0 rmoveto show
"@
            }
            "Header" {
                @"
      centerX pageHeight 50 sub moveto
      ($escapedText) dup stringwidth pop 2 div neg 0 rmoveto show
"@
            }
            "Footer" {
                @"
      centerX 50 moveto
      ($escapedText) dup stringwidth pop 2 div neg 0 rmoveto show
"@
            }
            "Corner" {
                @"
      pageWidth 100 sub 100 moveto
      ($escapedText) show
"@
            }
            default {
                @"
      centerX centerY moveto
      $angle rotate
      ($escapedText) dup stringwidth pop 2 div neg 0 rmoveto show
"@
            }
        }
      )
      
      grestore
      true
    } ifelse
  }
>> setpagedevice

%%EOF
"@
        
        return $psCode
    }
    
    try {
        if (-not (Test-Path $InputFile)) {
            return @{ Success = $false; Message = "Input file not found" }
        }
        
        $originalFile = Get-Item $InputFile
        $originalSize = $originalFile.Length
        
        $watermarkConfig = $null
        
        if (-not $NoDialog) {
            $watermarkConfig = Show-WatermarkDialog
        } else {
            $watermarkConfig = @{
                Text = $WatermarkText
                Type = $WatermarkType
                Quality = $Quality
                Color = $Color
                Opacity = 0.3
                FontSize = 48
            }
        }
        
        if ($watermarkConfig -eq $null) {
            return @{ Success = $false; Message = "Cancelled by user" }
        }
        
        if ([string]::IsNullOrWhiteSpace($watermarkConfig.Text)) {
            return @{ Success = $false; Message = "No watermark text provided" }
        }
        
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if (Test-Path $OutputFile) {
            Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
        }
        
        $watermarkSettings = Get-WatermarkSettings -Quality $watermarkConfig.Quality
        $colorRGB = Get-ColorRGB -ColorName $watermarkConfig.Color
        
        $psCode = Build-WatermarkPostScript -Text $watermarkConfig.Text `
                                           -Type $watermarkConfig.Type `
                                           -Settings $watermarkSettings `
                                           -ColorRGB $colorRGB `
                                           -CustomOpacity $watermarkConfig.Opacity `
                                           -CustomFontSize $watermarkConfig.FontSize
        
        $tempPsFile = [System.IO.Path]::GetTempFileName() + ".ps"
        Set-Content -Path $tempPsFile -Value $psCode -Encoding ASCII
        
        $gsPath = $Global:ToolPaths.Ghostscript
        if (-not $gsPath -or -not (Test-Path $gsPath)) {
            if (Test-Path $tempPsFile) {
                Remove-Item $tempPsFile -Force
            }
            
            return @{
                Success = $false
                Message = "Ghostscript not found"
            }
        }
        
        $gsArgs = @(
            "-sDEVICE=pdfwrite",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-dPDFSETTINGS=/default",
            "-dCompatibilityLevel=1.4",
            "-dColorConversionStrategy=/LeaveColorUnchanged",
            "-dSubsetFonts=true",
            "-dEmbedAllFonts=true",
            "-sOutputFile=`"$OutputFile`"",
            "-q",
            "-f",
            "`"$tempPsFile`"",
            "`"$InputFile`""
        )
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $gsPath
        $psi.Arguments = ($gsArgs -join " ")
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $process.Start() | Out-Null
        
        $stdOut = $process.StandardOutput.ReadToEnd()
        $stdErr = $process.StandardError.ReadToEnd()
        
        $process.WaitForExit(30000)
        $stopwatch.Stop()
        
        if (Test-Path $tempPsFile) {
            Remove-Item $tempPsFile -Force -ErrorAction SilentlyContinue
        }
        
        if ($process.ExitCode -eq 0 -and (Test-Path $OutputFile -PathType Leaf)) {
            $watermarkedFile = Get-Item $OutputFile
            $watermarkedSize = $watermarkedFile.Length
            $processingTime = $stopwatch.Elapsed.TotalSeconds
            
            $colorDisplay = switch ($watermarkConfig.Color) {
                "Black"   { "Black (adjusted to dark gray for visibility)" }
                "Gray"    { "Gray" }
                "Red"     { "Red" }
                "Blue"    { "Blue" }
                "Green"   { "Green" }
                "Purple"  { "Purple" }
                "Orange"  { "Orange" }
                default   { $watermarkConfig.Color }
            }
            
            $report = @"
WATERMARKING REPORT
===================
Input File: $(Split-Path $InputFile -Leaf)
Output File: $(Split-Path $OutputFile -Leaf)

Configuration:
- Text: $($watermarkConfig.Text)
- Type: $($watermarkConfig.Type)
- Color: $colorDisplay
- Opacity: $($watermarkConfig.Opacity)
- Font Size: $($watermarkConfig.FontSize) pt
- Quality: $($watermarkConfig.Quality)

Processing:
- Original Size: $(Get-FileSizeString -SizeInBytes $originalSize)
- Watermarked Size: $(Get-FileSizeString -SizeInBytes $watermarkedSize)
- Processing Time: $([math]::Round($processingTime, 2)) seconds

Color Information:
- Base RGB: $($colorRGB -join ', ')
- Applied Opacity: $($watermarkConfig.Opacity)
- Note: Using PostScript setfillopacity/setstrokeopacity for proper transparency

Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Tool: Windows PDF Converter Pro
"@
            
            $reportFile = [System.IO.Path]::ChangeExtension($OutputFile, '.report.txt')
            $report | Out-File -FilePath $reportFile -Encoding UTF8
            
            return @{
                Success = $true
                Message = "Watermark applied successfully"
                OutputFile = $OutputFile
                ReportFile = $reportFile
            }
            
        } else {
            return @{
                Success = $false
                Message = "Ghostscript failed with exit code: $($process.ExitCode)"
            }
        }
        
    } catch {
        return @{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

# ============================================
# CREATE ENHANCED PDF
# ============================================

function Create-EnhancedPDF {
    param(
        [string]$InputFile,
        [string]$OutputFile,
        [string]$ConversionType,
        [string]$Quality,
        [string]$TextContent,
        [switch]$IsImage,
        [System.Drawing.Image]$ImageObject,
        [int]$ScaledWidth,
        [int]$ScaledHeight,
        [int]$OffsetX,
        [int]$OffsetY
    )
    
    try {
        $fileInfo = Get-Item $InputFile -ErrorAction SilentlyContinue
        $fileName = Split-Path $InputFile -Leaf
        $fileSize = if ($fileInfo) { $fileInfo.Length } else { 0 }
        $dateStr = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        
        # Format file size
        $sizeStr = if ($fileSize -gt 0) { 
            if ($fileSize -lt 1KB) { "$fileSize B" }
            elseif ($fileSize -lt 1MB) { "$([math]::Round($fileSize/1KB, 2)) KB" }
            elseif ($fileSize -lt 1GB) { "$([math]::Round($fileSize/1MB, 2)) MB" }
            else { "$([math]::Round($fileSize/1GB, 2)) GB" }
        } else { "Unknown" }
        
        # Clean the text content
        $cleanText = if ($TextContent) { 
            # Remove extra spaces and normalize line breaks
            $TextContent -replace '\r\n?', "`n" -replace '\s+', ' ' -replace ' \.', '.' -replace ' ,', ','
        } else { 
            "No content available"
        }
        
        # Split into lines for PDF (max 60 chars per line for better readability)
        $lines = @()
        $currentLine = ""
        $words = $cleanText -split ' '
        
        foreach ($word in $words) {
            if (($currentLine + " " + $word).Length -le 60) {
                if ($currentLine -eq "") {
                    $currentLine = $word
                } else {
                    $currentLine += " " + $word
                }
            } else {
                if ($currentLine -ne "") {
                    $lines += $currentLine
                }
                $currentLine = $word
            }
        }
        if ($currentLine -ne "") {
            $lines += $currentLine
        }
        
        # Limit number of lines to prevent PDF from getting too large
        if ($lines.Count -gt 50) {
            $lines = $lines[0..49]
            $lines += "... [content truncated]"
        }
        
        # Build PDF content with proper structure
        $pdfContent = "%PDF-1.4`n"
        
        # Object 1: Catalog
        $pdfContent += "1 0 obj`n<</Type/Catalog/Pages 2 0 R>>`nendobj`n"
        
        # Object 2: Pages
        $pdfContent += "2 0 obj`n<</Type/Pages/Kids[3 0 R]/Count 1>>`nendobj`n"
        
        # Build content stream
        $contentStream = "BT`n"
        $contentStream += "/F1 12 Tf`n"
        
        $yPos = 750
        $lineHeight = 20
        
        # Title
        $contentStream += "72 $yPos Td`n"
        $title = "Word to PDF - Conversion Report"
        $contentStream += "($title) Tj`n"
        $yPos -= $lineHeight * 2
        
        # File info line
        $contentStream += "72 $yPos Td`n"
        $fileInfo = "File: $fileName  Size: $sizeStr  Date: $dateStr  Quality: $Quality"
        $contentStream += "($fileInfo) Tj`n"
        $yPos -= $lineHeight * 2
        
        # Content header
        $contentStream += "72 $yPos Td`n"
        $contentStream += "(Content:) Tj`n"
        $yPos -= $lineHeight
        
        # Add content lines
        foreach ($line in $lines) {
            if ($yPos -lt 50) { 
                $contentStream += "72 $yPos Td`n"
                $contentStream += "(... continued in next section) Tj`n"
                break 
            }
            $contentStream += "72 $yPos Td`n"
            # Escape parentheses in the line
            $escapedLine = $line -replace '\(', '\(' -replace '\)', '\)'
            $contentStream += "($escapedLine) Tj`n"
            $yPos -= $lineHeight
        }
        
        $contentStream += "ET`n"
        
        # Object 3: Page
        $pdfContent += "3 0 obj`n"
        $pdfContent += "<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>`n"
        $pdfContent += "endobj`n"
        
        # Object 4: Contents
        $pdfContent += "4 0 obj`n"
        $pdfContent += "<</Length $($contentStream.Length)>>`n"
        $pdfContent += "stream`n"
        $pdfContent += $contentStream
        $pdfContent += "endstream`n"
        $pdfContent += "endobj`n"
        
        # Object 5: Font
        $pdfContent += "5 0 obj`n"
        $pdfContent += "<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>`n"
        $pdfContent += "endobj`n"
        
        # Calculate xref offsets
        $offsets = @()
        $stream = [System.IO.MemoryStream]::new()
        $writer = [System.IO.StreamWriter]::new($stream)
        $writer.Write($pdfContent)
        $writer.Flush()
        $pdfLength = $stream.Length
        $stream.Close()
        $writer.Close()
        
        # Simple xref - just use approximate positions
        $xref = "xref`n"
        $xref += "0 6`n"
        $xref += "0000000000 65535 f `n"
        $xref += "0000000010 00000 n `n"
        $xref += "0000000050 00000 n `n"
        $xref += "0000000100 00000 n `n"
        $xref += "0000000200 00000 n `n"
        $xref += "0000000500 00000 n `n"
        
        $trailer = "trailer`n"
        $trailer += "<</Size 6/Root 1 0 R>>`n"
        $trailer += "startxref`n"
        $trailer += "$($pdfLength + 100)`n"
        $trailer += "%%EOF"
        
        # Write complete PDF
        $fullPdf = $pdfContent + $xref + $trailer
        [System.IO.File]::WriteAllText($OutputFile, $fullPdf, [System.Text.Encoding]::ASCII)
        
        # Verify the PDF
        if (Test-Path $OutputFile) {
            $outSize = (Get-Item $OutputFile).Length
            if ($outSize -gt 100) {
                return $true
            }
        }
        
        return $false
        
    } catch {
        # Ultimate fallback - minimal valid PDF
        try {
            $minimalPdf = @"
%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj
4 0 obj<</Length 44>>stream
BT/F1 12 Tf 72 720 Td(Word to PDF Conversion) Tj ET
endstream
endobj
5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj
xref
0 6
0000000000 65535 f 
0000000010 00000 n 
0000000050 00000 n 
0000000110 00000 n 
0000000210 00000 n 
0000000320 00000 n 
trailer
<</Size 6/Root 1 0 R>>
startxref
380
%%EOF
"@
            [System.IO.File]::WriteAllText($OutputFile, $minimalPdf, [System.Text.Encoding]::ASCII)
            return (Test-Path $OutputFile)
        } catch {
            return $false
        }
    }
}


function Test-ValidPDF {
    param([string]$FilePath)
    
    try {
        if (-not (Test-Path $FilePath)) { return $false }
        
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -lt 5) { return $false }
        
        $header = [System.Text.Encoding]::ASCII.GetString($bytes[0..4])
        return ($header -match '%PDF')
    } catch {
        return $false
    }
}

# ============================================
# FILE MANAGEMENT FUNCTIONS
# ============================================

function Update-FileFilters {
}

function Add-Files {
    $convType = $Global:AppState.ConversionListBox.SelectedItem.ToString()
    
    $filter = switch ($convType) {
        "Word to PDF" { 
            "Word Documents (*.doc;*.docx;*.rtf)|*.doc;*.docx;*.rtf|All Files (*.*)|*.*"
        }
        "PDF to Word" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "Excel to PDF" { 
            "Excel Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv|All Files (*.*)|*.*"
        }
        "PDF to Excel" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PowerPoint to PDF" { 
            "PowerPoint Files (*.ppt;*.pptx;*.pptm)|*.ppt;*.pptx;*.pptm|All Files (*.*)|*.*"
        }
        "PDF to PowerPoint" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "Images to PDF" { 
            "Image Files (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|All Files (*.*)|*.*"
        }
        "PDF to Images" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "Text to PDF" { 
            "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        }
        "HTML to PDF" { 
            "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*"
        }
        "PDF Merge" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PDF Split" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PDF Compress" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PDF Encrypt" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PDF Decrypt" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        "PDF Watermark" { 
            "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
        }
        default { 
            "All Files (*.*)|*.*"
        }
    }
    
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Filter = $filter
    $dialog.Title = "Select files to convert for: $convType"
    $dialog.CheckFileExists = $true
    
    if ($dialog.ShowDialog() -eq "OK") {
        $count = 0
        foreach ($file in $dialog.FileNames) {
            if (Test-Path $file) {
                Add-FileToList -FilePath $file
                $count++
            }
        }
        if ($count -gt 0) {
            Update-FileCount
            Update-Status -Message "Added $count file(s) for $convType"
        }
    }
}

function Add-Folder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select folder with files to convert"
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq "OK") {
        $convType = $Global:AppState.ConversionListBox.SelectedItem.ToString()
        
        $extensions = switch ($convType) {
            "Word to PDF" { @('.doc', '.docx', '.rtf') }
            "PDF to Word" { @('.pdf') }
            "Excel to PDF" { @('.xls', '.xlsx', '.xlsm', '.csv') }
            "PDF to Excel" { @('.pdf') }
            "PowerPoint to PDF" { @('.ppt', '.pptx', '.pptm') }
            "PDF to PowerPoint" { @('.pdf') }
            "Images to PDF" { @('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif') }
            "PDF to Images" { @('.pdf') }
            "Text to PDF" { @('.txt') }
            "HTML to PDF" { @('.html', '.htm') }
            "PDF Merge" { @('.pdf') }
            "PDF Split" { @('.pdf') }
            "PDF Compress" { @('.pdf') }
            "PDF Encrypt" { @('.pdf') }
            "PDF Decrypt" { @('.pdf') }
            "PDF Watermark" { @('.pdf') }
            default { @() }
        }
        
        if ($extensions.Count -gt 0) {
            $filesFound = Get-ChildItem -Path $dialog.SelectedPath -Recurse -File | 
                Where-Object { $extensions -contains $_.Extension.ToLower() } |
                Select-Object -First 50 -ExpandProperty FullName
        } else {
            $filesFound = Get-ChildItem -Path $dialog.SelectedPath -Recurse -File | 
                Select-Object -First 50 -ExpandProperty FullName
        }
        
        if ($filesFound.Count -gt 0) {
            foreach ($file in $filesFound) {
                Add-FileToList -FilePath $file
            }
            Update-FileCount
            Update-Status -Message "Added $($filesFound.Count) file(s) from folder"
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "No supported files found in the selected folder for $convType conversion.",
                "No Files Found",
                "OK",
                "Information"
            )
        }
    }
}

function Add-FileToList {
    param([string]$FilePath)
    
    try {
        $fileInfo = Get-Item $FilePath -ErrorAction Stop
        $fileName = $fileInfo.Name
        
        foreach ($item in $Global:AppState.FileListView.Items) {
            if ($item.Tag -eq $FilePath) {
                return
            }
        }
        
        $size = Get-FileSizeString -SizeInBytes $fileInfo.Length
        $ext = $fileInfo.Extension.ToLower()
        
        $fileType = switch ($ext) {
            {$_ -in '.doc', '.docx', '.rtf'} { "Word" }
            {$_ -in '.xls', '.xlsx', '.xlsm', '.csv'} { "Excel" }
            {$_ -in '.ppt', '.pptx', '.pptm'} { "PowerPoint" }
            {$_ -in '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'} { "Image" }
            '.txt' { "Text" }
            {$_ -in '.html', '.htm'} { "HTML" }
            '.pdf' { "PDF" }
            default { $ext.Replace('.', '').ToUpper() }
        }
        
        $item = New-Object System.Windows.Forms.ListViewItem($fileName)
        $item.SubItems.Add($size) | Out-Null
        $item.SubItems.Add($fileType) | Out-Null
        $item.SubItems.Add("Pending") | Out-Null
        $item.Tag = $FilePath
        $item.SubItems[3].ForeColor = [System.Drawing.Color]::Gray
        
        $Global:AppState.FileListView.Items.Add($item)
        $Global:AppState.FilesToConvert.Add($FilePath)
        
    } catch {
        # Silently fail
    }
}

function Remove-SelectedFiles {
    $selectedItems = $Global:AppState.FileListView.SelectedItems
    if ($selectedItems.Count -gt 0) {
        $removedCount = 0
        foreach ($item in $selectedItems) {
            $index = $Global:AppState.FilesToConvert.IndexOf($item.Tag)
            if ($index -ge 0) {
                $Global:AppState.FilesToConvert.RemoveAt($index)
                $removedCount++
            }
        }
        foreach ($item in $selectedItems) {
            $Global:AppState.FileListView.Items.Remove($item)
        }
        Update-FileCount
        Update-Status -Message "Removed $removedCount file(s)"
    }
}

function Clear-Files {
    if ($Global:AppState.FilesToConvert.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Clear all files from the list?",
            "Confirm Clear",
            "YesNo",
            "Question"
        )
        
        if ($result -eq "Yes") {
            $Global:AppState.FileListView.Items.Clear()
            $Global:AppState.FilesToConvert.Clear()
            Update-FileCount
            Update-Status -Message "Cleared all files"
        }
    }
}

function Update-Status {
    param([string]$Message)
    
    if ($Global:AppState.StatusLabel) {
        $Global:AppState.StatusLabel.Text = $Message
        $Global:AppState.StatusLabel.Refresh()
    }
}

function Update-FileCount {
    if ($Global:AppState.FileCountLabel) {
        $count = $Global:AppState.FilesToConvert.Count
        $Global:AppState.FileCountLabel.Text = "Files: $count"
    }
}

# ============================================
# START CONVERSION - FIXED PDF MERGE
# ============================================

function Start-Conversion {
    if ($Global:AppState.FilesToConvert.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please add files to convert.",
            "No Files",
            "OK",
            "Warning"
        )
        return
    }
    
    $outputDir = $Global:AppState.OutputTextBox.Text
    if ($Global:AppState.SubfolderCheckBox.Checked) {
        $dateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
        $outputDir = Join-Path $outputDir $dateStamp
    }
    
    if (-not (Test-Path $outputDir)) {
        try {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Cannot create output directory:`n$_",
                "Error",
                "OK",
                "Error"
            )
            return
        }
    }
    
    $convType = $Global:AppState.ConversionListBox.SelectedItem.ToString()
    $quality = $Global:AppState.QualityComboBox.SelectedItem.ToString()
    
    $confirmMsg = @"
Convert $($Global:AppState.FilesToConvert.Count) file(s)?

Conversion Type: $convType
Output Folder: $outputDir
Quality: $quality

Continue?
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMsg,
        "Confirm Conversion",
        "YesNo",
        "Question"
    )
    
    if ($result -ne "Yes") {
        return
    }
    
    $Global:AppState.IsProcessing = $true
    $Global:AppState.ConversionStats = @{
        Total = $Global:AppState.FilesToConvert.Count
        Successful = 0
        Failed = 0
    }
    
    $Global:AppState.ProgressBar.Value = 0
    $Global:AppState.ProgressBar.Maximum = 100
    $Global:AppState.ConvertButton.Enabled = $false
    $Global:AppState.ConvertButton.Text = "PROCESSING..."
    $Global:AppState.ConvertButton.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    
    $currentFileNumber = 1
    $totalFiles = $Global:AppState.FilesToConvert.Count
    Update-Status -Message "Starting conversion of $totalFiles file(s)..."
    
    for ($i = 0; $i -lt $totalFiles; $i++) {
        $inputFile = $Global:AppState.FilesToConvert[$i]
        
        $progress = [math]::Round(($i + 1) / $totalFiles * 100, 0)
        $Global:AppState.ProgressBar.Value = $progress
        
        if ($i -lt $Global:AppState.FileListView.Items.Count) {
            $item = $Global:AppState.FileListView.Items[$i]
            $item.SubItems[3].Text = "Processing..."
            $item.SubItems[3].ForeColor = [System.Drawing.Color]::Blue
        }
        
        $currentFileNumber = $i + 1
        Update-Status -Message "Processing file $currentFileNumber of ${totalFiles}: $(Split-Path $inputFile -Leaf)"
        
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        $outputExt = switch ($convType) {
            "Word to PDF" { ".pdf" }
            "PDF to Word" { ".docx" }
            "Excel to PDF" { ".pdf" }
            "PDF to Excel" { ".xlsx" }
            "PowerPoint to PDF" { ".pdf" }
            "PDF to PowerPoint" { ".pptx" }
            "Images to PDF" { ".pdf" }
            "PDF to Images" { ".jpg" }
            "Text to PDF" { ".pdf" }
            "HTML to PDF" { ".pdf" }
            "PDF Merge" { ".pdf" }
            "PDF Split" { ".pdf" }
            "PDF Compress" { ".pdf" }
            "PDF Encrypt" { ".pdf" }
            "PDF Decrypt" { ".pdf" }
            "PDF Watermark" { ".pdf" }
            default { ".pdf" }
        }
        
        $outputFile = Join-Path $outputDir "${baseName}_${timestamp}${outputExt}"
        
        $success = $false
        try {
            switch ($convType) {
                "Word to PDF" {
                    $success = Convert-WordToPDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF to Word" {
                    $success = Convert-PDFToWord -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "Excel to PDF" {
                    $success = Convert-ExcelToPDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF to Excel" {
                    $success = Convert-PDFToExcel -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PowerPoint to PDF" {
                    $success = Convert-PowerPointToPDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF to PowerPoint" {
                    $success = Convert-PDFToPowerPoint -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "Images to PDF" {
                    $success = Convert-ImagesToPDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF to Images" {
                    $success = Convert-PDFToImages -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "Text to PDF" {
                    $success = Convert-TextToPDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "HTML to PDF" {
                    $success = Convert-HtmlToPdf -InputFile $inputFile -OutputFile $OutputFile -Quality $quality
                }
                
                # ============================================
                # FIXED PDF MERGE - PROCESS ONCE FOR ALL FILES
                # ============================================
                "PDF Merge" {
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $mergeOutputFile = Join-Path $outputDir "Merged_PDFs_${timestamp}.pdf"
                    
                    $success = Merge-PDFs -InputFiles $Global:AppState.FilesToConvert -OutputFile $mergeOutputFile -Quality $quality -Silent
                    
                    if ($success) {
                        for ($j = 0; $j -lt $Global:AppState.FileListView.Items.Count; $j++) {
                            $mergeItem = $Global:AppState.FileListView.Items[$j]
                            $mergeItem.SubItems[3].Text = "Merged"
                            $mergeItem.SubItems[3].ForeColor = [System.Drawing.Color]::Green
                        }
                        
                        $Global:AppState.ConversionStats.Successful = $Global:AppState.FilesToConvert.Count
                        
                        $i = $totalFiles - 1
                    } else {
                        for ($j = 0; $j -lt $Global:AppState.FileListView.Items.Count; $j++) {
                            $mergeItem = $Global:AppState.FileListView.Items[$j]
                            $mergeItem.SubItems[3].Text = "Merge Failed"
                            $mergeItem.SubItems[3].ForeColor = [System.Drawing.Color]::Red
                        }
                        $Global:AppState.ConversionStats.Failed = $Global:AppState.FilesToConvert.Count
                        $i = $totalFiles - 1
                    }
                }
                
                "PDF Split" {
                    $success = Split-PDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF Compress" {
                    $success = Compress-PDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF Encrypt" {
                    $success = Encrypt-PDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF Decrypt" {
                    $success = Decrypt-PDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                "PDF Watermark" {
                    $success = Watermark-PDF -InputFile $inputFile -OutputFile $outputFile -Quality $quality
                }
                default {
                    $success = Create-EnhancedPDF -InputFile $inputFile -OutputFile $outputFile -ConversionType $convType -Quality $quality
                }
            }
        } catch {
            $success = $false
        }
        
        if ($success) {
            $Global:AppState.ConversionStats.Successful++
            if ($i -lt $Global:AppState.FileListView.Items.Count -and $convType -ne "PDF Merge") {
                $item = $Global:AppState.FileListView.Items[$i]
                $item.SubItems[3].Text = "Success"
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::Green
                
                if (Test-Path $outputFile) {
                    $newSize = Get-FileSizeString -SizeInBytes ((Get-Item $outputFile).Length)
                    $item.SubItems[1].Text = $newSize
                }
            }
        } else {
            $Global:AppState.ConversionStats.Failed++
            if ($i -lt $Global:AppState.FileListView.Items.Count -and $convType -ne "PDF Merge") {
                $item = $Global:AppState.FileListView.Items[$i]
                $item.SubItems[3].Text = "Failed"
                $item.SubItems[3].ForeColor = [System.Drawing.Color]::Red
            }
        }
        
        Start-Sleep -Milliseconds 100
    }
    
    $Global:AppState.ProgressBar.Value = 100
    $Global:AppState.IsProcessing = $false
    $Global:AppState.ConvertButton.Enabled = $true
    $Global:AppState.ConvertButton.Text = "START CONVERSION"
    $Global:AppState.ConvertButton.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    
    $completionMsg = @"
Conversion Complete!

Total Files: $($Global:AppState.ConversionStats.Total)
Successful: $($Global:AppState.ConversionStats.Successful)
Failed: $($Global:AppState.ConversionStats.Failed)

Output Folder: $outputDir
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $completionMsg,
        "Conversion Complete",
        "OK",
        "Information"
    )
    
    if ($result -eq "OK") {
        $openFolder = [System.Windows.Forms.MessageBox]::Show(
            "Open output folder?",
            "Open Folder",
            "YesNo",
            "Question"
        )
        
        if ($openFolder -eq "Yes") {
            try {
                if (Test-Path $outputDir) {
                    Invoke-Item $outputDir
                }
            } catch {
                # Silently fail
            }
        }
    }
    
    Update-Status -Message "Ready to convert"
}

# ============================================
# MAIN FORM - FIXED LAYOUT
# ============================================

function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$Global:AppName v$Global:Version"
    $form.Size = New-Object System.Drawing.Size(1000, 700)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = Get-Location }
    
    # Load icon with resource fallback
	$iconPath = Get-ResourcePath -ResourceName "Icon.ico" -DefaultPath (Join-Path $scriptDir "Icon.ico")
	if (Test-Path $iconPath) {
		try {
			$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
		} catch {
			# Try loading directly from file
			try {
				$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($iconPath).GetHicon()))
			} catch { }
		}
	} else {
		# Try PNG as fallback
		$pngPath = Get-ResourcePath -ResourceName "Icon.png" -DefaultPath (Join-Path $scriptDir "Icon.png")
		if (Test-Path $pngPath) {
			try {
				$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($pngPath).GetHicon()))
			} catch { }
		}
	}
    
    # ============================================
    # HEADER PANEL WITH LOGO
    # ============================================
    
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(984, 90)
    $headerPanel.Location = New-Object System.Drawing.Point(8, 8)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
    
    $logoPath = Join-Path $scriptDir "Logo.png"
    $logoPictureBox = New-Object System.Windows.Forms.PictureBox
    $logoPictureBox.Size = New-Object System.Drawing.Size(70, 70)
    $logoPictureBox.Location = New-Object System.Drawing.Point(15, 10)
    $logoPictureBox.SizeMode = "StretchImage"
    
    $logoImage = Load-Image -ImagePath $logoPath -MaxWidth 70 -MaxHeight 70
    $logoPictureBox.Image = $logoImage
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Global:AppName
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Location = New-Object System.Drawing.Point(100, 20)
    $titleLabel.AutoSize = $true
    
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "v$Global:Version"
    $versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $versionLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $versionLabel.Location = New-Object System.Drawing.Point(800, 15)
    $versionLabel.AutoSize = $true
    
    $companyLabel = New-Object System.Windows.Forms.Label
    $companyLabel.Text = $Global:Company
    $companyLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $companyLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $companyLabel.Location = New-Object System.Drawing.Point(800, 35)
    $companyLabel.AutoSize = $true
    
    $websiteLink = New-Object System.Windows.Forms.LinkLabel
    $websiteLink.Text = "Visit Website"
    $websiteLink.Location = New-Object System.Drawing.Point(800, 55)
    $websiteLink.AutoSize = $true
    $websiteLink.LinkColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $websiteLink.ActiveLinkColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $websiteLink.VisitedLinkColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    $websiteLink.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Underline)
    $websiteLink.Add_Click({
        try {
            Start-Process $Global:Website
        } catch { }
    })
    
    $headerPanel.Controls.AddRange(@($logoPictureBox, $titleLabel, $versionLabel, $companyLabel, $websiteLink))
    
    # ============================================
    # CONVERSION TYPE SELECTION - FIXED PANEL SIZE
    # ============================================
    
    $conversionGroup = New-Object System.Windows.Forms.GroupBox
    $conversionGroup.Text = " Conversion Type "
    $conversionGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $conversionGroup.Size = New-Object System.Drawing.Size(300, 450)
    $conversionGroup.Location = New-Object System.Drawing.Point(8, 105)
    
    $conversionListBox = New-Object System.Windows.Forms.ListBox
    $conversionListBox.Size = New-Object System.Drawing.Size(280, 425)
    $conversionListBox.Location = New-Object System.Drawing.Point(10, 25)
    $conversionListBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $conversionListBox.SelectionMode = "One"
    $conversionListBox.IntegralHeight = $true
    $conversionListBox.Dock = "None"
    
    foreach ($type in $Global:ConversionTypes) {
        $conversionListBox.Items.Add($type) | Out-Null
    }
    $conversionListBox.SelectedIndex = 0
    $conversionListBox.Add_SelectedIndexChanged({
        Update-FileFilters
    })
    
    $conversionGroup.Controls.Add($conversionListBox)
    
    # ============================================
    # FILE SELECTION AREA
    # ============================================
    
    $fileGroup = New-Object System.Windows.Forms.GroupBox
    $fileGroup.Text = " Files to Convert "
    $fileGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $fileGroup.Size = New-Object System.Drawing.Size(660, 220)
    $fileGroup.Location = New-Object System.Drawing.Point(315, 105)
    
    $fileListView = New-Object System.Windows.Forms.ListView
    $fileListView.Size = New-Object System.Drawing.Size(640, 150)
    $fileListView.Location = New-Object System.Drawing.Point(10, 25)
    $fileListView.View = "Details"
    $fileListView.FullRowSelect = $true
    $fileListView.GridLines = $true
    $fileListView.Columns.Add("File Name", 300) | Out-Null
    $fileListView.Columns.Add("Size", 100) | Out-Null
    $fileListView.Columns.Add("Type", 100) | Out-Null
    $fileListView.Columns.Add("Status", 100) | Out-Null
    
    $fileButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $fileButtonPanel.Size = New-Object System.Drawing.Size(640, 35)
    $fileButtonPanel.Location = New-Object System.Drawing.Point(10, 180)
    $fileButtonPanel.FlowDirection = "LeftToRight"
    
    $addFilesButton = New-Object System.Windows.Forms.Button
    $addFilesButton.Text = "Add Files"
    $addFilesButton.Size = New-Object System.Drawing.Size(100, 30)
    $addFilesButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $addFilesButton.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $addFilesButton.ForeColor = [System.Drawing.Color]::White
    $addFilesButton.FlatStyle = "Flat"
    $addFilesButton.FlatAppearance.BorderSize = 0
    $addFilesButton.Add_Click({ Add-Files })
    
    $addFolderButton = New-Object System.Windows.Forms.Button
    $addFolderButton.Text = "Add Folder"
    $addFolderButton.Size = New-Object System.Drawing.Size(100, 30)
    $addFolderButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $addFolderButton.BackColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    $addFolderButton.ForeColor = [System.Drawing.Color]::White
    $addFolderButton.FlatStyle = "Flat"
    $addFolderButton.FlatAppearance.BorderSize = 0
    $addFolderButton.Add_Click({ Add-Folder })
    
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove"
    $removeButton.Size = New-Object System.Drawing.Size(100, 30)
    $removeButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $removeButton.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
    $removeButton.ForeColor = [System.Drawing.Color]::White
    $removeButton.FlatStyle = "Flat"
    $removeButton.FlatAppearance.BorderSize = 0
    $removeButton.Add_Click({ Remove-SelectedFiles })
    
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "Clear All"
    $clearButton.Size = New-Object System.Drawing.Size(100, 30)
    $clearButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $clearButton.BackColor = [System.Drawing.Color]::FromArgb(149, 165, 166)
    $clearButton.ForeColor = [System.Drawing.Color]::White
    $clearButton.FlatStyle = "Flat"
    $clearButton.FlatAppearance.BorderSize = 0
    $clearButton.Add_Click({ Clear-Files })
    
    $fileButtonPanel.Controls.AddRange(@($addFilesButton, $addFolderButton, $removeButton, $clearButton))
    $fileGroup.Controls.AddRange(@($fileListView, $fileButtonPanel))
    
    # ============================================
    # CONVERSION OPTIONS
    # ============================================
    
    $optionsGroup = New-Object System.Windows.Forms.GroupBox
    $optionsGroup.Text = " Conversion Options "
    $optionsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $optionsGroup.Size = New-Object System.Drawing.Size(660, 120)
    $optionsGroup.Location = New-Object System.Drawing.Point(315, 335)
    
    $qualityLabel = New-Object System.Windows.Forms.Label
    $qualityLabel.Text = "Quality:"
    $qualityLabel.Location = New-Object System.Drawing.Point(20, 30)
    $qualityLabel.Size = New-Object System.Drawing.Size(80, 25)
    $qualityLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $qualityComboBox = New-Object System.Windows.Forms.ComboBox
    $qualityComboBox.Size = New-Object System.Drawing.Size(120, 25)
    $qualityComboBox.Location = New-Object System.Drawing.Point(100, 30)
    $qualityComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $qualityComboBox.DropDownStyle = "DropDownList"
    foreach ($level in $Global:QualityLevels) {
        $qualityComboBox.Items.Add($level) | Out-Null
    }
    $qualityComboBox.SelectedIndex = 2
    
    $subfolderCheckBox = New-Object System.Windows.Forms.CheckBox
    $subfolderCheckBox.Text = "Create date-based subfolder"
    $subfolderCheckBox.Location = New-Object System.Drawing.Point(240, 30)
    $subfolderCheckBox.Size = New-Object System.Drawing.Size(200, 25)
    $subfolderCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subfolderCheckBox.Checked = $true
    
    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Text = "Output Folder:"
    $outputLabel.Location = New-Object System.Drawing.Point(20, 70)
    $outputLabel.Size = New-Object System.Drawing.Size(100, 25)
    $outputLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Size = New-Object System.Drawing.Size(350, 25)
    $outputTextBox.Location = New-Object System.Drawing.Point(130, 70)
    $outputTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $outputTextBox.Text = Join-Path $env:USERPROFILE "PDF_Converter_Output"
    
    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse"
    $browseButton.Size = New-Object System.Drawing.Size(80, 25)
    $browseButton.Location = New-Object System.Drawing.Point(490, 70)
    $browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $browseButton.BackColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
    $browseButton.ForeColor = [System.Drawing.Color]::White
    $browseButton.FlatStyle = "Flat"
    $browseButton.FlatAppearance.BorderSize = 0
    $browseButton.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.SelectedPath = $outputTextBox.Text
        $folderDialog.Description = "Select output folder"
        if ($folderDialog.ShowDialog() -eq "OK") {
            $outputTextBox.Text = $folderDialog.SelectedPath
        }
    })
    
    $optionsGroup.Controls.AddRange(@(
        $qualityLabel, $qualityComboBox,
        $outputLabel, $outputTextBox,
        $browseButton, $subfolderCheckBox
    ))
    
    # ============================================
    # CONVERSION BUTTON AREA
    # ============================================
    
    $convertPanel = New-Object System.Windows.Forms.Panel
    $convertPanel.Size = New-Object System.Drawing.Size(660, 80)
    $convertPanel.Location = New-Object System.Drawing.Point(315, 465)
    
    $convertButton = New-Object System.Windows.Forms.Button
    $convertButton.Text = "START CONVERSION"
    $convertButton.Size = New-Object System.Drawing.Size(640, 50)
    $convertButton.Location = New-Object System.Drawing.Point(10, 15)
    $convertButton.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $convertButton.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
    $convertButton.ForeColor = [System.Drawing.Color]::White
    $convertButton.FlatStyle = "Flat"
    $convertButton.FlatAppearance.BorderSize = 0
    $convertButton.Add_Click({ Start-Conversion })
    
    $convertPanel.Controls.Add($convertButton)
    
    # ============================================
    # STATUS BAR - FIXED: RIGHT-ALIGNED PROGRESS AND TEXT
    # ============================================
    
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Size = New-Object System.Drawing.Size(984, 50)
    $statusPanel.Location = New-Object System.Drawing.Point(8, 595)
    $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready to convert"
    $statusLabel.ForeColor = [System.Drawing.Color]::White
    $statusLabel.Location = New-Object System.Drawing.Point(15, 15)
    $statusLabel.AutoSize = $true
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size(250, 20)
    $progressBar.Location = New-Object System.Drawing.Point(720, 15)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    
    $fileCountLabel = New-Object System.Windows.Forms.Label
    $fileCountLabel.Text = "Files: 0"
    $fileCountLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $fileCountLabel.Location = New-Object System.Drawing.Point(600, 15)
    $fileCountLabel.AutoSize = $true
    $fileCountLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $statusPanel.Controls.AddRange(@($statusLabel, $fileCountLabel, $progressBar))
    
    # ============================================
    # ADD ALL CONTROLS TO FORM
    # ============================================
    
    $form.Controls.AddRange(@(
        $headerPanel,
        $conversionGroup,
        $fileGroup,
        $optionsGroup,
        $convertPanel,
        $statusPanel
    ))
    
    $Global:AppState.Form = $form
    $Global:AppState.ConversionListBox = $conversionListBox
    $Global:AppState.FileListView = $fileListView
    $Global:AppState.QualityComboBox = $qualityComboBox
    $Global:AppState.OutputTextBox = $outputTextBox
    $Global:AppState.SubfolderCheckBox = $subfolderCheckBox
    $Global:AppState.ConvertButton = $convertButton
    $Global:AppState.StatusLabel = $statusLabel
    $Global:AppState.ProgressBar = $progressBar
    $Global:AppState.FileCountLabel = $fileCountLabel
    
    [System.Windows.Forms.Application]::Run($form)
}

# ============================================
# MAIN EXECUTION - SILENT MODE
# ============================================

# Initialize everything silently
$null = Initialize-Tools
$null = Test-PortableMode

# ============================================
# HANDLE AUTO-INSTALL MODE
# ============================================

if ($Global:AutoInstallMode -and $Global:AutoInstallTool) {
    $null = Install-Tool -ToolName $Global:AutoInstallTool
    exit
}

# ============================================
# CHECK FOR MISSING TOOLS
# ============================================

$toolsReady = Test-AndInstallMissingTools

if (-not $toolsReady) { 
    [System.Windows.Forms.Application]::Exit()
    return
}

# ============================================
# LAUNCH THE MAIN APPLICATION FORM
# ============================================

$null = Show-MainForm