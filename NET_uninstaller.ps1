<#
.SYNOPSIS
    Uninstalls all .NET Desktop Runtime versions 5.0 through 7.0 (x86 and x64)
.DESCRIPTION
    Queries the registry for installed .NET Desktop Runtimes and uninstalls them using 
    the actual uninstaller executables from the Package Cache.
.NOTES
    Author: SCCM Task Sequence Script
    Date: 2025-10-29
    Requires: PowerShell 5.1+, Administrative privileges
#>

# Set error action preference
$ErrorActionPreference = "Continue"

# Log function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    
    $logPath = "C:\DotNetUninstalls\DotNetUninstall.log"
    Add-Content -Path $logPath -Value $logMessage -ErrorAction SilentlyContinue
}

Write-Log "Starting .NET Desktop Runtime uninstallation process"

# Get all installed .NET Desktop Runtimes from registry
$runtimes = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", `
                               "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
    Where-Object { 
        $_.DisplayName -like "Microsoft Windows Desktop Runtime - 5.*" -or 
        $_.DisplayName -like "Microsoft Windows Desktop Runtime - 6.*" -or 
        $_.DisplayName -like "Microsoft Windows Desktop Runtime - 7.*" 
    }

$foundCount = ($runtimes | Measure-Object).Count
$uninstalledCount = 0
$failedCount = 0
$notFoundCount = 0

Write-Log "Found $foundCount .NET Desktop Runtime(s) to uninstall"

# Function to find the uninstaller executable in Package Cache
function Find-UninstallerExe {
    param([string]$GUID)
    
    $packageCachePath = "C:\ProgramData\Package Cache\$GUID"
    
    if (Test-Path $packageCachePath) {
        # Look for windowsdesktop-runtime exe files
        $exeFiles = Get-ChildItem -Path $packageCachePath -Filter "windowsdesktop-runtime*.exe" -ErrorAction SilentlyContinue
        
        if ($exeFiles) {
            return $exeFiles[0].FullName
        }
    }
    
    return $null
}

# Uninstall each runtime
foreach ($runtime in $runtimes) {
    $name = $runtime.DisplayName
    $version = $runtime.DisplayVersion
    $guid = $runtime.PSChildName
    
    Write-Log "Processing: $name (Version: $version)" "INFO"
    Write-Log "GUID: $guid" "INFO"
    
    # Find the uninstaller executable
    $uninstallerPath = Find-UninstallerExe -GUID $guid
    
    if ($uninstallerPath) {
        Write-Log "Found uninstaller: $uninstallerPath" "INFO"
        
        try {
            # Use the actual uninstaller with /uninstall /quiet flags
            $process = Start-Process -FilePath $uninstallerPath -ArgumentList "/uninstall /quiet" -Wait -PassThru -NoNewWindow
            
            if ($process.ExitCode -eq 0) {
                Write-Log "Successfully uninstalled: $name" "INFO"
                $uninstalledCount++
            }
            elseif ($process.ExitCode -eq 3010) {
                Write-Log "Successfully uninstalled (reboot required): $name" "INFO"
                $uninstalledCount++
            }
            else {
                Write-Log "Uninstall failed with exit code $($process.ExitCode): $name" "ERROR"
                $failedCount++
            }
            
            # Small delay between uninstalls
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Log "Exception during uninstall of $name : $_" "ERROR"
            $failedCount++
        }
    }
    else {
        Write-Log "WARNING: Uninstaller executable not found in Package Cache for $name" "WARN"
        Write-Log "Expected path: C:\ProgramData\Package Cache\$guid" "WARN"
        $notFoundCount++
    }
}

# Summary
Write-Log "================================================" "INFO"
Write-Log "Uninstallation Summary:" "INFO"
Write-Log "  Total runtimes found: $foundCount" "INFO"
Write-Log "  Successfully uninstalled: $uninstalledCount" "INFO"
Write-Log "  Failed to uninstall: $failedCount" "INFO"
Write-Log "  Uninstaller not found: $notFoundCount" "INFO"
Write-Log "================================================" "INFO"

# Exit with appropriate code
if ($failedCount -gt 0) {
    Write-Log "Completed with errors - some uninstalls failed" "ERROR"
    exit 1
}
elseif ($uninstalledCount -gt 0) {
    if ($notFoundCount -gt 0) {
        Write-Log "Completed successfully - $notFoundCount duplicate/orphaned registry entries were skipped" "INFO"
    }
    else {
        Write-Log "All found runtimes uninstalled successfully" "INFO"
    }
    exit 0
}
else {
    if ($notFoundCount -gt 0) {
        Write-Log "No valid uninstallers found - only orphaned registry entries detected" "WARN"
        exit 0
    }
    else {
        Write-Log "No .NET Desktop Runtimes found to uninstall" "INFO"
        exit 0
    }
}