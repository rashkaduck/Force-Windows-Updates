<#
.SYNOPSIS
  Unattended Windows Update script that never prompts the user.

.DESCRIPTION
  - Runs completely without prompts or questions
  - All messages in English
  - Uses built-in Windows Update methods
  - No external modules needed
  - Installs all updates automatically
#>

param(
    [bool]$UseMicrosoftUpdate = $true,
    [bool]$InstallUpdates = $true,
    [bool]$AutoReboot = $true,
    [string]$LogFile = "C:\Temp\WindowsUpdate.log",
    [switch]$Silent
)

# ----------------------------
# Global unattended configuration
# ----------------------------
$ErrorActionPreference = 'Stop'
$ConfirmPreference = 'None'
if ($Silent) {
    $VerbosePreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
} else {
    $VerbosePreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'
}

# Create temp folder if it doesn't exist
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null
}

# ----------------------------
# Simple logging function
# ----------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$Level] $Message"
    
    # Write to file
    $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    
    # Write to console if not silent mode
    if (-not $Silent) {
        switch ($Level) {
            'ERROR'   { Write-Host $line -ForegroundColor Red }
            'WARNING' { Write-Host $line -ForegroundColor Yellow }
            default   { Write-Host $line }
        }
    }
}

# ----------------------------
# Main function to run Windows Update
# ----------------------------
function Run-WindowsUpdate {
    Write-Log "==== STARTING WINDOWS UPDATE PROCESS ====" "INFO"
    
    try {
        # 1. Restart Windows Update services
        Write-Log "Restarting Windows Update services..." "INFO"
        
        $services = @('wuauserv', 'BITS', 'cryptsvc')
        
        foreach ($service in $services) {
            try {
                $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
                if ($svc.Status -eq 'Running') {
                    Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
                }
                Start-Sleep -Seconds 2
                Start-Service -Name $service -ErrorAction SilentlyContinue
                Write-Log "Restarted service: $service" "DEBUG"
            } catch {
                Write-Log "Could not restart $service" "WARNING"
            }
        }
        
        # 2. Clear Windows Update cache
        Write-Log "Clearing Windows Update cache..." "INFO"
        
        try {
            # Stop Windows Update service first
            Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
            
            # Clear cache folders
            $cacheFolders = @(
                "$env:SystemRoot\SoftwareDistribution\Download",
                "$env:SystemRoot\SoftwareDistribution\DataStore"
            )
            
            foreach ($folder in $cacheFolders) {
                if (Test-Path $folder) {
                    Get-ChildItem -Path $folder -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                    Write-Log "Cleared: $folder" "DEBUG"
                }
            }
            
            # Start service again
            Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            
        } catch {
            Write-Log "Could not clear cache" "WARNING"
        }
        
        # 3. Use UsoClient to search for updates
        Write-Log "Searching for updates with UsoClient..." "INFO"
        
        try {
            Start-Process "UsoClient" -ArgumentList "StartInteractiveScan" -Wait -WindowStyle Hidden
            Write-Log "UsoClient scan completed" "DEBUG"
        } catch {
            Write-Log "Could not run UsoClient" "WARNING"
        }
        
        # 4. Search for updates with COM object
        Write-Log "Checking for available updates..." "INFO"
        
        $foundUpdates = $false
        $updatesCount = 0
        
        try {
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $searchResult = $searcher.Search("IsInstalled=0 and Type='Software'")
            
            $updatesCount = $searchResult.Updates.Count
            
            if ($updatesCount -gt 0) {
                $foundUpdates = $true
                Write-Log "Found $updatesCount update(s)" "INFO"
                
                # Log first 5 updates
                for ($i = 0; $i -lt [Math]::Min(5, $updatesCount); $i++) {
                    $update = $searchResult.Updates.Item($i)
                    $title = $update.Title
                    if ($title.Length -gt 80) { 
                        $title = $title.Substring(0, 77) + "..." 
                    }
                    Write-Log "  $($i + 1). $title" "INFO"
                }
                
                if ($updatesCount -gt 5) {
                    Write-Log "  ... and $($updatesCount - 5) more updates" "INFO"
                }
            } else {
                Write-Log "No updates found" "INFO"
            }
            
        } catch {
            Write-Log "Error searching for updates: $($_.Exception.Message)" "ERROR"
        }
        
        # 5. Install updates if requested and found
        if ($InstallUpdates -and $foundUpdates) {
            Write-Log "Preparing to install updates..." "INFO"
            
            try {
                # Create collection of all updates to install
                $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                
                for ($i = 0; $i -lt $searchResult.Updates.Count; $i++) {
                    $update = $searchResult.Updates.Item($i)
                    $updatesToInstall.Add($update) | Out-Null
                }
                
                # Create downloader and download updates
                Write-Log "Downloading updates..." "INFO"
                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $updatesToInstall
                $downloadResult = $downloader.Download()
                
                if ($downloadResult.ResultCode -eq 2) { # Downloaded successfully
                    Write-Log "All updates downloaded" "INFO"
                    
                    # Create installer and install
                    Write-Log "Installing updates..." "INFO"
                    $installer = $session.CreateUpdateInstaller()
                    $installer.Updates = $updatesToInstall
                    $installer.ForceQuiet = $true  # NO prompts
                    $installer.IsForced = $true    # Force installation
                    
                    $installationResult = $installer.Install()
                    
                    if ($installationResult.ResultCode -eq 2) { # Installed successfully
                        Write-Log "All updates installed successfully" "INFO"
                        
                        # Check if reboot is required
                        if ($installationResult.RebootRequired) {
                            Write-Log "Reboot required after installation" "INFO"
                            
                            if ($AutoReboot) {
                                Write-Log "Rebooting system in 60 seconds..." "INFO"
                                # Reboot in 60 seconds, force all applications to close
                                shutdown /r /t 60 /c "Windows Update installation complete" /f
                            }
                        }
                        
                        return $true
                    } else {
                        Write-Log "Installation failed with code: $($installationResult.ResultCode)" "ERROR"
                    }
                } else {
                    Write-Log "Download failed with code: $($downloadResult.ResultCode)" "ERROR"
                }
                
            } catch {
                Write-Log "Error during installation: $($_.Exception.Message)" "ERROR"
            }
        } elseif (-not $foundUpdates) {
            Write-Log "No updates to install" "INFO"
            return $false
        } else {
            Write-Log "InstallUpdates is disabled, skipping installation" "INFO"
            return $false
        }
        
    } catch {
        Write-Log "Serious error in Windows Update process: $($_.Exception.Message)" "ERROR"
        return $false
    }
    
    return $false
}

# ----------------------------
# RUN MAIN PROGRAM
# ----------------------------
try {
    # Set execution policy for this process
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
    
    # Run Windows Update
    $result = Run-WindowsUpdate
    
    if ($result) {
        Write-Log "Windows Update process completed" "INFO"
    } else {
        Write-Log "No updates could be found or installed" "INFO"
    }
    
} catch {
    Write-Log "Critical error in main script: $($_.Exception.Message)" "ERROR"
} finally {
    Write-Log "==== SCRIPT FINISHED ====" "INFO"
}