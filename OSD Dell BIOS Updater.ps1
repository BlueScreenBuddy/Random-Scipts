#Initialize Task Sequence Environment
$TSEnv=New-Object -COMObject Microsoft.SMS.TSEnvironment
$SMSTSLogPath=$TSEnv.Value("_SMSTSLogPath")
$OSDTargetSystemDrive=$TSEnv.Value("OSDTargetSystemDrive")

#Disable Progress bars to speed up Invoke-WebRequest
$ProgressPreference = 'SilentlyContinue'

#Define Log file path
$Script:Logfile="$($SMSTSLogPath)\Update-BIOS.log"

#Function to write a log entry, with handling for file locks
Function Write-LogEntry
{
    Param($Value)
    $Tries=0
    $Success=$False
    Do {
        Try{
            if ($Value-ne " "){
                $Timestamp="[$((Get-Date).ToString())] "
            } else {
                $Timestamp=""
            }
            Add-Content -Value "$($Timestamp)$($Value)" -Path $Script:LogFile -Force -ErrorAction Stop
            $Success=$True
        }
        Catch{
            $Success=$False
            $Tries++
            Start-Sleep -Seconds 1
        }  
    }
    Until ($Success -eq $True -or $Tries -ge 5)
}

#Function to delete files created by the script
Function Cleanup-Temp
{
    Write-LogEntry -Value " "
    Write-LogEntry -Value "Cleaning up Temp files"
    Try{
        Remove-Item "$($TempFolder)" -Recurse -Force -ErrorAction Stop
        Write-LogEntry -Value "Successfully cleaned up Temp files"
    }
    Catch {
        Write-LogEntry -Value "Failed to clean up Temp files, you will need to manually remove $($TempFolder)"
    }
}

#Function to clean up and provide exit code before exiting script
Function Exit-Script
{
    Param($ExitCode)
    Cleanup-Temp
    Write-LogEntry -Value " "
    Write-LogEntry -Value "Script Completed with exit code $($ExitCode)"
    Exit $ExitCode
}

#Create Log File
If (Test-Path $LogFile){
    Remove-Item $LogFile -Force
}

Write-LogEntry -Value "Started BIOS Upgrade Script" 

#Create Temp Folder
Write-LogEntry -Value " " 
Write-LogEntry -Value "Creating BIOS folder in Temp" 
$TempFolder="$($env:TEMP)\BIOS"
If (!(Test-Path $TempFolder)){
    Try{
        New-Item -Path "$($env:Temp)" -Name "BIOS" -ItemType Directory -Force -ErrorAction Stop
        Write-LogEntry -Value "Successfully created $($TempFolder) folder" 
    }
    Catch{
        Write-LogEntry -Value "Failed to create $($TempFolder) folder"
        Exit-Script -ExitCode 100
    }
} else {
    Write-LogEntry -Value "$($TempFolder) folder already exists" 
}

#Get System SKU
Write-LogEntry -Value " " 
Write-LogEntry -Value "Getting System SKU value from WMI" 
Try {
    $Script:SystemSku="$((Get-WMIObject -Namespace root\WMI -Class MS_SystemInformation -ErrorAction Stop).SystemSku)"
    Write-LogEntry -Value "Retrieved System SKU value: $($Script:SystemSku)" 
}
Catch{
    Write-LogEntry -Value "Failed to retrieve System SKU value"
    Exit-Script -ExitCode 101
}

#Get System BIOS Version and convert to Version string
Write-LogEntry -Value " " 
Write-LogEntry -Value "Getting BIOS version" 
$Script:BIOSVersion=(Get-WMIObject -Class win32_bios -Property smbiosbiosversion).smbiosbiosversion
[string]$Script:BIOSPaddedVersion=$Script:BIOSVersion -replace "[^0-9.]", "" -replace "^0", ""
$Script:BIOSPaddedVersion += ".0" * (4 - $Script:BIOSPaddedVersion.Split('.').Count)
[version]$Script:BIOSPaddedVersion=$Script:BIOSPaddedVersion
Write-LogEntry -Value "Retrieved BIOS Version: $($Script:BIOSVersion)" 
Write-LogEntry -Value "Converted BIOS Version: $([string]$Script:BIOSPaddedVersion)" 

#Download BIOS Flash Tool
Write-LogEntry -Value " " 
Write-LogEntry -Value "Downloading BIOS Flash Tool" 
Try{
    $Script:FlashToolZIP="FlashVer3.3.28.zip"
    Invoke-WebRequest -Uri "https://downloads.dell.com/FOLDER12288556M/1/$($Script:FlashToolZIP)" -OutFile "$($TempFolder)\$($Script:FlashToolZIP)" -ErrorAction Stop
    Write-LogEntry -Value "Successfully downloaded https://downloads.dell.com/FOLDER12288556M/1/$($Script:FlashToolZIP)" 
}
Catch{
    Write-LogEntry -Value "Failed to download https://downloads.dell.com/FOLDER12288556M/1/$($Script:FlashToolZIP)"
    Exit-Script -ExitCode 102 
}

#Download Model Catalog XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Downloading Model Catalog XML" 
Try{
    $Script:CatalogCABName="CatalogIndexPC.cab"
    Invoke-WebRequest -Uri "https://downloads.dell.com/catalog/$Script:CatalogCABName" -OutFile "$($TempFolder)\$Script:CatalogCABName" -ErrorAction Stop
    Write-LogEntry -Value "Successfully downloaded https://downloads.dell.com/catalog/$Script:CatalogCABName" 
}
Catch{
    Write-LogEntry -Value "Failed to download https://downloads.dell.com/catalog/$Script:CatalogCABName"
    Exit-Script -ExitCode 103
}
      
#Extract BIOS Flash Tool
Write-LogEntry -Value " " 
Write-LogEntry -Value "Extracting BIOS Flash Tool" 
Try{
    Expand-Archive -Path "$($TempFolder)\$($Script:FlashToolZIP)" -DestinationPath "$($TempFolder)" -Force -ErrorAction Stop
    $Flash64W=(Get-ChildItem -Path $($TempFolder) -Recurse -ErrorAction Stop | Where-Object {$_.Name -eq "Flash64W.exe"}).FullName
    Move-Item -Path $Flash64W -Destination $($TempFolder) -Force -ErrorAction Stop
    Write-LogEntry -Value "Successfully extracted Flash64W.exe from $($Script:FlashToolZIP)" 
}
Catch{
    Write-LogEntry -Value "Failed to extract Flash64W.exe from $($Script:FlashToolZIP)"
    Exit-Script -ExitCode 104 
}

#Extract Catalog XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Extracting Model Catalog XML" 
Try{
    $Script:CatalogXMLName = $Script:CatalogCABName -Replace "cab","xml"
    Start-Process "$($env:windir)\system32\expand.exe" -WindowStyle Hidden -ArgumentList "`"$($TempFolder)\$($Script:CatalogCABName)`" `"$($TempFolder)\$($Script:CatalogXMLName)`"" -Wait -ErrorAction Stop
    Write-LogEntry -Value "Successfully extracted $($Script:CatalogXMLName) from $($Script:CatalogCABName)" 
}
Catch{
    Write-LogEntry -Value "Failed to extract $($Script:CatalogXMLName) from $($Script:CatalogCABName)"
    Exit-Script -ExitCode 105     
}

#Get Model URL from Catalog XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Get Model URL from Catalog XML" 
Try{
    [xml]$Script:ModelCatalog = Get-Content "$($TempFolder)\CatalogIndexPC.xml" -ErrorAction Stop
    $Script:ModelURL="https://$($ModelCatalog.ManifestIndex.baseLocation)/$(($Script:ModelCatalog.ManifestIndex.GroupManifest | Where-Object {$_.supportedsystems.brand.model.systemID -eq $Script:SystemSKU}).manifestinformation.path)"
    Write-LogEntry -Value "Successfully retrieved Model URL: $($Script:ModelURL)" 
}
Catch{
    Write-LogEntry -Value "Failed to retrieve Model URL"
    Exit-Script -ExitCode 106    
}

#Download Model XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Downloading Model Catalog XML" 
Try{
    $Script:BIOSCABName = Split-Path -Path $Script:ModelURL -Leaf
    Invoke-WebRequest -Uri $Script:ModelURL -OutFile "$($TempFolder)\$($Script:BIOSCABName)" -ErrorAction Stop
    Write-LogEntry -Value "Successfully downloaded $($Script:ModelURL)" 
}
Catch{
    Write-LogEntry -Value "Failed to download $($Script:ModelURL)"
    Exit-Script -ExitCode 107 
}

#Extract Model XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Extracting Model XML" 
Try{
    $Script:BIOSXMLName=$Script:BIOSCABName -Replace "cab","xml"
    Start-Process "$($env:windir)\system32\expand.exe" -WindowStyle Hidden -ArgumentList "`"$($TempFolder)\$($Script:BIOSCABName)`" `"$($TempFolder)\$($Script:BIOSXMLName)`"" -Wait -ErrorAction Stop
    Write-LogEntry -Value "Successfully extracted $Script:BIOSXMLName from $Script:BIOSCABName" 
}
Catch{
    Write-LogEntry -Value "Failed to extract $Script:BIOSXMLName from $Script:BIOSCABName"
    Exit-Script -ExitCode 108     
}

#Get Highest BIOS Version from XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Finding Latest BIOS Version" 
Try{
    [xml]$Script:ModelXML = Get-Content "$($TempFolder)\$($Script:BIOSXMLName)" -ErrorAction Stop
    $Script:BIOSPacks=@($Script:ModelXML.Manifest.SoftwareComponent | Where-Object {$_.ComponentType.Value -eq "BIOS"})
    $Script:BIOSPacks = $Script:BIOSPacks | ForEach-Object {
        [PSCustomObject]@{
            Version = $_.vendorVersion
            PaddedVersion=[version]$(
                $Padded=$_.vendorVersion -replace "[^0-9.]", "" -replace "^0", ""
                $Padded += ".0" * (4 - $Padded.Split('.').Count)
                $Padded
            )
            DownloadPath = "https://$($Script:ModelXML.Manifest.baselocation)/$($_.path)"
        }
    }
    $Script:BIOSPack = $Script:BIOSPacks| Sort-Object -Property PaddedVersion -Descending | Select-Object -First 1
    Write-LogEntry -Value "Successfully retrieved BIOS version: $($Script:BIOSPack.Version)" 
} Catch{
    Write-LogEntry -Value "Failed to retrieve latest BIOS version"
    Exit-Script -ExitCode 109
}

#Compare Latest BIOS to System BIOS
Write-LogEntry -Value " " 
Write-LogEntry -Value "Checking if BIOS update is required" 
$Script:UpgradeRequired=$False
If ($Script:BIOSPack.PaddedVersion -gt $Script:BIOSPaddedVersion){
    Write-LogEntry -Value "Latest BIOS, version $($Script:BIOSPack.Version) is greater than System BIOS, version $($Script:BIOSVersion)" 
    Write-Logentry -Value "BIOS update required"
    $Script:UpgradeRequired=$True
} elseif ($Script:BIOSPack.PaddedVersion -eq $Script:BIOSPaddedVersion) {
    Write-LogEntry -Value "Latest BIOS, version $($Script:BIOSPack.Version) is the same as System BIOS, version $($Script:BIOSVersion)" 
} else {
    Write-LogEntry -Value "Latest BIOS, version $($Script:BIOSPack.Version) is older than System BIOS, version $($Script:BIOSVersion)" 
}

If ($Script:UpgradeRequired -eq $True){
    #Download BIOS Package
    Write-LogEntry -Value " " 
    Write-LogEntry -Value "Downloading BIOS Package" 
    Try{
        $Script:BIOSEXEName = Split-Path -Path $Script:BIOSPack.DownloadPath -Leaf
        Invoke-WebRequest -Uri $Script:BIOSPack.DownloadPath -OutFile "$($TempFolder)\$($Script:BIOSEXEName)" -ErrorAction Stop
        Write-LogEntry -Value "Successfully downloaded $($Script:BIOSPack.DownloadPath)" 
    }
    Catch{
        Write-LogEntry -Value "Failed to download $($Script:BIOSPack.DownloadPath)"
        Exit-Script -ExitCode 110
    }

    #Install BIOS Package
    Write-LogEntry -Value " " 
    Write-LogEntry -Value "Installing BIOS Package"
    Try {
        $BIOS=Start-Process "$($TempFolder)\Flash64W.exe" -ArgumentList "/b=$($TempFolder)\$($Script:BIOSEXEName) /s" -Wait -PassThru
        $Script:BIOSExitCode=$BIOS.ExitCode
        Switch ($BIOSExitCode){
            0{
				Write-LogEntry -Value "Successfully installed $($TempFolder)\$($Script:BIOSEXEName)"
				Exit-Script -ExitCode $BIOSExitCode
				break
			}
            2{
				Write-LogEntry -Value "Successfully installed $($TempFolder)\$($Script:BIOSEXEName) - Reboot Required"
				Exit-Script -ExitCode $BIOSExitCode
				break
			}
            3{
				Write-LogEntry -Value "Same version is already installed"
				Exit-Script -ExitCode $BIOSExitCode
				break
			}
			8{
				Write-LogEntry -Value "Newer version is already installed"
				Exit-Script -ExitCode $BIOSExitCode
				break
			}
            default {throw $BIOSExitCode}
        }
        
    }
    Catch {
        Switch ($_){
            10{
				Write-LogEntry -Value "Failed to install $($TempFolder)\$($Script:BIOSEXEName) because the battery level is too low"
				Exit-Script -ExitCode $BIOSExitCode
				break
			}
            default {
				Write-LogEntry -Value "Failed to install $($TempFolder)\$($Script:BIOSEXEName) with exit code $($_)"
				Exit-Script -ExitCode $BIOSExitCode
			}
        }
        
    }

} else {
    Write-Logentry -Value "BIOS update not required"
    Exit-Script -ExitCode 3
}