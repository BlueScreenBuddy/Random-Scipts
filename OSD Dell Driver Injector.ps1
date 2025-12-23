#Initialize Task Sequence Environment
$TSEnv=New-Object -COMObject Microsoft.SMS.TSEnvironment
$SMSTSLogPath=$TSEnv.Value("_SMSTSLogPath")
$OSDTargetSystemDrive=$TSEnv.Value("OSDTargetSystemDrive")

#Disable Progress bars to speed up Invoke-WebRequest
$ProgressPreference = 'SilentlyContinue'

#Define Log file path
$Script:Logfile="$($SMSTSLogPath)\Apply-Drivers.log"


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
Write-LogEntry -Value "Started Driver Injection Script" 

#Create Temp Folder
Write-LogEntry -Value " " 
Write-LogEntry -Value "Creating Drivers folder in Temp" 
$TempFolder="$($OSDTargetSystemDrive)\Drivers"
If (!(Test-Path $TempFolder)){
    Try{
        New-Item -Path "$($OSDTargetSystemDrive)" -Name "Drivers" -ItemType Directory -Force -ErrorAction Stop
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

#Download Model Catalog XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Downloading Driver Pack Catalog XML" 
Try{
    $Script:CatalogCABName="DriverPackCatalog.cab"
    Invoke-WebRequest -Uri "https://downloads.dell.com/catalog/$Script:CatalogCABName" -OutFile "$($TempFolder)\$Script:CatalogCABName" -ErrorAction Stop
    Write-LogEntry -Value "Successfully downloaded https://downloads.dell.com/catalog/$Script:CatalogCABName" 
}
Catch{
    Write-LogEntry -Value "Failed to download https://downloads.dell.com/catalog/$Script:CatalogCABName"
    Exit-Script -ExitCode 102
}

#Extract Catalog XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Extracting Driver Pack Catalog XML" 
Try{
    $Script:CatalogXMLName = $Script:CatalogCABName -Replace "cab","xml"
    Start-Process "$($env:windir)\system32\expand.exe" -WindowStyle Hidden -ArgumentList "`"$($TempFolder)\$($Script:CatalogCABName)`" `"$($TempFolder)\$($Script:CatalogXMLName)`"" -Wait -ErrorAction Stop
    Write-LogEntry -Value "Successfully extracted $($Script:CatalogXMLName) from $($Script:CatalogCABName)" 
}
Catch{
    Write-LogEntry -Value "Failed to extract $($Script:CatalogXMLName) from $($Script:CatalogCABName)"
    Exit-Script -ExitCode 103    
}

#Get Highest Driver Pack Version from XML
Write-LogEntry -Value " " 
Write-LogEntry -Value "Finding Latest Driver Pack Version"
$Script:DriversFound -eq $False
Try{
    [xml]$Script:DriversXML= Get-Content "$($TempFolder)\$($Script:CatalogXMLName)" -ErrorAction Stop
    $Script:DriverPacks=@($Script:DriversXML.DriverPackManifest.DriverPackage | Where-Object {$_.SupportedSystems.Brand.Model.systemID -eq $Script:SystemSku -and $_.supportedoperatingsystems.operatingsystem.oscode -like "*11"})
    $Script:DriverPacks = $Script:DriverPacks | ForEach-Object {
        [PSCustomObject]@{
            Version = $_.dellVersion
            PaddedVersion=[version]$(
                $Padded=$_.dellVersion -replace "[^0-9.]", "" -replace "^0", ""
                $Padded += ".0" * (4 - $Padded.Split('.').Count)
                $Padded
            )
            DownloadPath = "https://$($Script:DriversXML.DriverPackManifest.baselocation)/$($_.path)"
            MD5Hash=$_.hashMD5
        }
    }
    $Script:DriverPack = $Script:DriverPacks | Sort-Object -Property PaddedVersion -Descending | Select-Object -First 1
    Write-LogEntry -Value "Successfully retrieved Driver Pack version: $($Script:DriverPack.Version)"
} Catch{
    Write-LogEntry -Value "Failed to retrieve latest Driver Pack version"
    Exit-Script -ExitCode 104
}

#Download Driver Pack
Write-LogEntry -Value " " 
Write-LogEntry -Value "Downloading Driver Pack" 
Try{
    $Script:DriversEXEName = Split-Path -Path $Script:DriverPack.DownloadPath -Leaf
    Invoke-WebRequest -Uri $Script:DriverPack.DownloadPath -OutFile "$($TempFolder)\$($Script:DriversEXEName)" -ErrorAction Stop
    Write-LogEntry -Value "Successfully downloaded $($Script:DriverPack.DownloadPath)" 
}
Catch{
    Write-LogEntry -Value "Failed to download $($Script:DriverPack.DownloadPath)"
    Exit-Script -ExitCode 105
}

#Verify Driver Pack MD5 Hash
Write-LogEntry -Value " " 
Write-LogEntry -Value "Verifying integrity of download"
Try{
    $MD5=(Get-FileHash "$($TempFolder)\$($Script:DriversEXEName)" -Algorithm MD5).Hash
    If ($MD5 -eq $Script:DriverPack.MD5Hash){
        Write-LogEntry -Value "Successfully verified integrity of download"
    } else {
        Write-LogEntry -Value "Failed to verify integrity of download"
        Exit-Script -ExitCode 106
    }
}
Catch{
    Write-LogEntry -Value "Failed to verify integrity of download"
    Exit-Script -ExitCode 106

}

#Extract Driver Pack
Write-LogEntry -Value " " 
Write-LogEntry -Value "Extracting Driver Pack"
Try {
    If (!(Test-Path $($TempFolder))){
        New-Item -Path "$($TempFolder)" -Name "DriverPack" -ItemType Directory -Force
    }
    $Drivers=Start-Process "$($TempFolder)\$($Script:DriversEXEName)" -ArgumentList "/s /e=`"$($TempFolder)\DriverPack`" /l=`"$($TempFolder)\dell.log`"" -Wait -PassThru -ErrorAction Stop
    $Script:DriversExitCode=$Drivers.ExitCode
    Switch ($DriversExitCode){
        0{Write-LogEntry -Value "Successfully extracted $($TempFolder)\$($Script:DriversEXEName) to $($TempFolder)\DriverPack"}
        default {throw $DriversExitCode}
    }
}
Catch {
    Switch ($_){
        default {Write-LogEntry -Value "Failed to extract $($TempFolder)\$($Script:DriversEXEName) with exit code code $($_)"}
    }
    Exit-Script -ExitCode 107
}

#Inject Drivers into image
Write-LogEntry -Value " " 
Write-LogEntry -Value "Injecting Driver Pack"
Try{
    $DRIVERS=Start-Process "Dism.exe" -WindowStyle Hidden -ArgumentList "/Image:$($OSDTargetSystemDrive)\ /Add-Driver /Driver:$($TempFolder)\DriverPack /Recurse" -Wait -PassThru
    $DRIVERSExitcode=$DRIVERS.Exitcode
    Switch ($DRIVERSExitcode){
       0 {Write-LogEntry -Value "Driver injection completed successfully"}
       default {Write-LogEntry -Value "Driver injection completed with exit code $($DRIVERSExitcode)"}
    }  
}
Catch{
    Write-LogEntry "Failed to inject drivers from $($TempFolder)\DriverPack into image"
    Exit-Script -ExitCode 108
}

#Exit  the script with success code
Exit-Script -ExitCode 0