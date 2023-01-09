#Author: Zach McGill
#Date: 11/4/2022
#Written in PowerShell
#Run through Desktop Authority to get computer information and add it to a CSV file
#being used as the "database"

$outputCSV = "\\OutputFile\compInventory.csv"
$outputMessage = @()
$curDate = (Get-Date -Format "yyyy-MM-dd HH-mm-ss")
$newFile = $false

#Check to see if file exists, if not create it and the headers
if(!(Test-Path $outputCSV)){
    #Create the file
    New-Item -ItemType File -Path "\\server\Inventory\" -Name ("compInventory.csv")
    $newFile = $true
    #Create CSV Headers
    $outputMessage += "Name,Status,LastConnected,Model,Manufacturer,SerialNum,MAC,BIOSRelease,Processor,Memory,HDDSize,OS,Version,Updates,OSInstall,NumUsers,Run"
}

#Runs this command to get all of the info we want
$results = (Get-ComputerInfo -Property CsName,CsManufacturer,CsModel, BIOSSeralNumber,BIOSReleaseDate, 
    WindowsProductName,OsVersion, OsInstallDate, OsNumberOfUsers)

#Assign values to variables
$comp = $results.CsName
$computerModel = $results.CsModel
$computerSerial = $results.BIOSSeralNumber
$computerManu = $results.CsManufacturer
$BIOSRelease = $results.BIOSReleaseDate
$osProductName = $results.WindowsProductName
$osVersion = $results.OsVersion
$osInstallDate = $results.OsInstallDate
$osUserCount = $results.OsNumberOfUsers

#Get RamSize
$computerMem = (Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | %{[Math]::Round(($_.sum / 1GB),2)})

#Get HDD Size; DriveType=3 is the local drive
$hddSize = (Get-WmiObject -Class win32_logicaldisk -Filter "DriveType=3" | Measure-Object -Property Size -Sum | %{[Math]::Round(($_.sum / 1GB),2)})

#Get MAC Address
$macAddr = (Get-CimInstance win32_networkadapterconfiguration | select macaddress)
$macAddr = [string]$macAddr.macaddress

#Get Processor Info
$processor = Get-WmiObject win32_processor | select name
$processor = $processor.name
$computerProcessor = [string]$processor

#Gets the list of installed Updates
$updates = (Get-WmiObject -Class Win32_QuickFixEngineering | select HotFixId)
$osUpdates = [string]$updates.HotFixId

#If we just created the file, this does not need to be run
if($newFile -ne $true){
    $csvData = Import-Csv -path $outputCSV

    #Check if the computer is already in the "database"
    $indexOfRow = [array]::IndexOf($csvData.serialnum,$computerSerial)

    $count = [int]$csvData[$indexOfRow].Run
    $count = $count + 1

    #Find the row where the computer is
    if($indexOfRow -ne -1){
        #update the entire row with the new data
        $csvData[$indexOfRow].name = $comp
        $csvData[$indexOfRow].status = "Online"
        $csvData[$indexOfRow].LastConnected = $curDate
        $csvData[$indexOfRow].model = $computerModel
        $csvData[$indexOfRow].manufacturer = $computerManu
        $csvData[$indexOfRow].serialnum = $computerSerial
        $csvData[$indexOfRow].mac = $macAddr
        $csvData[$indexOfRow].biosrelease = $BIOSRelease
        $csvData[$indexOfRow].processor = $computerProcessor 
        $csvData[$indexOfRow].memory = $computerMem
        $csvData[$indexOfRow].hddsize = $hddSize
        $csvData[$indexOfRow].os = $osProductName
        $csvData[$indexOfRow].version = $osVersion
        $csvData[$indexOfRow].updates = $osUpdates
        $csvData[$indexOfRow].osinstall = $osInstallDate
        $csvData[$indexOfRow].numusers = $osUserCount
        $csvData[$indexOfRow].run = $count

        $csvData | Export-Csv -Path $outputCSV -nti    
    }else{
        #The computer is new and we need to just need to append its information
        $outputMessage += "$comp,Online,$curDate,$computerModel,""$computerManu"",$computerSerial,$macAddr,$BIOSRelease,$computerProcessor,$computerMem,$hddSize,$osProductName,$osVersion,$osUpdates,$osInstallDate,$osUserCount,1"
        $outputMessage | Out-File $outputCSV -Encoding utf8 -Append
    }
}else{
    #The computer is new and we need to just need to append its information
    $outputMessage += "$comp,Online,$curDate,$computerModel,""$computerManu"",$computerSerial,$macAddr,$BIOSRelease,$computerProcessor,$computerMem,$hddSize,$osProductName,$osVersion,$osUpdates,$osInstallDate,$osUserCount,1"
    $outputMessage | Out-File $outputCSV -Encoding utf8 -Append
}