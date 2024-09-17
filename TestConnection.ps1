$userInput = Read-Host "Please select an option and press enter:`n[1] Manual Input`n[2] .csv Input"  
$compList = @()

$outputFilePath = "C:\Temp\ConnectionTest.csv"

function CreateOutputFile([string]$outputFilePath) {
    If(!(Test-Path $outputFilePath)){
        New-Item -path $outputFilePath | Out-Null
    }
}

function IsOutputFileWritable {
    $testFile = New-Object System.IO.FileInfo $outputFilePath
    try{
        $openFile = $testFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None) 
        $openFile.Close()
    } catch {
        Write-Host("***ERROR***`nThe output file is unaccessible (" + $outputFilePath + "). Confirm you have the correct privledges and that the file is closed. Press Enter to exit the program.") -ForegroundColor Red -BackgroundColor White
        Read-Host("Press any key to exit...")
        exit(1)
    }
}

function IsConnected([string]$comp){
    
    If(Test-Connection -ComputerName $comp -Count 1 -Quiet) {
        $isConnected = $true
    } else {
        $isConnected = $false
    }
    return $isConnected
}

class ComputerObject {
    [string]$computerName
    [switch]$isOnline
    ComputerObject ([string]$computerName, [switch]$isOnline) {
        $This.computerName = $computerName
        $This.isOnline = $isOnline
    }
}

<#function SingleHost {
    CreateOutputFile $outputFilePath
    IsOutputFileWritable
    $userInput = Read-Host "`nPlease enter the hostname or IP address: "
    Write-Host("Processing Request, please wait...")
    $computerName = $userInput
    $isOnline = IsConnected $userInput
    $newComputerObject = [ComputerObject]::new($computerName,$isOnline)
    $newComputerObject | Export-Csv -LiteralPath $outputFilePath -NoTypeInformation -Force
    start $outputFilePath
    Write-Host("Evaluation Completed: " + $comp)
}
#>

function ManualInput {
    CreateOutputFile $outputFilePath
    IsOutputFileWritable
    $userInput = Read-Host("`nPlease input ALL computer hostnames or IPs. If you are attempting to check multiple clients, separate each with a comma, then press Enter.`nEX: computer1a,computer2b,computer3c")
    $computerArray = $userInput.Split(',')
    $compArray = $userInput.Split(',')
    $compObjArray = @()
    $currentCompNum = 1
    $totalCompNum = $compArray.Length
    foreach($comp in $compArray){
        Write-Host("Evaluating Host(" + $currentCompNum + "/" + $totalCompNum + "): " + $comp)
        $computerName = $comp
        $isOnline = IsConnected $comp
        $newComputerObject = [ComputerObject]::new($computerName,$isOnline)
        $compObjArray += $newComputerObject 
        $currentCompNum++
    }
    $compObjArray | Export-Csv -LiteralPath $outputFilePath -NoTypeInformation -Force
    start $outputFilePath
}

function CsvInput {
    CreateOutputFile $outputFilePath
    IsOutputFileWritable 
     Write-Host("`nPlease select the input .csv file")
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Filter = 'Documents (*.csv)|*.csv'
    }
    $clickResult = $fileBrowser.ShowDialog()
    if($clickResult -eq [System.Windows.Forms.DialogResult]::Cancel){
        Exit
    }
    Write-Host("Loaded File: " + $fileBrowser.FileName)    
    $importFile = Import-Csv -LiteralPath $fileBrowser.FileName | select Name
    $compObjArray = @()
    $currentCompNum = 1
    $totalCompNum = $importFile.Length
    foreach($comp in $importFile){
        Write-Host("Evaluating Host(" + $currentCompNum + "/" + $totalCompNum + "): " + $comp.Name)
        $computerName = $comp.Name
        $isOnline = IsConnected $comp.Name
        $newComputerObject = [ComputerObject]::new($computerName,$isOnline)
        $compObjArray += $newComputerObject 
        $currentCompNum++
    }
    $compObjArray | Export-Csv -LiteralPath $outputFilePath -NoTypeInformation -Force
    start $outputFilePath
}


#######################

if($userInput -eq 1){
    ManualInput
} elseif($userInput -eq 2) {
    CsvInput
} else {
    Write-Host("Please input a valid option")
}
########################################
