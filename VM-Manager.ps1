#######################################
###### Function Definitions ###########
#######################################

function MainMenu {

    Write-Host "`r`n*********** Menu ***********`r`n" -ForegroundColor Yellow
    $actions = 'Power On', 'Power Off', 'Suspend', 'Create Snapshots', 'Delete Snapshots', 'Revert Snapshots', 'Show VM List','Extract Archives', 'Exit'
    $ln = 0
    $actionList = $actions | ForEach-Object {New-Object psObject -Property @{'ID' = $ln; 'Action'= $_.ToString()};$ln ++}
    $actionList | ForEach-Object {Write-Host "[$($_.ID)] $($_.Action)";}
    #RefreshvmList

    ### Get input ###
    Write-Host "`r`nSelect Action: " -ForegroundColor Yellow -NoNewline
    $actionSelected = Read-Host 
    switch($actionSelected)
    {
        0 {PowerOn}
        1 {PowerOff}
        2 {Suspend}
        3 {CreateSnapshot}
        4 {DeleteSnapshot}
        5 {RevertSnapshot}
        6 {clear; UpdateVMState; $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize}
        7 {ExtractArchives}
        8 {Write-Host 'Bye!';exit}
    }
}

function CreateVMList {
	
    Write-Host "Creating VM List..." -ForegroundColor Yellow

    # Get all VMs on disk and build table
    $dir = Get-ChildItem $vmPath -Recurse
    $i = 0
    $script:VMList = $dir | where{$_.Extension -eq '.vmx'} | % {
        New-Object psObject -Property @{
            'ID' = $i
            'VM'= $_.Name
            'Path' = $_.FullName
            'IsRunning' = 'No'
            'Snapshots' = if(!$script:snapsRetreived){'Not Checked'}
        }; 
        $i ++
    }
    UpdateVMState
}

function UpdateVMState {
	
    Write-Host "Updating VM State..." -ForegroundColor Yellow

    # Detect running VMs
    $runningVMs = & $vmrun list | select -skip 1;
    $script:VMList | % {
        if($runningVMs -and ($runningVMs -contains $_.Path)){
            $_.IsRunning = 'Yes';
        }else{
            $_.IsRunning = 'No';
        }
    }
}

function UpdateSnapshots {
	
	Write-Host "Updating VM Snapshots..." -ForegroundColor Yellow
	$i = 0
	$script:VMList | % {
        Write-Progress -Activity "Updating VM Snapshots..." -percentComplete (($i / $script:VMList.length)*100);
        if($runningVMs -and ($runningVMs -contains $_.Path)){
            $_.IsRunning = 'Yes';
        }else{
            $_.IsRunning = 'No';
        }
        $_.Snapshots = ((& $vmrun listsnapshots $($_.Path))| select -skip 1)
        $i++
    }
	Write-Progress -Completed $true
}

function ParseSelection {
    Param([string]$selection)

    $parsed = @();
    if($selection){
        $selection.split(',') | % {
            $_.trim() | Out-Null;
            if($_.ToString().Contains('-')){
                $tmp = $_.substring(0,($_.indexof('-')))..$_.substring(($_.lastindexof('-') + 1))
                $parsed = $parsed + $tmp;
            }else{
                $parsed += $_
            }
        }
    }else{
        $parsed="";
    }
    $parsed = $parsed | Select -Unique
    return $parsed;
}

function PowerOn {
    ### Display Menu ###    
    clear
    UpdateVMState
    Write-Host "`r`nSelect the VMs to Power On:`r`n" -ForegroundColor Yellow
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        $vmSelected | % {
            Write-Progress -Activity "Powering On...." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun start '$($script:VMList[$_.ToString()].Path)'"
            & $vmrun start $($script:VMList[$_.ToString()].Path)
            $i++;
        }
    }

    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n" -ForegroundColor Yellow
}

function PowerOff {
    ### Display Menu ###
    clear
    UpdateVMState
    Write-Host "`r`nSelect the VMs to Power Off:`r`n" -ForegroundColor Yellow
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        $vmSelected | % {
            Write-Progress -Activity "Powering Off...." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun stop '$($script:VMList[$_.ToString()].Path)'"
            & $vmrun stop $($script:VMList[$_.ToString()].Path)
            $i++;
        }
    }
    Write-Progress -id 1 -Completed $true;
    
    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n" -ForegroundColor Yellow
}

function Suspend {
    ### Display Menu ###
    clear
    UpdateVMState
    Write-Host "`r`nSelect the VMs to Suspend:`r`n" -ForegroundColor Yellow
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        $vmSelected | % {
            Write-Progress -Activity "Suspending...." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun suspend '$($script:VMList[$_.ToString()].Path)'"
            & $vmrun suspend $($script:VMList[$_.ToString()].Path)
            $i++;
        }
    }

    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n" -ForegroundColor Yellow
}

function CreateSnapshot {
    ### Display Menu ###
    clear
    UpdateVMState
    UpdateSnapshots
    Write-Host "`r`nSelect the VMs to Snapshot:`r`n" -ForegroundColor Yellow
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    Write-Host 'Snapshot Name [leave blank to use date]: ' -ForegroundColor Yellow -NoNewLine
    $snapshotName = ""
    $snapshotName = Read-Host
        
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        if(!$snapshotName){$snapshotName=(Get-Date -format yyyyMMdd)}
        $vmSelected | % {
            Write-Progress -Activity "Creating Snapshot..." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun snapshot '$($script:VMList[$_.ToString()].Path)' $snapshotName"
            & $vmrun snapshot $($script:VMList[$_.ToString()].Path) $snapshotName
            $i++;
        }
    }

    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n" -ForegroundColor Yellow
}

function DeleteSnapshot {
    ### Display Menu ###
    clear
    UpdateVMState
    UpdateSnapshots
    Write-Host "`r`nSelect a VM to delete a Snapshot:`r`n"
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    Write-Host 'Snapshot Name: ' -ForegroundColor Yellow -NoNewLine
    $snapshotName = Read-Host
    
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        $vmSelected | % {
            Write-Progress -Activity "Deleting Snapshots...." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun deletesnapshot '$($script:VMList[$_.ToString()].Path)' $snapshotName"
            & $vmrun deletesnapshot $($script:VMList[$_.ToString()].Path) $snapshotName
            $i++;
        }
    }

    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n"
}

function RevertSnapshot {
 
    ### Display Menu ###
    clear
    UpdateVMState
    UpdateSnapshots
    Write-Host "`r`nSelect a VM to revert to a Snapshot:`r`n"
    $script:VMList | Format-Table -Property ID, VM, IsRunning, Snapshots -AutoSize
    Write-Host "`r`nList VM numbers (0,1-4,6): " -ForegroundColor Yellow -NoNewLine
    $vmSelected = ParseSelection(Read-Host)
    Write-Host 'Snapshot Name: ' -ForegroundColor Yellow -NoNewLine
    $snapshotName = Read-Host
    
    ### Perform Action ###
    $i = 0
    $vmCount = $vmSelected.length
    if($vmCount -ne 0) {
        $vmSelected | % {
            Write-Progress -Activity "Reverting to Snapshots...." -percentComplete (($i / $vmCount)*100);
            Write-Host "vmrun reverttosnapshot '$($script:VMList[$_.ToString()].Path)' $snapshotName"
            & $vmrun reverttosnapshot $($script:VMList[$_.ToString()].Path) $snapshotName
            $i++;
        }
    }

    ### Finish ###
    Write-Progress -Completed $true
    Write-Host "`r`nDone!`r`n"
}

function ExtractArchives {

    # Get Start Time
    $StartTime = (Get-Date)

    # Execute
    Get-ChildItem $zipsSource*.7z | % {& "C:\Program Files\7-Zip\7z.exe" "x" $_.fullname "-o$vmPath"}

    # Get End Time
    $endDTM = (Get-Date)

    # Echo Time elapsed
    Write-Host ""
    "Elapsed Time: $([timespan]::fromseconds(((Get-Date)-$StartTime).Totalseconds).ToString(“hh\:mm\:ss”)) hh:mm:ss"
}

function MenuHandler {
    
	if($mustExit){exit}

    ### Display Menu ###
    $actions = 'MainMenu', 'Exit'
    $ln = 0
    $actionList = $actions | % {New-Object psObject -Property @{'ID' = $ln; 'Action'= $_.ToString()};$ln ++}
    $actionList | ForEach-Object {Write-Host "[$($_.ID)] $($_.Action)";}
    
    Write-Host "`r`nSelect Action: " -ForegroundColor Yellow -NoNewline
    $actionSelected = Read-Host

    ### Perform Action ###
    switch($actionSelected)
    {
        0 {clear; $mustExit = $false;}
        1 {Write-Host 'Bye!';exit}
    }
    
}

function GetParms {

    #### Get variables from parms.txt file ########
    #
    # Optional Values Are:
    #
    # vmrun = C:\Program Files (x86)\VMware\VMware Workstation\vmrun.exe
    # vmPath = C:\Users\agivens\Documents\vms
    # unzipDestination = C:\destination
    # zipsSource = C:\source
    #

    $parmsPath = "$PSScriptRoot\parms.txt"

    if(Test-Path $parmsPath) {
        Write-Host "Pulling parameters from $parmsPath" -ForegroundColor Yellow
        Write-Host "The following parameters were found:" -ForegroundColor Yellow
        Get-Content $parmsPath | Foreach-Object{
           $var = $_.Split('=')
           New-Variable -Name $var[0] -Value $var[1] -Force -Scope Script
           Write-Host "     $($var[0]) = $($var[1])"
        }
    } else {
        Write-Host "No valid parms file found." -ForegroundColor Yellow
        Write-Host "Please create a 'parms.txt' file in the directory of this script, and define the following values:" -ForegroundColor Yellow
        Write-Host "`nvmrun = C:\Program Files (x86)\VMware\VMware Workstation\vmrun.exe `nvmPath = C:\Users\agivens\Documents\vms `nunzipDestination = C:\destination `nzipsSource = C:\source"
    }
    ### Check variables ###
    If (-Not (Test-Path $vmrun)){Write-Host "vmrun.exe not found. Check the path in the parms.txt file."}
    If (-Not (Test-Path $vmPath)){Write-Host "vmPath is not valid. Check the path in the parms.txt file."}
    If (Test-Path variable:local:b) {. $b}
    
}

function StartAdmin {
    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
        Write-Host "Not running as administrator! Attempting to raise permission level...." -ForegroundColor Yellow
        $arguments = "& '" + $myinvocation.mycommand.definition + "'"
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        Break
    } else {
        Write-Host "Running as administrator!" -ForegroundColor Yellow
    }
}

#######################################
############## Main ###################
#######################################

### Start Program ###
$script:snapsRetreived = $false
$script:VMList = @()

GetParms
#StartAdmin
CreateVMList
#UpdateSnapshots

while($true){
    MainMenu
}