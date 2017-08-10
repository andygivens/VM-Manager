# VM-Manager
Menu driven PowerShell front-end for VMRun executable, allowing bulk actions on VMs in VMware Workstation.

### Usage
Fill in the correct parameters in the parms.txt file below:
```powershell
vmrun=C:\Program Files (x86)\VMware\VMware Workstation\vmrun.exe
vmPath=C:\Users\....\Documents\vms
unzipDestination=x:\destination
zipsSource=X:\source
```
Then right-click and "Run with PowerShell" on the VM-Manager.ps1

PS1 and parms files must live in the same diretory.
