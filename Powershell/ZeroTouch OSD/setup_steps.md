Set-ExecutionPolicy RemoteSigned -Force

Install-Module OSD -Force

Import-Module OSD -Force

New-OSDCloudTemplate

New-OSDCloudWorkspace -WorkspacePath C:\OSDCloud

New-OSDCloudUSB

Edit-OSDCloudwinPE -workspacepath C:\OSDCloud -CloudDriver * -WebPSScript https://gist.githubusercontent.com/SammySDGH/fea2e4bee4bf1aebf57df44b617d9d60/raw/a1801bd56940bb6bd7b1913de365a2ffc089800b/sd11_osdcloud_config.ps1 -Verbose

New-OSDCloudISO

Update-OSDCloudUSB