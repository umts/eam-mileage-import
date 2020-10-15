$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol =
  [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

# Read in the config file
$config = Get-Content $(Get-ChildItem "$pwd\config") | ConvertFrom-StringData

# Download the file
$dest = "$pwd\usage.csv"

$wc = New-Object System.Net.WebClient
$wc.DownloadFile($config.source, $dest)
$dest = Get-ChildItem $dest

# Find and replace for Absolute paths
(Get-Content $dest) |
Foreach-Object { $_ -replace "#DIRECTORY#", $pwd } |
Set-Content $dest

# Create "command" file
$cmd_file = "$pwd\mileage.cmd"
New-Item $cmd_file -ItemType File -Force -Value "2151 $dest"

# Create output files
If (Test-Path "$pwd\usage.log") { $log = Get-ChildItem "$pwd\usage.log" }
Else { $log = New-Item -ItemType File -Path "$pwd\usage.log" }

If (Test-Path "$pwd\usage.err") { $err = Get-ChildItem "$pwd\usage.err" }
Else { $err = New-Item -ItemType File -Path "$pwd\usage.err" }

# Launch FA
$fa_args = @(
  "$($config.fa_address):$($config.fa_port)",
  "$($config.fa_user)/$($config.fa_password)",
  $cmd_file
)

Add-Content $log "[$(Get-Date -Format s)] - Starting Import"
& $config.fa_gui $fa_args | Wait-Process
Add-Content $log "[$(Get-Date -Format s)] - Import Complete"

# Archive inputs/outputs
$day = Get-Date -Format dd
Copy-Item -Path $dest -Destination ($dest.BaseName + $day + $dest.Extension)

Get-Content $log | Select -Last 500 | Set-Content $log
Get-Content $err | Select -Last 500 | Set-Content $err
