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

If (Test-Path "$pwd\usage.rej") { $rej = Get-ChildItem "$pwd\usage.rej" }
Else { $rej = New-Item -ItemType File -Path "$pwd\usage.rej" }

# Launch FA
$fa_args = @(
  "$($config.fa_address):$($config.fa_port)",
  "$($config.fa_user)/$($config.fa_password)",
  $cmd_file
)

Add-Content $log "[$(Get-Date -Format s)] - Starting Import"
& $config.fa_gui $fa_args | Wait-Process
Add-Content $log "[$(Get-Date -Format s)] - Import Complete"

# Analyze rejections
$errors = @()

# Reject file contains rejected lines with a 2-line header.
if (($rej_lines = Get-Content $rej).Length -gt 2) {
  $rej_lines = $rej_lines | Select-Object -Skip 2
  $rej_count = $rej_lines.Count

  $err_lines = Get-Content $err | Select-Object -Last $rej_count

  # Match-up the rejected lines with the error lines from the log
  for ($i = 0; $i -lt $rej_count; $i++) {
    $r = $rej_lines[$i]
    $e = $err_lines[$i]

    # 49: "Meter 1 value is less than last reading"
    # 437: "Meter 1 value is out of edit range from last reading"
    if ($e -match "PR-ERR:\s+(\d+)" -and $Matches[1] -in @("49", "437")) {
      $errors += @{
        Rejected = $r
        Error = $e
      }
    }
  }

  # Send email if there are errors and there are email settings configured
  if ($errors.Count -gt 0 -and $config.smtp_server) {
    $mail_body = "The following mileage records were rejected by FA during import:`n`n"
    foreach ($e in $errors) {
      $mail_body += "Rejected: $($e.Rejected)`nError: $($e.Error)`n`n"
    }

    Send-MailMessage -To ($config.email_to -split "\s*,\s*") `
                     -From $config.email_from `
                     -Subject "FA Mileage Import Errors" `
                     -Body $mail_body `
                     -SmtpServer $config.smtp_server `
  }
}

# Archive inputs/outputs
$day = Get-Date -Format dd
Copy-Item -Path $dest -Destination ($dest.BaseName + $day + $dest.Extension)
Copy-Item -Path $rej -Destination ($rej.BaseName + $day + $rej.Extension)

Get-Content $log | Select -Last 500 | Set-Content $log
Get-Content $err | Select -Last 500 | Set-Content $err
