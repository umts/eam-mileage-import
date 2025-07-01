Powershell script for retrieving a mileage export over HTTP and importing it
into Trapeeze EAM

# Configuration

Copy the `config.example` file to `config` and update the values as
appropriate:

* `source`: An HTTP(S) address which the script will fetch that contains
  the CSV-ish data that EAM will import. See below for a description of the
  file contents.
* `fa_gui`: The path to `fa_gui.exe` from a locally-installed copy of the
  EAM GUI.
* `fa_address`: The domain name or IP address of the EAM application server.
* `fa_port`: The port that the application service is running on.
* `fa_user`: The username of a user that has rights to insert meter readings
  (and ideally no other permissions).
* `fa_password`: That user's password.
* `smtp_server`: The SMTP server to use for sending email notifications for
  some rejection errors. If unset, no email will be sent.
* `email_to`: The email address(es) to send notifications to, comma-separated.
* `email_from`: The email address to use as the "from" address in the
  notification emails.

# Input file format

The fetched file is CSV-like with the following format:

**Line 1**: Full path to the error log file followed by a semicolon followed by
the full path to the log file. The powershell script assumes these to be called
`usage.err` and `usage.log` in the working directory of the script. The script
will also gsub the special string "`#DIRECTORY#` with the working directory.

**Line 2**: `2151` (Which is the screen id from EAM)

All remaining lines:

```
I,0,1,1410,2,<MILEAGE-TIME>,3,1,0,8,1411,2,<EQUIPMENT-NAME>,1411,4,<MILEAGE-TIME>,1411,5,<MILEAGE>,1411,6,NO EQ UPD,1411,9,,1411,10,,1411,11,,1411,14,-1,0,0,
```

This magic incantation _does_ technically depend on your specific EAM install.
But, its been pretty stable since version 6.2.2 as long as you haven't done
_major_ modifications to the screens and controls. It was generated using the
"Batch processing" setup instructions in the FASuite help.

# Other files created by the script

* `mileage.cmd`: The actual "command" file that EAM is given to run. It contains
  only the screen number to start on and the path to the downloaded CSV.
* `usage.err`: A log file containing the import errors from the last run. n.b.
  the line numbers listed in this file _exclude_ the first two lines in the CSV.
  The script only keeps the last 500 lines.
* `usage.log`: A log file containing a summary of each run (start time, number
  of committed records, number of rejected records, end time). The script only
  keeps the last 500 lines.
* `usage.rej`: A file containing the rejected _lines_ from the CSV. EAM makes
  this file because you could, in theory, manually edit the file and only
  re-process the rejections.
* `usageNN.csv` a backup copy of the downloaded file from day NN of the month.
* `usageNN.rej` a backup copy of the rejection file from day NN of the month.
