# Set timestamp for use in filenames
$Date = Get-Date -Format "dd-MM-yyyy"
$Timestamp = "$Date 13.00"

# Create addresslist
$Addresses = New-Object System.Collections.ArrayList

#Add shared addresses to list
$Addresses.Add("nexus.officeapps.live.com") > $null
$Addresses.Add("roaming.officeapps.live.com") > $null
$Addresses.Add("vortex-win.data.microsoft.com") > $null

# Check which random OneDrive address was selected at 3:00 and add that address to list
If (Test-Path ".\$Date 03.00 DB5SCH101110729.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101110729.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH101101044.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101101044.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB6SCH102090410.wns.windows.com.temp.txt"){$RandomAddress = "DB6SCH102090410.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH101110724.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101110724.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB6SCH102090309.wns.windows.com.temp.txt"){$RandomAddress = "DB6SCH102090309.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH103090414.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH103090414.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH101110428.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101110428.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH101110114.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101110114.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH101101324.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH101101324.wns.windows.com"}
If (Test-Path ".\$Date 03.00 DB5SCH103100117.wns.windows.com.temp.txt"){$RandomAddress = "DB5SCH103100117.wns.windows.com"}
$Addresses.Add($RandomAddress) > $null

# Add OneDrive addresses to list
$Addresses.Add("settings-win.data.microsoft.com") > $null
$Addresses.Add("skydrive.wns.windows.com") > $null
$Addresses.Add("login.microsoftonline.com") > $null
$Addresses.Add("odc.officeapps.live.com") > $null

#Add Outlook address to list
$Addresses.Add("login.windows.net") > $null
$Addresses.Add("outlook.office365.com") > $null

# Add OneNote address to list
$Addresses.Add("ols.officeapps.live.com") > $null

#Add recommended addresses to list
$Addresses.Add("portal.office.com") > $null
$Addresses.Add("www.yammer.com") > $null

# Create nexus.officeapps.live.com.title.txt with seperator lines and time
Add-Content ".\$Timestamp nexus.officeapps.live.com.title.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Timestamp nexus.officeapps.live.com.title.txt" "|                                           13:00                                           |"
Add-Content ".\$Timestamp nexus.officeapps.live.com.title.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Timestamp nexus.officeapps.live.com.title.txt" ""

# Create title.txt files with addresses
foreach ($Address in $Addresses) {Add-Content ".\$Timestamp $Address.title.txt" "$Address`:443" -NoNewline}

# Add Comments to title files
Add-Content ".\$Timestamp nexus.officeapps.live.com.title.txt" " (Shared infrastructure, Word Excel PowerPoint 100%, Outlook 100%, OneNote 100%)" -NoNewLine
Add-Content ".\$Timestamp roaming.officeapps.live.com.title.txt" " (Unknown, Word Excel PowerPoint 100%, Outlook 100%, OneNote 100%)" -NoNewLine
Add-Content ".\$Timestamp vortex-win.data.microsoft.com.title.txt" " (Unknown, Word Excel PowerPoint 80%, OneDrive 10% , Outlook 80%, OneNote 100%)" -NoNewLine

Add-Content ".\$Timestamp $RandomAddress.title.txt" " (Unknown, Random OneDrive 10%/100%)" -NoNewLine

Add-Content ".\$Timestamp settings-win.data.microsoft.com.title.txt" " (Unknown, OneDrive 10%)" -NoNewLine
Add-Content ".\$Timestamp skydrive.wns.windows.com.title.txt" " (Unknown, OneDrive 100%)" -NoNewLine
Add-Content ".\$Timestamp login.microsoftonline.com.title.txt" " (Authentication and identity, OneDrive 100%)" -NoNewLine
Add-Content ".\$Timestamp odc.officeapps.live.com.title.txt" " (OneDrive for Business: Determines consumer v commercial, OneDrive 100%)" -NoNewLine

Add-Content ".\$Timestamp login.windows.net.title.txt" " (Authentication and identity, Outlook 10%)" -NoNewLine
Add-Content ".\$Timestamp outlook.office365.com.title.txt" " (Exchange Online, Outlook 100%)" -NoNewLine

Add-Content ".\$Timestamp ols.officeapps.live.com.title.txt" " (Unknown, OneNote 100%)" -NoNewLine

Add-Content ".\$Timestamp portal.office.com.title.txt" " (Office 365 Portal, Microsoft recommendation)" -NoNewLine
Add-Content ".\$Timestamp www.yammer.com.title.txt" " (Yammer, Microsoft recommendation)" -NoNewLine

# Ping all addresses, write output to temp.result.txt files
foreach ($Address in $Addresses) {Start-Process .\psping64.exe -NoNewWindow -RedirectStandardOutput ".\$Timestamp $Address.temp.result.txt" "$Address`:443 -nobanner -accepteula -n 10s -i 1 -q -h 10"}

# Make sure pings are done and files are free
Start-Sleep 15

# Delete unwanted newlines from temp.result.txt files and write to result.txt files
foreach ($Address in $Addresses) {Get-Content ".\$Timestamp $Address.temp.result.txt" | Where-Object {$_ -replace "`n",""} | Set-Content ".\$Timestamp $Address.result.txt"}

# Merge title.txt and result.txt files to temp.txt files
foreach ($Address in $Addresses) {Get-Content ".\$Timestamp $Address.title.txt",".\$Timestamp $Address.result.txt" | Set-Content ".\$Timestamp $Address.temp.txt"}

# Delete all title.txt and result.txt files
Remove-Item .\*.title.txt,.\*.result.txt