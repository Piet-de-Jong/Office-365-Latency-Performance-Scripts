# Set date for use in filenames
$Date = Get-Date -Format "dd-MM-yyyy"

# Create addresslist
$Addresses = New-Object System.Collections.ArrayList

# Add shared addresses to list
$Addresses.Add("nexus.officeapps.live.com") > $null
$Addresses.Add("roaming.officeapps.live.com") > $null
$Addresses.Add("vortex-win.data.microsoft.com") > $null

# Check which random OneDrive address was selected and add that address to list
If (Test-Path ".\$Date 03.00 DB5SCH101110729.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101110729.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH101101044.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101101044.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB6SCH102090410.wns.windows.com.temp.txt"){$Addresses.Add("DB6SCH102090410.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH101110724.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101110724.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB6SCH102090309.wns.windows.com.temp.txt"){$Addresses.Add("DB6SCH102090309.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH103090414.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH103090414.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH101110428.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101110428.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH101110114.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101110114.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH101101324.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH101101324.wns.windows.com") > $null}
If (Test-Path ".\$Date 03.00 DB5SCH103100117.wns.windows.com.temp.txt"){$Addresses.Add("DB5SCH103100117.wns.windows.com") > $null}

# Add OneDrive addresses to list
$Addresses.Add("settings-win.data.microsoft.com") > $null
$Addresses.Add("skydrive.wns.windows.com") > $null
$Addresses.Add("login.microsoftonline.com") > $null
$Addresses.Add("odc.officeapps.live.com") > $null

# Add Outlook addresses to list
$Addresses.Add("login.windows.net") > $null
$Addresses.Add("outlook.office365.com") > $null

# Add OneNote address to list
$Addresses.Add("ols.officeapps.live.com") > $null

# Add recommended addresses to list
$Addresses.Add("portal.office.com") > $null
$Addresses.Add("www.yammer.com") > $null

# Load charting requirements
[void][Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms.DataVisualization")

foreach ($Address in $Addresses) {

# Create chartobject
$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart

# Set chartobject variables
[void]$Chart.Titles.Add(“$Address”)
$Chart.Width = 500
$Chart.Height = 400

if ($Address -eq "nexus.officeapps.live.com"){[void]$Chart.Titles.Add(“Shared infrastructure, Word Excel PowerPoint 100%, Outlook 100%, OneNote 100%”)}
if ($Address -eq "roaming.officeapps.live.com"){[void]$Chart.Titles.Add(“Unknown, Word Excel PowerPoint 100%, Outlook 100%, OneNote 100%”)}
if ($Address -eq "vortex-win.data.microsoft.com"){[void]$Chart.Titles.Add(“Unknown, Word Excel PowerPoint 80%, OneDrive 10% , Outlook 80%, OneNote 100%”)}

if ($Address -eq "DB5SCH101110729.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH101101044.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB6SCH102090410.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH101110724.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB6SCH102090309.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH103090414.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH101110428.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH101110114.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH101101324.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}
if ($Address -eq "DB5SCH103100117.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, Random OneDrive 10%/100%”)}

if ($Address -eq "settings-win.data.microsoft.com"){[void]$Chart.Titles.Add(“Unknown, OneDrive 10%”)}
if ($Address -eq "skydrive.wns.windows.com"){[void]$Chart.Titles.Add(“Unknown, OneDrive 100%”)}
if ($Address -eq "login.microsoftonline.com"){[void]$Chart.Titles.Add(“Authentication and identity, OneDrive 100%”)}
if ($Address -eq "odc.officeapps.live.com"){[void]$Chart.Titles.Add(“OneDrive for Business: Determines consumer v commercial, OneDrive 100%”)}

if ($Address -eq "login.windows.net"){[void]$Chart.Titles.Add(“Authentication and identity, Outlook 10%”)}
if ($Address -eq "outlook.office365.com"){[void]$Chart.Titles.Add(“Exchange Online, Outlook 100%”)}

if ($Address -eq "ols.officeapps.live.com" ){[void]$Chart.Titles.Add(“Unknown, OneNote 100%”)}

if ($Address -eq "portal.office.com"){[void]$Chart.Titles.Add(“Office 365 Portal, Microsoft recommendation”)}
if ($Address -eq "www.yammer.com"){[void]$Chart.Titles.Add(“Yammer, Microsoft recommendation")}

# Get average value of ping results of 03:00 from temp.txt files
$Night = Get-Content ".\$Date 03.00 $Address.temp.txt" | Where-Object {$_ -match 'Average'}
$Night = $Night  -match "X*Average = (?<value>.*)ms"
$Night = [decimal]$Matches["value"]

# Get average value of ping results of 08:30 from temp.txt files
$Login = Get-Content ".\$Date 08.30 $Address.temp.txt" | Where-Object {$_ -match 'Average'}
$Login = $Login  -match "X*Average = (?<value>.*)ms"
$Login = [decimal]$Matches["value"]

# Get average value of ping results of 13:00 from temp.txt files
$Midday = Get-Content ".\$Date 13.00 $Address.temp.txt" | Where-Object {$_ -match 'Average'}
$Midday = $Midday  -match "X*Average = (?<value>.*)ms"
$Midday = [decimal]$Matches["value"]

# Get average value of ping results of 16:30 from temp.txt files
$Logout = Get-Content ".\$Date 16.30 $Address.temp.txt" | Where-Object {$_ -match 'Average'}
$Logout = $Logout  -match "X*Average = (?<value>.*)ms"
$Logout = [decimal]$Matches["value"]

# Create chartarea
$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea

# Set chartarea variables 
$ChartArea.AxisX.Title = “Time”
$ChartArea.AxisX.Minimum = 1
$ChartArea.AxisX.Maximum = 4
$ChartArea.AxisY.Title = “Latency (ms)”
$ChartArea.AxisY.IsStartedFromZero = 0

# Add chartarea to chartobject
$Chart.ChartAreas.Add($ChartArea)

# Add empty chart
$Results= @{"03:00"="00.00";"08:30"="00.00";"13:00"="00.00";"16:30"="00.00"}
$Baseline= @{"03:00"="00.00";"08:30"="00.00";"13:00"="00.00";"16:30"="00.00"}

# Set chart content based on address
$Results.Keys = "03:00","08:30","13:00","16:30"
$Results.Values = $Night,$Login,$Midday,$Logout
$Baseline.Keys = "03:00","08:30","13:00","16:30"

if ($Address -eq "nexus.officeapps.live.com") {$Baseline.Values = "105.56","127.63","122.26","130.03"}
if ($Address -eq "roaming.officeapps.live.com") {$Baseline.Values = "20.51","23.07","20.37","24.08"}
if ($Address -eq "vortex-win.data.microsoft.com") {$Baseline.Values = "30.21","35.58","31.98","31.58"}

if ($Address -eq "DB5SCH101110729.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH101101044.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB6SCH102090410.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH101110724.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB6SCH102090309.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH103090414.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH101110428.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH101110114.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH101101324.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}
if ($Address -eq "DB5SCH103100117.wns.windows.com") {$Baseline.Values = "31.43","39.66","33.17","31.58"}

if ($Address -eq "settings-win.data.microsoft.com") {$Baseline.Values = "31.14","35.45","30.09","31.08"}
if ($Address -eq "skydrive.wns.windows.com") {$Baseline.Values = "31.52","35.63","32.23","33.59"}
if ($Address -eq "login.microsoftonline.com") {$Baseline.Values = "11.01","15.06","15.13","13.18"}
if ($Address -eq "odc.officeapps.live.com") {$Baseline.Values = "19.31","21.68","20.22","20.88"}

if ($Address -eq "login.windows.net") {$Baseline.Values = "10.44","15.40","11.98","12.40"}
if ($Address -eq "outlook.office365.com") {$Baseline.Values = "28.70","32.23","32.37","27.56"}

if ($Address -eq "ols.officeapps.live.com") {$Baseline.Values = "116.59","117.86","114.53","115.14"}

if ($Address -eq "portal.office.com") {$Baseline.Values = "10.24","13.78","11.52","10.52"}
if ($Address -eq "www.yammer.com") {$Baseline.Values = "10.12","12.83","11.38","12.87"}

# Add chart content
[void]$Chart.Series.Add(“Results”)
[void]$Chart.Series.Add(“Baseline”)
$Chart.Series[“Results”].Points.DataBindXY($Results.Keys, $Results.Values)
$Chart.Series[“Baseline”].Points.DataBindXY($Baseline.Keys, $Baseline.Values)

# Set chart variables
$Chart.Series["Results"].BorderWidth  = 3
$Chart.Series["Baseline"].BorderWidth  = 2
$Chart.Series["Results"].Color  = "blue"
$Chart.Series["Baseline"].Color  = "green"

# Set charttype
$Chart.Series[“Results”].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
$Chart.Series[“Baseline”].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line

# Change text fonts for readability
$Chart.Titles[0].Font = new-object system.drawing.font("calibri",12,[system.drawing.fontstyle]::Bold)
$Chart.Titles[1].Font = new-object system.drawing.font("calibri",10,[system.drawing.fontstyle]::Regular)
$ChartArea.AxisX. Titlefont = new-object system.drawing.font("calibri",12,[system.drawing.fontstyle]::Bold)
$ChartArea.AxisY.Titlefont = new-object system.drawing.font("calibri",12,[system.drawing.fontstyle]::Bold)
$Chart.chartAreas[0].AxisX.LabelStyle.Font = new-object system.drawing.font("calibri",12,[system.drawing.fontstyle]::Regular)
$Chart.chartAreas[0].AxisY.LabelStyle.Font = new-object system.drawing.font("calibri",12,[system.drawing.fontstyle]::Regular)

# Add legend 
$Legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
$Legend.name = "Legend1"
$Chart.Legends.Add($legend)

# Save charts
$Chart.SaveImage(“.\$Date-$Address.png”, “PNG”)

# Uncomment to write results to address.baseline.txt 
# Add-Content ".\$Address.baseline.txt" "$Date"
# Add-Content ".\$Address.baseline.txt" "03:00   08:30   13:00   16:30"
# if ($Night -lt 10) {Add-Content ".\$Address.baseline.txt" "$Night    " -NoNewline} ElseIf ($Night -ge 100) {Add-Content ".\$Address.baseline.txt" "$Night  " -NoNewline} Else {Add-Content ".\$Address.baseline.txt" "$Night   " -NoNewline}
# if ($Login -lt 10) {Add-Content ".\$Address.baseline.txt" "$Login    " -NoNewline} ElseIf ($Login -ge 100) {Add-Content ".\$Address.baseline.txt" "$Login  " -NoNewline} Else {Add-Content ".\$Address.baseline.txt" "$Login   " -NoNewline}
# if ($Midday -lt 10) {Add-Content ".\$Address.baseline.txt" "$Midday    " -NoNewline} ElseIf ($Midday -ge 100) {Add-Content ".\$Address.baseline.txt" "$Midday  " -NoNewline} Else {Add-Content ".\$Address.baseline.txt" "$Midday   " -NoNewline}
# if ($Logout -lt 10) {Add-Content ".\$Address.baseline.txt" "$Logout    " -NoNewline} ElseIf ($Logout -ge 100) {Add-Content ".\$Address.baseline.txt" "$Logout  " -NoNewline} Else {Add-Content ".\$Address.baseline.txt" "$Logout" -NoNewline}
# Add-Content ".\$Address.baseline.txt" ""
# Add-Content ".\$Address.baseline.txt" ""
}

# Append newline in all temp.txt files
foreach ($Address in $Addresses) {Add-Content ".\$Date 03.00 $Address.temp.txt" ""}
foreach ($Address in $Addresses) {Add-Content ".\$Date 08.30 $Address.temp.txt" ""}
foreach ($Address in $Addresses) {Add-Content ".\$Date 13.00 $Address.temp.txt" ""}
foreach ($Address in $Addresses) {Add-Content ".\$Date 16.30 $Address.temp.txt" ""}

# Append seperator line and new line to roaming.officeapps.live.com.temp.txt (after shared addresses)
Add-Content ".\$Date 03.00 vortex-win.data.microsoft.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 03.00 vortex-win.data.microsoft.com.temp.txt" ""
Add-Content ".\$Date 08.30 vortex-win.data.microsoft.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 08.30 vortex-win.data.microsoft.com.temp.txt" ""
Add-Content ".\$Date 13.00 vortex-win.data.microsoft.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 13.00 vortex-win.data.microsoft.com.temp.txt" ""
Add-Content ".\$Date 16.30 vortex-win.data.microsoft.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 16.30 vortex-win.data.microsoft.com.temp.txt" ""

# Append seperator line and new line to RandomAddress.temp.txt (after OneDrive)
Add-Content ".\$Date 03.00 odc.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 03.00 odc.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 08.30 odc.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 08.30 odc.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 13.00 odc.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 13.00 odc.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 16.30 odc.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 16.30 odc.officeapps.live.com.temp.txt" ""

# Append seperator line and new line to ols.officeapps.live.com.temp.txt (after OneNote)
Add-Content ".\$Date 03.00 ols.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 03.00 ols.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 08.30 ols.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 08.30 ols.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 13.00 ols.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 13.00 ols.officeapps.live.com.temp.txt" ""
Add-Content ".\$Date 16.30 ols.officeapps.live.com.temp.txt" "---------------------------------------------------------------------------------------------"
Add-Content ".\$Date 16.30 ols.officeapps.live.com.temp.txt" ""

# Merge temp.txt files to $Timestamp source.txt without lines containing 'warmpup' and 'for'
foreach ($Address in $Addresses) {Get-Content ".\$Date 03.00 $Address.temp.txt" | Where-Object {$_ -notmatch 'warmup'} | Where-Object {$_ -notmatch 'for'} | Add-Content ".\$Date.txt"}
foreach ($Address in $Addresses) {Get-Content ".\$Date 08.30 $Address.temp.txt" | Where-Object {$_ -notmatch 'warmup'} | Where-Object {$_ -notmatch 'for'} | Add-Content ".\$Date.txt"}
foreach ($Address in $Addresses) {Get-Content ".\$Date 13.00 $Address.temp.txt" | Where-Object {$_ -notmatch 'warmup'} | Where-Object {$_ -notmatch 'for'} | Add-Content ".\$Date.txt"}
foreach ($Address in $Addresses) {Get-Content ".\$Date 16.30 $Address.temp.txt" | Where-Object {$_ -notmatch 'warmup'} | Where-Object {$_ -notmatch 'for'} | Add-Content ".\$Date.txt"}

# Format emailbody as HTML
$Bodystring = foreach ($Address in $Addresses) {"<img src=cid:$Date-$Address.png><br>"}
$Body = "<html><body>$Bodystring</body></html>"

# Create list of attachments to send with email
$Attachments = foreach ($Address in $Addresses) {"$Date-$Address.png"}
$Attachments = $Attachments+".\$Date.txt"

# Send email with attachments as HTML
Send-MailMessage -To "Office 365 Performance<yourgmailaddress@gmail.com>" -From "Office 365 Performance<yourgmailaddress@gmail.com>" -Subject "Office 365 Latency Performance Scripts Demo $Date" -Attachments $Attachments -BodyAsHtml $Body -SmtpServer smtp.gmail.com -Port 587 -UseSsl -Credential (Get-Credential)

# Delete all temp.txt files
Remove-Item .\*.temp.txt

# Create subdirectory
New-Item -ItemType Directory -Path "$Date" > $null

# Select all .png files
$Images = Get-ChildItem -Path "*.png"

# Move files to subdirectory
Move-Item $Images $Date
Move-Item "$Date.txt" $Date