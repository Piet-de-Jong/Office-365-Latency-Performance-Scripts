These PowerShell scripts were written as part of my graduation assignment. These are the demo versions I used when presenting my thesis. I wanted to put the code online in case someone could find a use for it. A friend suggested I put it up on GitHub, so here it is. Feel free to use the code as review material for your own work. 

To be able to run these PowerShell scripts you need at least PowerShell 5.0, which is part of Windows Management Framework 5.0, which can be downloaded here: https://www.microsoft.com/en-us/download/details.aspx?id=50395 

Windows 10 has PowerShell 5.0 installed by default.

You’ll also need PsPing, which is part of PsTools, which can be downloaded here: https://technet.microsoft.com/en-us/sysinternals/psping.aspx

The scripts were designed to use the 64 bits version of PsPing, but that can easily be changed.

The scripts were designed to run daily at the designated times. Log Office 365 Latency 03.00.ps1 at 03:00, Log Office 365 Latency 08.30.ps1 at 08:30, Log Office 365 Latency 13.00.ps1 at 13:00 and Log Office 365 Latency 16.30.ps1 at 16:30. At the end of the day another script, Mail Log Summary.ps1, would use the data gathered throughout the day to send a summary by mail of the latency performance that day compared to a baseline. 

Although these scripts were written to log the office 365 latency performance and compare them to a baseline, the addresses can easily be changed to something else. The scripts were also designed to be used from different networks to be able to compare performance. All formatting and naming conventions are as requested by the organization these scripts were designed for. I hope someone may find use for these scripts or snippets of code. Piet de Jong.
