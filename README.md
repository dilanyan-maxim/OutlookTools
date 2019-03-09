## Overview

This is a fork of original repository by Aman Dhally
amandhally@gmail.com
www.amandhally.net
http://newdelhipowershellusergroup.blogspot.in/

OutlookTools is a PowerShell module for Outlook. With the help of this module you can create new:
* task
* contact
* appointment
* email

## Changed from original module

Code cleanup: replaced Write-Host with Write-Verbose, Write-Warning with Write-Error
Code errors correction

Added new functions:
* New-OutlookTask

## Installation

With PowerShell:
```
$uri = 'https://raw.githubusercontent.com/dilanyan-maxim/OutlookTools/master'
$files = @('OutlookTools.psd1','OutlookTools.psm1')
$docs = [Environment]::GetFolderPath('MyDocuments')
$PSmodules = Join-Path -Path "$docs" -ChildPath 'WindowsPowerShell\Modules'
if (!(Test-Path -Path "$PSmodules\OutlookTools")) {
    New-Item -Path "$PSmodules" -Name 'OutlookTools' -ItemType Directory 
}
foreach ($file in $files) {
    $source = "$uri/$file"
    $dest = "$PSmodules\OutlookTools\$file"
    (New-Object System.Net.WebClient).DownloadFile($source, $dest)
}
Remove-Variable -Name uri,files,docs,PSmodules,source,dest
```

## Usage

```
Import-Module -Name OutlookTools
```
