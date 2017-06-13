#This script makes the administrative changes necessary for your system to install and run XXX well.

#Map DSRC and Developer_tools to drives for this PS admin session if not already mapped
if(!(Test-Path("Z:")))
{
net use Z: \\DSRC.dgo.dcc.ccn\TOSE
"TOSE Mapped to drive Z:"
}
if(!(Test-Path("W:")))
{
net use W: \\DSHFS01.dgo.dcc.ccn\developer_tools
"Developer Tools Mapped to drive W:"
}

#The next two lines are one statement.  Changes Windows registry value to allow additional heap size.
set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems\' -name Windows -Value `
(((Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\SubSystems\' -name Windows).windows).replace("768","1024"))
"Heap size increased to 1024"

#This block creates a firewall rule for XXX then adds it to the active rules list.

#Creates a firewall rules list variable and adds a rule variable if there isn't one yet.
$fw= (new-object -ComObject HNetCfg.FwPolicy2).rules 

if(!(($fw|%{$_.name}) -contains "XXX"))
{ 
    $rule=new-object -ComObject HNetcfg.fwrule
    #Name rule, set to XXX, specify port numbers, set as XXX directed rule
    $rule.Name="XXX"
    $rule.Protocol=1 
    $rule.LocalPorts="1,2,3"
    $rule.Direction=2 
    #Enables rule and adds it to active list
    $rule.enabled=$true 
    $fw.add($rule)
    "XXX Firewall rule added"
} 

#Changes local security policy settings to disable FIPS algorithm compliance. 
secedit /export /cfg c:\secpol.cfg 
(gc "C:\secpol.cfg")|%{$_=$_ -replace "FIPSAlgorithmPolicy\\Enabled=4,1","FIPSAlgorithmPolicy\Enabled=4,0";$_}|Out-file c:\secpol.cfg
secedit /configure /db c:\windows\security\local.sdb /cfg c:\secpol.cfg /areas SECURITYPOLICY
rm -force C:\secpol.cfg -confirm:$false
"Disabled FIPS Algorithm Compliance"

#Creates reminder popup to contact SOC
(new-object -ComObject wscript.shell).popup("Contact (Call) the SOC office.  Let them know the server name so they may add it to their variance list.",0,"Contact SOC",1)


#This section checks pre-reqs and installs them if missing

#Sets PathExt environment variable to allow common extensions
[Environment]::SetEnvironmentVariable("Pathext",".exe;.pl;.com;.bat;.cmd;.vbs;.vbe;.js;.jse;.wsf;.wsh;.msc;.zip","Machine")

#Create a temporary install folder on your desktop
$InstallDir="$env:USERPROFILE\Desktop\InstallTmp"

if(!(test-path($InstallDir)))
{
  mkdir $InstallDir
}

#MS.NET 3.5
cd "MSNET"
if(!(dir).count -gt 1){"MS.NET 3.5 must be enabled before Habitat installation, please do that now"; pause}

#Perl 5.20
try
{
  perl -v|out-null; if(!((perl -v) -match "v5.20"))
  {
    Copy-Item "Perl" -Destination $InstallDir
    Start-Process "$InstallDir\ActivePerl-5.20.1.2000-MSWin3-x86-64int-298557.msi"
    pause
  }
    else{"Perl 5.20 already installed"}
}    
catch
{
  Copy-Item "Perl" -Destination $InstallDir
  Start-Process "$InstallDir\ActivePerl-5.20.1.2000-MSWin3-x86-64int-298557.msi"
  pause
}

#Java 1.8.0_45
try
{
  java -version;if(!(($Error|Out-String) -match "1.8.0_45"))
  {
     copy-item "Java" -Destination $InstallDir
     Start-Process "$InstallDir\jre-8u45-windows-x64.exe"
     pause
  }
    else{"Java 1.8.0_45 already installed"}
}
catch
{
  copy-item "Java" -Destination $InstallDir
  Start-Process "$InstallDir\jre-8u45-windows-x64.exe"
  pause
}

#7-Zip
if(!(test-path ("7Zip")))
{
  Copy-Item "7Zip" -Destination $InstallDir
  Start-Process "$InstallDir\7z920-x64.msi"
  pause
}
else{"7-Zip already installed"}

#PDF-XChange
if(!(Test-Path ("PDF")))
{
  Copy-Item "PDF" -Destination $InstallDir
  Start-Process "$InstallDir\PDFXVwer.exe"
  pause
}
else{"PDF-XChange already installed"}

#Eterrabrowser
if(!(Test-Path ("ETB")))
{
  Copy-Item "ETB" -Destination $InstallDir
  Start-Process "$InstallDir\e-terrabrowser_361_X64_20110531.1.msi"
  Set-Content "File" (gc "File"|%{$_ -replace "xx" ,"yy"})
  pause
}
else{"ETB already installed"}

#Notepad++
if(!(Test-Path("Notepad++")))
{
  Copy-Item "Notepad++" -Destination $InstallDir
  Start-Process "$InstallDir\npp.6.7.4.Installer.exe"
  pause
}
else{"Notepad++ already installed"}

#TortoiseSVN
if(!(Test-Path("Tortoise")))
{
  copy-item "Tortoise" -Destination $InstallDir
  Start-Process "$InstallDir\tortoisesvn-1.8.4.24972-x64-svn-1.8.5.msi"
  pause
}
else{"TortoiseSVN already installed"}

#Time to install XXX
if(!(Test-Path("XXX")))
{
  Copy-Item "XXX" -Destination $InstallDir
  Start-Process "$InstallDir\habitat_510.msi"
  pause
}
else{"XXX already installed"}

pause

