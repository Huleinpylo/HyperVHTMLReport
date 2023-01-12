
$VmsDataArray =  New-Object -TypeName "System.Collections.ArrayList"
$VmsDataArrayDisque =  New-Object -TypeName "System.Collections.ArrayList"
function New-VmsData()
{
  param ($var)


  $VmsDataVar = new-object PSObject
  $VmsDataVar | add-member -type NoteProperty -Name HOSTNAME -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name VM_Nom -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name vCPU -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name vRam -value ""
  $VmsDataVar | add-member -type NoteProperty -Name IPAddresses -value ""
  #$VmsDataVar | add-member -type NoteProperty -Name Disque -value (New-Object -TypeName "System.Collections.ArrayList")
  return $VmsDataVar
}
function New-VHD()
{
  param ($var)


  $VmsDataVar = new-object PSObject
  $VmsDataVar | add-member -type NoteProperty -Name HOSTNAME -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name VM_Nom -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name Path -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name VhdFormat -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name VhdType -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name Max_Size -value ""
  $VmsDataVar | add-member -type NoteProperty -Name FileSize -value ""
   return $VmsDataVar
}
$vms = (Get-VM )

foreach ( $vm in $vms)
{
    $var=New-VmsData;
    $var.vCPU= (Get-VMProcessor -VMName $vm.Name).Count
    $getRAMInfo = Get-VMMemory -VMName $vm.Name

if($getRAMInfo.DynamicMemoryEnabled -eq $true)
{
#    Write-Line 'RAM type' 'Dynamic Memory'
    $var.vRam= ([string]($getRAMInfo.Startup / 1MB) + ' MB')
 #   Write-Line 'Minimum RAM' ([string]($getRAMInfo.Minimum / 1MB) + ' MB')
  #  Write-Line 'Maximum RAM' ([string]($getRAMInfo.Maximum / 1MB) + ' MB')
}else{
 # Write-Line 'RAM type' 'Static Memory'
  $var.vRam= ([string]($vm.MemoryStartup / 1MB) + ' MB')
}
$vmHDDs = (Get-VMHardDiskDrive -VMName $vm.Name  | Sort-Object Name)
   
ForEach($vmHDD in $vmHDDs)
{
    $vihd=New-VHD
    $vihd.Path= ($vmHDD.Path)

  $vmHDDVHD = $vmHDD.Path | Get-VHD  -ErrorAction SilentlyContinue
  if ($vmHDDVHD -ne $null) {
    $vihd.HOSTNAME=$env:COMPUTERNAME
    $vihd.VM_Nom=$vm.Name
    $vihd.VhdFormat= ($vmHDDVHD.VhdFormat)
    $vihd.VhdType= ($vmHDDVHD.VhdType)
    $vihd.Max_Size = ($vmHDDVHD.Size / 1GB).ToString() 
    $vihd.FileSize= ($vmHDDVHD.FileSize / 1GB).ToString()
    
  } else {
    Write-Error 'Disk Warning Error accessing virtual disk'  
  }
  $VmsDataArrayDisque.Add( $vihd)
}
$vmNetCards = (Get-VMNetworkAdapter -VMName $vm.Name)
 $var.IPAddresses = $vmNetCard.IPAddresses
 $var.HOSTNAME = $env:COMPUTERNAME
 $var.VM_Nom=$vm.Name
 $VmsDataArray.Add($var)
 }
 ##################################################################################
function New-Computer_Info()
{
  param ($var)
  $computer_info_d=Get-CimInstance -ClassName Win32_ComputerSystem

    $computer_info_d1=Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, ServicePackMajorVersion, OSArchitecture, CSName, WindowsDirectory,SerialNumber

  $Computer_Info = new-object PSObject
  $Computer_Info | add-member -type NoteProperty -Name Nom_Modele -Value $computer_info_d.Model
  $Computer_Info | add-member -type NoteProperty -Name Reference -Value (get-ciminstance win32_bios).SerialNumber
  $Computer_Info | add-member -type NoteProperty -Name Numero_de_serie -Value ((get-ciminstance win32_bios).SerialNumber)
  $Computer_Info | add-member -type NoteProperty -Name OS -Value $computer_info_d1.Caption
  $Computer_Info | add-member -type NoteProperty -Name OS_Build -Value $computer_info_d1.Version
  $Computer_Info | add-member -type NoteProperty -Name OS_Architecture -Value $computer_info_d1.OSArchitecture
  $Computer_Info | add-member -type NoteProperty -Name Ram -Value (Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % { "{0:N1} GB" -f ($_.sum / 1GB)})
  $Computer_Info | add-member -type NoteProperty -Name Date_Install -Value ([timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($(get-itemproperty 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion').InstallDate)))

  return $Computer_Info
}



$Computer_Info=New-Computer_Info
$computerRam=Get-WmiObject Win32_PhysicalMemory | select DeviceLocator, @{Name="Capacity";Expression={ "{0:N1} GB" -f ($_.Capacity / 1GB)}}, ConfiguredClockSpeed, ConfiguredVoltage,SerialNumber,Manufacturer 
$computerSystem = Get-CimInstance CIM_ComputerSystem
$computerBIOS = Get-CimInstance CIM_BIOSElement
$computerOs=Get-WmiObject win32_operatingsystem | select Caption, CSName, Version, @{Name="InstallDate";Expression={([WMI]'').ConvertToDateTime($_.InstallDate)}} , @{Name="LastBootUpTime";Expression={([WMI]'').ConvertToDateTime($_.LastBootUpTime)}}, @{Name="LocalDateTime";Expression={([WMI]'').ConvertToDateTime($_.LocalDateTime)}}, CurrentTimeZone, CountryCode, OSLanguage, SerialNumber, WindowsDirectory 
$computerCpu=Get-WmiObject Win32_Processor | select DeviceID, Name, Caption, Manufacturer, MaxClockSpeed, L2CacheSize, L2CacheSpeed, L3CacheSize, L3CacheSpeed 
$computerMainboard=Get-WmiObject Win32_BaseBoard | Select-Object Manufacturer,Model, Name,SerialNumber,SKU,Version,Product
# Get HDDs
$driveType = @{
   2="Removable disk "
   3="Fixed local disk "
   4="Network disk "
   5="Compact disk "}
$Hdds = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | select DeviceID, VolumeName, @{Name="DriveType";Expression={$driveType.item([int]$_.DriveType)}}, FileSystem,VolumeSerialNumber,@{Name="Size_GB";Expression={"{0:N1} GB" -f ($_.Size / 1Gb)}}, @{Name="FreeSpace_GB";Expression={"{0:N1} GB" -f ($_.FreeSpace / 1Gb)}}, @{Name="FreeSpace_percent";Expression={"{0:N1}%" -f ((100 / ($_.Size / $_.FreeSpace)))}} 
$OS_Name=(Get-CimInstance  -class win32_operatingsystem | select Caption, BuildNumber).Caption
$OS_Build=(Get-CimInstance  -class win32_operatingsystem | select Caption, BuildNumber).BuildNumber
$model=(Get-CimInstance -ClassName Win32_ComputerSystem |select Model).Model
$CPU=(Get-CimInstance -ClassName Win32_Processor ).Name
$RamData=Get-CimInstance win32_physicalmemory | select Manufacturer,Banklabel,Configuredclockspeed,Devicelocator,Capacity,Serialnumber | ConvertTo-Html -Fragment

Write-Host "Gathering Report Customization..." -ForegroundColor White

Write-Host "__________________________________" -ForegroundColor White


<###########################
         Dashboard
############################>

#Check for ReportHTML Module
$Mod = Get-Module -ListAvailable -Name "ReportHTML"

If ($null -eq $Mod)
{
	
	Write-Host "ReportHTML Module is not present, attempting to install it"
	
	Install-Module -Name ReportHTML -Force
    import-module C:\install\reporthtml
	Import-Module ReportHTML -ErrorAction SilentlyContinue
}
mkdir "C:\Script-Prep\"
$CompanyLogo = "https://cdn.logo.com/hotlink-ok/logo-social.png"
$RightLogo = "https://cdn.logo.com/hotlink-ok/logo-social.png"
$ReportSavePath = "C:\Script-Prep\"

Write-Host "Working on Dashboard Report..." -ForegroundColor Green

Write-Host "Done!" -ForegroundColor White

$tabarray = @('Host','VM', "Veeam")

Write-Host "Compiling Report..." -ForegroundColor Green

#Dashboard Report
$FinalReport = New-Object 'System.Collections.Generic.List[System.Object]'
$FinalReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo))
$FinalReport.Add($(Get-HTMLTabHeader -TabNames $tabarray))
$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format dd-MM-yyyy))))


$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Caracteristique systeme"))
$FinalReport.Add($(Get-HTMLContentTable $Computer_Info))
$FinalReport.Add($(Get-HTMLContentClose))

$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Caracteristique HDD, Ram"))
$FinalReport.Add($(Get-HTMLColumn1of2))
$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Ram Info'))
$FinalReport.Add($(Get-HTMLContentDataTable $computerRam -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLColumnClose))
$FinalReport.Add($(Get-HTMLColumn2of2))
$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Disque dure'))
$FinalReport.Add($(Get-HTMLContentDataTable $Hdds -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLColumnClose))
$FinalReport.Add($(Get-HTMLContentClose))
#---------------------------------------------------

$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "CPU and Carte mere"))
$FinalReport.Add($(Get-HTMLColumn1of2))
$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "CPU"))
$FinalReport.Add($(Get-HTMLContentDataTable $computerCpu -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLColumnClose))
$FinalReport.Add($(Get-HTMLColumn2of2))
$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Carte mere'))
$FinalReport.Add($(Get-HTMLContentDataTable $computerMainboard -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLColumnClose))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLTabContentClose))




#Groups Report
$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format dd-MM-yyyy))))


$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'VM info'))
$FinalReport.Add($(Get-HTMLContentTable ($VmsDataArray) -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 3 -HeaderText 'VM Disque'))
$FinalReport.Add($(Get-HTMLContentTable ($VmsDataArrayDisque) -HideFooter))
$FinalReport.Add($(Get-HTMLContentClose))
$FinalReport.Add($(Get-HTMLTabContentClose))


$FinalReport.Add($(Get-HTMLClosePage))


$Day = (Get-Date).Day
$Month = (Get-Date).Month
$Year = (Get-Date).Year
$Hour = (Get-Date).Hour
$Minute = (Get-Date).Minute

$ReportName = ("$Year-$Month-$day-$Hour-$Minute-$env:COMPUTERNAME")
mkdir -Path "C:\Script-Prep" -ErrorAction SilentlyContinue

Save-HTMLReport -ReportContent $FinalReport -ShowReport -ReportName "$ReportName.html" -ReportPath "C:\Script-Prep" 