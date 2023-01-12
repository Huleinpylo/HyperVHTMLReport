
$VmsDataArray =  New-Object -TypeName "System.Collections.ArrayList"
function New-VmsData()
{
  param ($var)


  $VmsDataVar = new-object PSObject
  $VmsDataVar | add-member -type NoteProperty -Name HOSTNAME -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name VM_Nom -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name vCPU -Value ""
  $VmsDataVar | add-member -type NoteProperty -Name vRam -value ""
  $VmsDataVar | add-member -type NoteProperty -Name IPAddresses -value ""
  $VmsDataVar | add-member -type NoteProperty -Name Disque -value (New-Object -TypeName "System.Collections.ArrayList")
  return $VmsDataVar
}
function New-VHD()
{
  param ($var)


  $VmsDataVar = new-object PSObject
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
    $vihd.VhdFormat= ($vmHDDVHD.VhdFormat)
    $vihd.VhdType= ($vmHDDVHD.VhdType)
    $vihd.Max_Size = ($vmHDDVHD.Size / 1GB).ToString() 
    $vihd.FileSize= ($vmHDDVHD.FileSize / 1GB).ToString()
    $var.Disque.Add( $vihd)
  } else {
    Write-Error 'Disk Warning Error accessing virtual disk'  
  }
}
$vmNetCards = (Get-VMNetworkAdapter -VMName $vm.Name)
 $var.IPAddresses = $vmNetCard.IPAddresses
 $var.HOSTNAME = $env:COMPUTERNAME
 $var.VM_Nom=$vm.Name
 $VmsDataArray.Add($var)
 }
