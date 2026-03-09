$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$cs  = Get-CimInstance Win32_ComputerSystem
$ram = Get-CimInstance Win32_PhysicalMemory
$gpu = Get-CimInstance Win32_VideoController | Sort-Object AdapterRAM -Descending | Select-Object -First 1
try {
   $storage = Get-PhysicalDisk | Sort-Object Size -Descending | ForEach-Object {
       "$($_.FriendlyName) - $($_.MediaType) - $([math]::Round($_.Size / 1GB)) GB"
   }
} catch {
   $storage = Get-CimInstance Win32_DiskDrive | Sort-Object Size -Descending | ForEach-Object {
       "$($_.Model) - $([math]::Round($_.Size / 1GB)) GB"
   }
}
$report = [pscustomobject]@{
   Computer     = $env:COMPUTERNAME
   Manufacturer = $cs.Manufacturer
   Model        = $cs.Model
   CPU          = $cpu.Name
   Cores        = $cpu.NumberOfCores
   Threads      = $cpu.NumberOfLogicalProcessors
   RAM_GB       = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
   RAM_Config   = ($ram | ForEach-Object { "$([math]::Round($_.Capacity / 1GB))GB@$($_.Speed)MHz" }) -join "; "
   GPU          = $gpu.Name
   GPU_RAM_GB   = $(if ($gpu.AdapterRAM) { [math]::Round($gpu.AdapterRAM / 1GB, 1) } else { "Unknown" })
   Storage      = $storage -join " | "
}
$path = "$env:USERPROFILE\Desktop\$env:COMPUTERNAME-specs.txt"
$report | Format-List | Tee-Object -FilePath $path