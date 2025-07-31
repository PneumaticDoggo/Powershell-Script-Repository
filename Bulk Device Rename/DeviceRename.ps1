$TypeDef = @" 
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
  
namespace Api
{
 public class Kernel32
 {
   [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   public static extern int OOBEComplete(ref int bIsOOBEComplete);
 }
}
"@ 
Add-Type -TypeDefinition $TypeDef -Language CSharp
  
$IsOOBEComplete = $false
$hr = [Api.Kernel32]::OOBEComplete([ref] $IsOOBEComplete)
if ($IsOOBEComplete) {
  Write-Host "Not in OOBE, nothing to do."
  exit 0
}
 
$systemEnclosure = Get-CimInstance -ClassName Win32_SystemEnclosure
$details = Get-ComputerInfo
 
if (($null -eq $systemEnclosure.SMBIOSAssetTag) -or ($systemEnclosure.SMBIOSAssetTag -eq "")) {
    if ($null -ne $details.BiosSerialNumber) {
        $assetTag = $details.BiosSerialNumber
    } else {
        $assetTag = $details.BiosSeralNumber
    }
} else {
    $assetTag = $systemEnclosure.SMBIOSAssetTag
}
if ($assetTag.Length -gt 13) {
    $assetTag = $assetTag.Substring(0, 13)
}
if ($details.CsPCSystemTypeEx -eq 1) {
    $newName = "Laptop-$assetTag"
} else {
    $newName = "Desktop-$assetTag"
}
 
if ($newName -ieq $details.CsName) {
    Write-Host "No need to rename computer, name is already set to $newName"
    Exit 0
}
 
Write-Host "Renaming computer to $($newName)"
Rename-Computer -NewName $newName -Force

#Change Prefix to whatever suits your needs $newName = "SOMEPREFIX-$assetTag" $newName = "ANOTHERPREFIX-$assetTag"