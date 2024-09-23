###########################################################
# Project FoxPro Launcher
#########################
#
# This generic script loads FoxPro in the Deploy Folder and
# ensures that config.fpw is loaded to set the environment
#
# This file should live in any Project's Root Folder
###########################################################

$vfp = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\{002D2B10-C1FA-4193-B134-D86EAECC5250}\LocalServer32")."(Default)"
if ($vfp -eq $null)
{
   Write-Host "Visual FoxPro not found. Please register FoxPro with ``vfp9.exe /regserver``" -ForegroundColor Red
   exit
}

$vfp = $vfp.Replace(" /automation","")
$vfp

$path = [System.IO.Path]::GetFullPath("Examples")
$path

cd .\Examples
& $vfp -c"$path\config.fpw"

