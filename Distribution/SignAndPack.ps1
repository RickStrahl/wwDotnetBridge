~\Dropbox\utl\signtool.exe sign /v /n "West Wind Technologies" /sm /s MY /tr "http://timestamp.digicert.com" /td SHA256 /fd SHA256 ".\wwdotnetbridge.dll"
~\Dropbox\utl\signtool.exe sign /v /n "West Wind Technologies" /sm /s MY /tr "http://timestamp.digicert.com" /td SHA256 /fd SHA256 ".\clrhost.dll"

~\Dropbox\utl\7z.exe  a -tzip -r .\wwDotnetBridge-Oss.zip .\wwdotnetbridge.dll .\clrhost.dll .\wwDotnetBridge.prg