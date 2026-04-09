
# download azcopy so that you can copy files via azcopy
Invoke-WebRequest -Uri "https://aka.ms/downloadazcopy-v10-windows" -OutFile "azcopy.zip"
Expand-Archive azcopy.zip -DestinationPath "C:\AzCopy"
cd C:\AzCopy\azcopy_windows_amd64_*
.\azcopy.exe --version

# use azcopy to copy file question in version
.\azcopy.exe copy "<SAS URL>" "C:\Temp\"
