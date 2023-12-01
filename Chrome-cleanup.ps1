#This scrip will look at the regesty and pull all of the current user profiles, Then it will look for a Chrome install in the user's appdata folder. If it's older than 90 days
#the chrome will be updated using the google update tool 

Set-ExecutionPolicy Bypass
#first we grab a list of all the users on the system
$userlist = Get-CimInstance -ClassName win32_userprofile |Select-Object localpath, lastusetime, Loaded

#then we parse out the user names and check to see how old they are
#due to a bug in Windows 10 post update 1709 when system updates are done it updates ALL NTUSER.DAT files. So you can't user that as a measuer of when they last logged in.
#so I use the IconCache.db file. Thanks to the folks on this thread for finding this. https://techcommunity.microsoft.com/t5/windows-deployment/issue-with-date-modified-for-ntuser-dat/m-p/102438

foreach ($user in $userlist) {
    if ($user.localpath.StartsWith("C:\Users\")) {
        $userArray = $user.localpath.Split("\")
        $userName = $userArray[2]
        #Check to see if the local install of chrome is present
        $fileExists = Test-Path "C:\Users\$userName\AppData\Local\Google\Chrome\Application\chrome.exe"
        if ($fileExists){
            #if it is figure out how old it is.
            $Dat = Get-Item "C:\Users\$userName\AppData\Local\Google\Chrome\Application\chrome.exe" -Force 
            $DatTime = $Dat.LastWriteTime
            $age = (Get-Date).adddays(-60)
            write-host $DatTime
            if ($DatTime -lt $age){
                $old = $true
                Write-Host $old
            }
            if ($old) {
                #If it's old, run the local google update to schedule an update. This should happen in the background and the user will be prompted to relunch when it's finished. 
            Invoke-Command -Scriptblock {"C:\Users\$userName\AppData\Local\Google\Update\GoogleUpdate.exe /ua /installsource scheduler"}
        }else {
            #if it's not old then say as much.
            Write-Host $userName" Chrome is up to date."
        }
        $old = $false
        }else{
            #if we can't find the local version then say it's on the globa install.
            Write-Output $userName"Is using the global chrome, no action taken."
        }
        Write-Host $old
        }
    }





