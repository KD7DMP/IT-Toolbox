#This scrip will look at the regesty and pull all of the current user profiles, and the last date which they used the profile. 
#it will then mark the oldest profiles and give you the option of removing them from the system.

Set-ExecutionPolicy Bypass
#first we grab a list of all the users on the system
$userlist = Get-CimInstance -ClassName win32_userprofile |Select-Object localpath, lastusetime, Loaded
$old = $false
#then we parse out the user names and check to see how old they are
#due to a bug in Windows 10 post update 1709 when system updates are done it updates ALL NTUSER.DAT files. So you can't user that as a measuer of when they last logged in.
#so I use the IconCache.db file. Thanks to the folks on this thread for finding this. https://techcommunity.microsoft.com/t5/windows-deployment/issue-with-date-modified-for-ntuser-dat/m-p/102438

foreach ($user in $userlist) {
    if ($user.localpath.StartsWith("C:\Users\")) {
        $userArray = $user.localpath.Split("\")
        $userName = $userArray[2]
        $fileExists = Test-Path "C:\Users\$userName\AppData\Local\IconCache.db"
        if ($fileExists){
            $Dat = Get-Item "C:\Users\$userName\AppData\Local\IconCache.db" -Force
        }else{
            $Dat = Get-Item "C:\Users\$userName\AppData\Local\Microsoft\Teams\" -Force
        }
            $DatTime = $Dat.LastWriteTime
            $age = (Get-Date).adddays(-375)
            if ($DatTime -lt $age){
                $old = $true
                Write-Host $old
            }
        
        #Now we remove this user from the system We make sure to ask the user before deleting a user just to be sure.
        if ($user.loaded -eq $false -and $old){
                $answer = Read-Host $userName, $DatTime, "Do you want to Delete this user? yes/no"
                
                if ($answer -eq 'yes'){
                    Get-WmiObject -class win32_userprofile | Where localpath -eq $user.localpath |Remove-WmiObject -Verbose
                    Write-Host "User Deleted"
                }else{
                    Write-Host $userName, $DatTime, "No Action, User Skiped"
                }           
        }
        Elseif($user.loaded -and $old){
            Write-Host $userName, $DatTime, "This is old, just still loaded. Leave it alone"
        }Elseif($userName -eq "pcmaint"){
            Write-Host $userName, $DatTime,"We Will leave pcmaint be"
        }
        else{
        Write-Host $userName, $DatTime, $old
        }
    }
}




