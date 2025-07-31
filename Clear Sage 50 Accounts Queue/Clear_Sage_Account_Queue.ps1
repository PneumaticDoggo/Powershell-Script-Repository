Taskkill /IM "sg50svc.exe" /f

Remove-Item "C:\Accounts\*ENTER SAGE COMPANY FOLDER NAME*\ACCDATA\QUEUE.DTA"

sleep -seconds 5

Start-Service -Name "Sage 50 Accounts Service v29"