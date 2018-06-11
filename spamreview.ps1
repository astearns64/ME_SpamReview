$Path = "C:\Program Files (x86)\Parallels\Plesk\Mail Servers\Mail Enable\Logging\SMTP"
$FileType = "SMTP-Activity-*.log"
$OutputFile = "C:\temp\spamreview.txt" #This is the location where it will print the results from the command.
$i = Get-ChildItem $Path -Filter $FileType|
        Get-Content |
            Select-String -Pattern 'FROM:<'
                
[Regex]::Matches($i, '(?<=\<).+?(?=\>)') | 
    %{$_.value} |
        Group-Object -NoElement |
            Sort-Object -Descending Count|
                FT -Auto |
		    Out-File -FilePath $OutputFile
