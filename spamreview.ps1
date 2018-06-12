#Specify the path of the SMTP logs
$Path = "C:\Program Files (x86)\Parallels\Plesk\Mail Servers\Mail Enable\Logging\SMTP"

#Specify the files you wish to scan
$FileType = "SMTP-Activity-*.log"

#Specify where you want the output to be logged
$OutputFile = "C:\temp\spamreview.log" #This is the location where it will print the results from the command.

#Get all lines from the log file(s) which contain "FROM:<"
$i = Get-ChildItem $Path -Filter $FileType|
        Get-Content |
            Select-String -Pattern 'FROM:<'

#Extract the email addresses from the prior results, get a count for each unique address, and print those results to a file
[Regex]::Matches($i, '(?<=\<).+?(?=\>)') | 
    %{$_.value} |
        Group-Object -NoElement |
            Sort-Object -Descending Count|
                FT -Auto |
		    Out-File -FilePath $OutputFile
