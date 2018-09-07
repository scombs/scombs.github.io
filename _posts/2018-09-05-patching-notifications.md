---
layout: post
title: Monthly Patching Email Notifications
subtitle: utilize powershell to send simple reminders
bigimg: /img/landscape.jpg
tags: [PowerShell,Windows,Task Scheduler,Patching]
---

One of the first tasks I took on when I started at my new employer was to look to see where we could improve on current processes. The one that jumped out at me was a string of emails being sent twice a month letting the IT department know about patch weekends. These emails were sent manually on the week after patch Tuesday and also the following week to notify IT about the up coming patching weekends for Dev and Production. Attached to these emails were a list of all servers, Prod Lan and DMZ along with TST/Dev/QA/STG Lan and DMZ. These reports were also maintained and updated manually. 

As you can see I had my work cut out for me, all I had to do was get to work on a solution. I've worked with PowerShells Send-MailMessage a few times before but mostly for testing SMTP and sending static messages. This on the other hand was going to be a bit different due to the scheduling of the emails and the times they were sent out. The first notification would be sent out the week of patch Tuesday specifying the date of the upcoming weekend and also the following weekend for Production. That's where this little piece of code helped out. 


~~~
$BaseDate = ( Get-Date -Day 12 ).Date
$PatchTuesday = $BaseDate.AddDays( 2 - [int]$BaseDate.DayOfWeek )
~~~


Tim Curwick has a nice write up on how this works so I won't dive into it here but I'll link it at the bottom. Now that we have our base patch Tuesday we can calculate the other dates.


~~~
$datesat=$patchtuesday.adddays(11)
$datesat=$datesat.ToString('MM-dd-yyyy')

$datesun=$patchtuesday.adddays(12)
$datesun=$datesun.ToString('MM-dd-yyyy')

$datedev=$patchtuesday.adddays(5)
$datedev=$datedev.ToString('MM-dd-yyyy')
~~~


So I have my dates, next I had to determine which servers were going to be involved in patching. With the way our naming structure is a mix of new and legacy names it would be a nightmare to generate reports based on that. The next best thing that came to mind is using the AD structure and query servers based on OU. This works well here since each patching day has it's own OU and is divided up by site location. This may seem a bit messy but it works for day and location based patching. What we end up with are variables that are going to be our search base for that days patching. 


~~~
$searchbasesat ="ou=mgmt-sat,ou=dr servers,dc=MyDomain,dc=com",
"ou=mgmt-sat,ou=office1 servers,dc=MyDomain,dc=com",
"ou=mgmt-sat,ou=office2 servers,dc=MyDomain,dc=com",
"ou=mgmt-sat,ou=office3 servers,dc=MyDomain,dc=com"

$searchbasesun ="ou=mgmt-sun,ou=dr servers,dc=MyDomain,dc=com",
"ou=mgmt-sun,ou=office1 servers,dc=MyDomain,dc=com",
"ou=mgmt-sun,ou=office2 prod servers,dc=MyDomain,dc=com"

$searchbasedev = "ou=mgmt-sun,ou=office1 dev,dc=MyDomain,dc=com",
"ou=mgmt-sun,ou=office2 dev,dc=MyDomain,dc=com"
~~~


We have our dates and we have our OUs where the servers reside, next we need to query those OUs and generate a report. The simplest solution was passing each of these variables into foreach loops and outputting the names to a csv. This gives us a CSV for each day so the department can flip through it and also have a record of what was patched and when. 


~~~
#temp directory to hold the files
$loc = "c:\temp"

#3rd saturdays loop
$forloopsat = foreach ($base in $searchbasesat) {
    get-adcomputer -filter * -SearchBase $base | Select-Object name | Sort-Object Name 
}
$forloopsat | export-csv "$loc\ProdSatReboots.csv"

#3rd sundays loop
$forloopsun = foreach ($base in $searchbasesun) {
    get-adcomputer -filter * -SearchBase $base | Select-Object name | Sort-Object Name
}
$forloopsun | export-csv "$loc\ProdSunReboots.csv"

#2nd sunday dev loop
$forloopdev = foreach ($base in $searchbasedev) {
    get-adcomputer -filter * -SearchBase $base | Select-Object name | Sort-Object Name
}
$forloopdev | export-csv "$loc\DevSunReboots.csv"
~~~


We also have servers residing on a DMZ, we haven't talked much about that but here we'll need to include those servers in the list. I have a similar process that mirrors the above that sits on a DMZ box to query AD and then copies the CSVs to the job scheduler server. That process is set to run 15 minutes before this one which will then allow us to run the following code block to combine the appropriate csvs. 


~~~
@(import-csv "$loc\ProdSunReboots.csv") + @(import-csv "$loc\dmzfiles\dmzprodsunreboots.csv") | sort-object name | export-csv "$loc\Production-Sunday.csv"
@(import-csv "$loc\ProdSatReboots.csv") + @(import-csv "$loc\dmzfiles\dmzprodsatreboots.csv") | sort-object name | export-csv "$loc\Production-Saturday.csv"
@(import-csv "$loc\DevSunReboots.csv") + @(import-csv "$loc\dmzfiles\dmzdevsunreboots.csv") | sort-object name |export-csv "$loc\Development-Sunday.csv"
~~~


Our reports are generated and ready to go, which brings us to the finale, sending them to the department. Here's where Send-MailMessage comes into play. With this command we can utilize spatting to make our variables easy to define and also allows us to easily format our body using HTML.


~~~
$messageparameters = @{
subject="PLEASE REVIEW: Microsoft Security Updates for All Servers - Scheduled Reboots"
body ="
Good Morning,
<p>The Infrastructure team will patch all Development and Production Servers over the next two weekends to ensure they are fully compliant with Microsoft Security Updates.  This will require a REBOOT or possibly multiple reboots until all patches are fully installed.  Please reference the below patching schedule and attached files for a full list of servers being patched each day.<br>
<br>
<b> Production Servers</b> $datesat at 1:00 AM local time and $datesun at 1:00 AM local time.<br>
<br>
<b> Development Servers</b> $datedev at 1:00 AM local time.<br>
<br>
<b>REMINDER: **Please remember to test functionality for ALL Applications on Tst, Dev, QA, and Stg Servers next week prior to scheduled Production Server patching.**</b><br>
<br>
<p>Thank you for your support! <br>
<br>
</P>"
From = "noreply@corp.com"
TO = "IT@corp.com"
Smtpserver= "smtp.corp.com"
Attachments = "$loc\Development-Sunday.csv","$loc\Production-Saturday.csv","$loc\Production-Sunday.csv"
}
Send-MailMessage @messageparameters -BodyAsHtml
~~~


That's it, the PowerShell work is done. Now you can set this up on your job scheduler server to run on the day you want. For me I'm just using Windows Task Scheduler with a service account with the appropriate permissions. The benefit in using Windows Task Scheduler is that every Windows admin should know how to use it, this allows me to configure it and not worry about others being unable to manage jobs. This will handle your initial communication but I divided out the reminder since we don't need to do double duty of querying AD a second time. You'll be able to grab the primary and reminder scripts from my PowerShell repo. 

With this in place, it saves over an hour of work a month of manual intervention and it includes a nice piece of mind that things will just work. 

First meaningful blogging experience, first time using Jekyll, first time using markdown in a more meaningful way and the first time sharing my everyday PowerShell experience. Today's a good day. 

Links:

Tim Curwick - https://www.madwithpowershell.com/2014/10/calculating-patch-tuesday-with.html

My PowerShell Repo - https://github.com/scombs/PowerShell