#========================================================================
# Code Generated By: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.22
# Generated On: 29-08-2013 18:36
# Generated By:  
# Organization:  
#========================================================================

	#==================| Satnaam Waheguru Ji |===============================
	#           
	#            Author  :     Aman Dhally 
	#            E-Mail  :      amandhally@gmail.com 
	#            website :    	 www.amandhally.net 
	#            twitter :       @AmanDhally 
	#            blog    :       http://newdelhipowershellusergroup.blogspot.in/
	#            facebook:   	http://www.facebook.com/groups/254997707860848/ 
	#            Linkedin:    	 http://www.linkedin.com/profile/view?id=23651495 
	# 
	#            Creation Date    : 12-08-2013 
	#            File    :  		  OutlookTools.psm1
	#            Purpose : 
	#            Version : 1 
	#			Updated On :          14-08-2013
	#
	#            My Pet Spider :          /^(o.o)^\  
	#========================================================================
	
try {
	Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
	$folders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
	$outlookApplication = New-Object -ComObject 'Outlook.Application' -ErrorAction 'Stop'
	
	$mapi = $outlookApplication.GetNameSpace("mapi")
	$tasks = $mapi.getDefaultFolder($folders::olFolderTasks).Items
	
	Write-Verbose -Message "Outlook Object Created."
	}
catch {
	Write-Warning -Message "Not able to create an Outlook Object,function won't work, please check if you have Outlook.Application com object installed."
}
	
	
	function New-OutlookCalendarMeeting {
	
		
	<#
		.SYNOPSIS
			Create Outlook Calendar Meetings using Powershell Console.
	
		.DESCRIPTION
			If you are a Powershell Scripter or Programmer, then most of your time is spent
			On the Powershell Console. I want to write a small function which helps me to
			Create a calendar invites from the Powershell console. So that I can add calendar
			Invites on the fly and add them as reminder.
	
		.PARAMETER  Subject
			Using -Subject parameter please provide the subject of the calendar meeting.
	
		.PARAMETER  Body
			Using -Body, you can add a more information in to the calendar invite.
	
		.PARAMETER  Location
			The location of your Meeting, for example can be meeting room1 or any country.
	
		.PARAMETER  Importance
			By Default the importance is set to 2 which is normal
			You can set to -Importance high by providing 2 as an argument
	    	0 = Low
	    	1 = Normal
	    	2 = High.
	
	
		.PARAMETER  AllDayEvent
			If you want to create an all day event mart it as $true.
	
	    .PARAMETER BusyStatus
	        To set your status to Busy, Free Tenative, or out of office, By default it is set to Busy
	        0 = Free
	        1 = Tentative
	        2 = Busy
	        3 = Out of Office
	
	
		.PARAMETER  EnableReminder
			By Default reminders are enabled. If you don’t want to enable Reminder set it to $false.
	
		.PARAMETER  MeetingStart
			Provide the Date and time of meeting to start from.
	
		.PARAMETER  MeetingDuration
			By default meeting duration is set to 30 Minutes. You can change the duration Of the meeting using -MeetingDuration Parameter.
	
		.PARAMETER  Reminder
			'By default you got reminder before 15 minutes of meting starts. 
	         You can use -Reminder to set the reminder duration. The value is in Minutes.'
	
	
		.EXAMPLE
			PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -AllDayEvent $true -EnableReminder $false
			
	
		.EXAMPLE
			PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -MeetingStart "08/08/2013 22:30" -Reminder 30 
		
	
		.EXAMPLE
			PS C:\>New-CalendarMeeting -Subject "Powershell" -Body "Show how to use Powershell with Outlook" -Location "Conf Room 1" -Importance 2 
	
	
		.NOTES
	        I worte this function for adding a quick calender meeting.
	        in this fucntion you can't add anyone and sent invites to someone.
	        In next version of the same function , i will add these functionality.
	        Thanks : Aman Dhally {amandhally@gmail.com}
	
		.LINK
			www.amandhally.net
	
		.LINK
			http://newdelhipowershellusergroup.blogspot.in/
	#>
		
	[cmdletBinding()]	
	param (
	
	
	
		# Subject Parameter	
	    [Parameter(
	        Mandatory = $True,
	        HelpMessage="Please provide a subject of your calendar invite.")]
	    [Alias('sub')]
	    [string] $Subject,
	
		#Body parameter
	    [Parameter(
	        Mandatory = $True,
	        HelpMessage="Please provide a description of your calendar invite.")]
	    [Alias('bod')]
	    [string] $Body,
	
		#Location Parameter
	    [Parameter(
	        Mandatory = $True,
	        HelpMessage="Please provide the location of your meeting.")]
	
	    [Alias('loc')]
	    [string] $Location,
	
		# Importance Parameter
		[int] $Importance = 1,
	
		# All Day event Parameter
		[bool] $AllDayEvent = $false,
	
		# Set Reminder Parameter
		[bool] $EnableReminder = $True,
	
		# Busy Status Parameter
		[string] $BusyStatus = 2,
	
		# Metting Start Time Parameter
		[datetime] $MeetingStart =(Get-Date),
	
		# Meeting time duration parameter
		[int] $MeetingDuration = 30, 
	
		# Meeting time End parameter
			#[datetime] $MeetingEnd = (Get-Date).AddMinutes(+30),
	
		# by Default Reminder Duration
		[int] $Reminder = 15
	
	
	
	)
	
	BEGIN { 
	        
	        Write-Verbose "Creating Outlook as an Object"
	        
	
	        # Creating a instatance of Calenders
	        $newCalenderItem = $outlookApplication.CreateItem('olAppointmentItem')
	
	
	
	      }
	
	
	PROCESS { 
	        
	         Write-Verbose "Creating Calender Invite"
	    
	         $newCalenderItem.AllDayEvent = $AllDayEvent
	         $newCalenderItem.Subject = $Subject
	         $newCalenderItem.Body = $Body
	         $newCalenderItem.Location  = $Location
	         $newCalenderItem.ReminderSet = $EnableReminder
	         $newCalenderItem.Importance = $importance
	
	
	         if ( ! ($AllDayEvent)) {
	
	         $newCalenderItem.Start = $MeetingStart
	         $newCalenderItem.Duration = $MeetingDuration
	         
	         }
	         
	         $newCalenderItem.ReminderMinutesBeforeStart = $Reminder
	         # 2 is busy, 3 is ou to office
	         $newCalenderItem.BusyStatus = $BusyStatus
	             
	    }
	
	END {
	    
	        Write-Verbose "Saving Calender Item"
	        $newCalenderItem.Save()
	      
	        # if you want to see the calener invite un-comment the below line
	            #un-comment it ==>  $newCalenderItem.Display($True)
	
	       }
	
		}
	
	
	
	function New-OutlookContact {
	
	
	<#
	
	.SYNOPSIS
		Create Outlook contacts using Powershell console.
	
	
		
		
	.PARAMETER FirstName
		This is a mandatory prameter, You have to provide a first Name to 
		add a contact.
	
	.PARAMETER Birthday
		To add a birthdate of the new contact use this parameter, and this parameter
		accept MM/DD/YY format.
		
	.PARAMETER	BusinessPhone
		To add a Business phone of the conact use this parameter.
		
	.PARAMETER	Company
		To add contact's company name use this parameter.
		
	.PARAMETER	EmailAddress
		To add contact's email-ID use this parameter.
		
	.PARAMETER	HomeAddress
		To add  contact's home address use this parameter.
		
	.PARAMETER	JobTitle
		To add new contact's JOB title use this parameter.
		
	.PARAMETER	LastName
		To add new contact's last name use this parameter.
		
	.PARAMETER	MobileNumber
		To add new contact mobile numner use this parameter.
		
	.PARAMETER	Notes
		To add extra notes/description use this parameter.
		
	.PARAMETER	Website
	
	.NOTES
	        
		Using this Function, you can add an Outlook Contact on the fly. 
		I converted it to the module. So that i can start adding more functionality to 
		it.
	
		
		
	.EXAMPLE
		New-OutlookContact -FirstName "jhonson" -LastName "Smith"
		
		
	.EXAMPLE
		New-OutlookContact -FirstName "Ajit" -LastName "Singh" -MobileNumber "9910129889"
		
		
	.EXAMPLE	
		New-OutlookContact -FirstName "Jujhar" -LastName "Singh" -Website "www.js.com" -Company "Jujhar Studio"
		
		
	.EXAMPLE	
		New-OutlookContact -FirstName "Jorawar" -LastName "Singh" -BusinessAddress "Sarhind, Punjab" -JobTitle "Warrior" -Notes "I saw him in a Sikh Temple"
		
	.EXAMPLE	
		New-OutlookContact -FirstName "Fateh" -LastName "Singh" -Birthday "05/01/1999" -BusinessPhone "0999999" -BusinessAddress 'Sarhind, Punjab' -Company "Khalsa" -EmailAddress 'fateh@khalasa.com' -HomeAddress 'Talwandi,Punjab' -JobTitle 'warror' -MobileNumber "90909090'
		
		
	.LINK
		www.amandhally.net
	
	.LINK
		http://newdelhipowershellusergroup.blogspot.in/
	
	#>
	
	[cmdletBinding()]	
	
	param(
	
	
	#a
	#b
	[Alias('bday')]
	[datetime]$Birthday,
	
	[Alias('bisph')]
	[string]$BusinessPhone, 
	
	[Alias('bisadd')]	
	[string]$BusinessAddress,
	#c
	[Alias('comp')]	
	[string]$Company,
	#d
	#e
	[Alias('email')]	
	[string]$EmailAddress,
	#f
	[Parameter(
		Mandatory = $True,
		HelpMessage="Please enter a First name of the person.")]
	[Alias('first')]	
	[string]$FirstName,
	#g
	#h
	[Alias('homead')]	
	[string]$HomeAddress,
	#i
	#j
	[Alias('Jobti')]	
	[string]$JobTitle,
	#k
	#l
	[Alias('last')]	
	[string]$LastName,
	#m
	[Alias('mobile')]	
	[string]$MobileNumber,
	#n
	[Alias('note')]	
	[string]$Notes,
	#o
	#p
	#q
	#r
	#s
	#t
	#u
	#v
	#w
	[Alias('web')]	
	[string]$Website
	#x
	#y
	#z
	
	) # enf d paramaters
	 
	
	
	
	BEGIN {
	
	    try 
	        {
	        
	        $contactObject = $outlookApplication.CreateItem('olContactItem')
	        }
	    
	    catch
	        {  
	        Write-Warning '$_.exception occured'
	        }
	
	
	
	
	
	} # end of begin
	
	
	PROCESS 
	{
	
	
	    try 
	        {
	        
	        $contactObject.FirstName = $FirstName
	        $contactObject.LastName  =$LastName
	        $contactObject.MobileTelephoneNumber = $MobileNumber
	        $contactObject.Email1Address = $EmailAddress
	        $contactObject.WebPage = $Website
	        $contactObject.Body = $Notes
	        $contactObject.CompanyName = $Company
	        $contactObject.JobTitle = $JobTitle  
	        if ( ! ( $Birthday -eq $null)){    
	        $contactObject.Birthday = $Birthday
	        }
	        $contactObject.BusinessTelephoneNumber = $BusinessPhone
	        $contactObject.BusinessAddress = $BusinessAddress
	        $contactObject.HomeAddress = $HomeAddres
			
	        
	
	        }
	
	    catch 
	        {
	        
	         Write-Warning '$_.exception occured'
	
	        } 
	
	
	
	} # end of process
	
	
	END 
	{
	
	    try 
	        {
	        $contactObject.Save()
			Write-Host "Contact $FirstName  created" -ForegroundColor 'Yellow'	
	        
	        }
	            
	
	    catch
	        {
	        Write-Warning '$_.exception occured'
	        }
	
	
	
	
	
	} # end of end
	
	
	
	
	
	} # end of the New-OutlookContact functions
	
	
	
	function New-OutlookTask {
	
	
	<#
		.SYNOPSIS
			This function creates a new task in Ms Outlook
		
		.DESCRIPTION
			By using this function, we can create a new task using Powershell console.
		
		
	
		.PARAMETER Subject
			This parameter is mandatory. We need to provide a subject for a task.
		
		.PARAMETER Notes
			If you want to add extra notes to new task you can use this prameter.
		
		.PARAMETER DueDate
			Due date parameter is used to set a due date for the task.
		
		.PARAMETER EnableReminder
			You can either enable or disable reminder, by default it is set to true.
		
		.PARAMETER StartDate
			To set the start date you can use this parameter.
		
		.PARAMETER  Importance
			By Default the importance is set to 2 which is normal
			You can set to -Importance high by providing 2 as an argument
	    	0 = Low
	    	1 = Normal
	    	2 = High.
		
		.PARAMETER ReminderTime
			To get a reminder of the task on a particular day and time, use this option.
	
		.EXAMPLE
		    New-OutlookTask -Subject "Going to visit Alice in wonderland"
	
	    .EXAMPLE
			New-OutlookTask -Subject "Visit Captain Jack Sparrow" -Notes " He lives near Caribbean"
	
	    .EXAMPLE
			New-OutlookTask -Subject "Metting with Mr. Harrry Potter" -Importance 2 -StartDate "08/15/2013 19:00 "
	
	    .EXAMPLE
			New-OutlookTask -Subject "Meeting with Tony" -Notes "ask him about Startreak" -EnableReminder $false
		
		.EXAMPLE
			New-OutlookTask -Subject "Meeting with Tony" -Notes "ask him about Startreak" -EnableReminder $false -Importance
		
		.EXAMPLE
			New-OutlookTask -Subject "Search for Avengers" -Notes "I goona find them" -Importance 2 -StartDate "08/15/2013"-DueDate "08/19/2013" -ReminderTime "08/17/2013 10:30"
		
	    
		
		.LINK
			www.amandhally.net
	
		.LINK
			http://newdelhipowershellusergroup.blogspot.in/
		
		.Link
			amandhally@gmail.com
		
	#>
		
	[cmdletBinding()]	
	Param (
	
	# Subject parameter , this is mandatory one.	
	[Parameter(
		Mandatory = $True,
		HelpMessage="Please enter a subject for the task.")]
	[Alias('subj')]		
	[string]$Subject,
	
	# Note Parameter , to add extra notes	
	[Alias('note')]	
	[string]$Notes,
	
	# Dute Date Parameter	
	[Alias('ddate')]		
	[DateTime]$DueDate,
	
	# Enable/Disable reminders	
	[Alias('enablerem')]		
	[boolean]$EnableReminder = $true,
	
	# Start Date of the parameter	
	[Alias('stdate')]		
	[DateTime]$StartDate,
	
	# Importance of the task	
	[Alias('impor')]		
	[string]$Importance = 1,
	
	# Reminder Time, when should you get reminder.	
	[Alias('remtime')]		
	$ReminderTime
	
	)
	
	
	
	BEGIN 
		{
			$newTaskObject = $outlookApplication.CreateItem("olTaskItem")
		}
	
	
	PROCESS 
		{
		    $newTaskObject.Subject = $Subject
	        $newTaskObject.Body = $Notes
		    if ( !( $DueDate -eq $null)) {$newTaskObject.DueDate = $DueDate}
		    $newTaskObject.ReminderSet = $EnableReminder
		    if (! ($StartDate -eq $null)) { $newTaskObject.StartDate = $StartDate }
	        $newTaskObject.Importance = $Importance
		    if (! ($ReminderTime -eq $null)) {$newTaskObject.ReminderTime = $ReminderTime }
		}
	
	END {
			$newTaskObject.Save()
			Write-Host "Task => $Subject <= created fucessfully." -ForegroundColor Green
		}
	
	} # end of the New-OutlookTask Function.
	
	
	
	function New-OutlookNote {
	 
	 
	 <#
		.SYNOPSIS
			Create Outlook Note using Powershell Console.
	
		.DESCRIPTION
			Create a quick note in Outlook, using Powershell console.
	
		.PARAMETER  Title
			This is a mandatory parameter. Use this to create a title of the note.
	
		.PARAMETER  Body
			Using -Body, you can add a more information note.
	
		.PARAMETER  Color
			Use this parameter to create a custom color note.
			This property is in int value. 
			0 = Blue
			1 = Green
			2 = Red
			3 = Yellow
			4 = white
		
		.EXAMPLE
			PS C:\>New-OutlookNote -Title "Hello" -Body "Hello World"
	
		.EXAMPLE
			PS C:\>New-OutlookNote -Title "Brainstroming" -Body "need to purcashe some brain strom soft" -Color 2
		
		.LINK
			www.amandhally.net
	
		.LINK
			http://newdelhipowershellusergroup.blogspot.in/
	#>
		
	[cmdletBinding()] 
	
	param (
	
	 #title 
	[Parameter(
	        Mandatory = $True,
	        HelpMessage="Please provide title for a note.")]
	 [Alias('titl')]	
	 $Title,
	
	 #Body
	 [Alias('bod')]
	 $Body,
	
	 #color
	[Alias('col')]
	[int]$Color = 3 
	 
	 
	 )
	
	# end of parameters.
		
		BEGIN 
	      {
	  
	        $outlookApplication = New-Object -ComObject 'Outlook.Application'
			#Creating a note object.
	        $noteItem = $outlookApplication.CreateItem('olNoteItem')
	 
	
	      }
	 
	 
	 PROCESS 
	        {
	 		
			#Setting Note body.
	        $noteItem.Body = $Title + "`n" +"-------------" + "`n" +  $Body
	        
			#Setting Note Color.
			$noteItem.Color = $color
	        
	              
	        }
	
	
	 END 
	    {
	        # saving the note
	        $noteItem.Save()
			
			#writing a prompt message.
	        Write-Host "Note: $Title is credted." -ForegroundColor Yellow
	    
	    }
	 
	
	 
	 } # end of New-OutlookNote the function 
	
	
	 
	
	
	
	function New-OutlookEmail {
	
		
	<#
		.SYNOPSIS
		Using this Function You can send an email, and sent email will be saved in 
		the SentItems
		
		.DESCRIPTION
		A function to create and send email using Powetshell console
		
		.PARAMETER TO
		This parameter is mandatory. Provide the TO email address.
		
		.PARAMETER CC
		Use this parameter to carbon copy someone.
		
		.PARAMETER BCC
		Use this Parameter to add Blank Carbon Copy.
		
		.PARAMETER Subject
		Use this paramter to add Subject to the email.
		
		.PARAMETER StartDate
		To set the start date you can use this parameter.
		
		.PARAMETER  Importance
			By Default the importance is set to 2 which is normal
			You can set to -Importance high by providing 2 as an argument
	    	0 = Low
	    	1 = Normal
	    	2 = High.
		
		.PARAMETER Attachment
		Use this paramter to add attachedment to the email. as per now this support
		attaching one file only.
	
		.PARAMETER Body
		Use this parameter to write the body of the email.
		
		.EXAMPLE
		    New-OutlookTask -Subject "Going to visit Alice in wonderland"
	
	    .EXAMPLE
	
	    .EXAMPLE
	
	    .EXAMPLE
	    
		
		.LINK
			www.amandhally.net
	
		.LINK
			http://newdelhipowershellusergroup.blogspot.in/
		
		.Link
			amandhally@gmail.com
		
	#>
	
	[Cmdletbinding()]
	
	param(
	
		
	[Parameter(
		Mandatory = $True,
		HelpMessage="Please Enter the Recipient Email-ID.")]		
	[string]$To,
	
			
	[string]$Cc,
	
		
	[string]$Bcc,
	
	
	[Parameter(
		Mandatory = $True,
		HelpMessage="Please add the subject of the email.")]
	[Alias('subj')]		
	$Subject,
	
	[Alias('attc')]		
	$Attachment = $null,
	
	[Alias('bod')]		
	[string]$Body,
	
	[Alias('impo')]		
	$importance = 1
	
	
	
	)
	
	
	BEGIN 
	    {
			#Creating a note object.
	        $newMailObject  = $outlookApplication.CreateItem('olMailItem')
	    }
	
	
	PROCESS 
	    {
		
	    $newMailObject.To = $To
	    $newMailObject.CC = $Cc
	    $newMailObject.BCC = $Bcc
	    $newMailObject.Subject = $Subject
	    
		if ( ! ($Attachment -eq $null) ) {
		    $attch = $newMailObject.Attachments
	    	$attch.add("$Attachment")
	    
	    }
	    $newMailObject.Importance = $Importance
	    $newMailObject.HTMLBody  = $Body

	    }
	
	END 
	    {
			
	    $newMailObject.Send() 
			
		# if you want to save the email, unchecke the below #
		#$newMailObject.Save()
	    
	    }
	
	} # end of New-OutlookEmail function
	
	# START Get-OutlookTask function
	function Get-OutlookTask {
	[Cmdletbinding()]
	param()
	BEGIN {
		$subject = $tasks.Subject
		$importance = $tasks.Importance
		$StartDate = $tasks.StartDate
		
		$tasks | Select-Object Subject,Importance,StartDate,DueDate,Categories
		
		Write-Host "something"
	}
	PROCESS {
		
	}
	END {
		
	}
	
	}
	# END Get-OutlookTask function

	
	#-------------------------------------------------------------------------------
	
	# Exporting Module members
	Export-ModuleMember -alias * -Function *
