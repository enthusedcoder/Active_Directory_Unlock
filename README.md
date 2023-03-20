# Active_Directory_Unlock
## Why I wrote it

![My tool](http://i.imgur.com/uJxdtDA.png)

> I wrote this script while working at the corporate offices of Arbys Restaurant Group.  My job was to provide technical support for all the corporate store locations in America, and it was quite a challenge considering the resources (or the lack thereof) available to me and my co-workers.  So, the goal was to create a toolkit that would be able to significantly reduce the amount of time needed to perform certain processes that were 1) time consuming and 2) frequently performed.  I am very proud of the end result and used it very often during my time working at Arbys while continuing to add extra features as I thought of them.  Even a few of my coworkers, intrigued when they saw me using the tool, wanted to use it too, and I was more than happy to provide them the link to this GitHub page.  But without further ado, I will now go over the functionality of the tool.

## Prerequisites

> The tool requires powershell 5 to be installed on the machine running the tool.  Depending on your operating system, you may already have powershell 5 and its prerequisites installed, but if you are unsure, I included the Windows installer for .Net 4.6.2 and 4.7.2 (the most recent release) in the "required_installs" folder, which is a prerequisite for powershell 5.  As for installing powershell 5, I would recommend using [Chocolatey](https://chocolatey.org/install) to simplify this process.  All you have to do is open an elevated command prompt, which can be done by (if you have Windows 7) opening your start menu, typing "cmd" into the "Search programs" input, right clicking the only result, and selecting "run as administrator", or (if you have Windows 8 or above) by right clicking your start menu and selecting "Command Prompt (admin)".  You would then type each of the commands, one after the other **_WHILE PERFORMING RESTARTS AS NEEDED_** as indicated by the console output:
1. @powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
> This will install chocolatey

2. call "C:\ProgramData\chocolatey\redirects\RefreshEnv.cmd"
> This step is optional, as you can either run this command or simply close the command prompt window and reopen a new one for the same effect.

3. choco install powershell -y
> This will install the latest version of powershell on your device.  This will also require a restart.

> Arbys also uses Salesforce for its ticketing system and also stores important data that is needed by the script in order to perform the tasks it needs to.  So the final install is a set of Salesforce powershell cmdlets which allow powershell to link to salesforce (using a key which the user will need to obtain from salesforce and input into the script when first run) and query information stored in salesforce for use in the script.

## Using the tool

> When you first launch the tool (using the "launch.bat" script), a shortcut will be placed on your desktop which you will use to launch it in the future.  The tool will verify that you have the prerequisites required for its use, obtain information from the user such as Windows credentials (to identify the user as well as for Active Directory manipulation), the salesforce key needed for powershell integration, and the default password that will be used when resetting active directory users' passwords.

### Active Directory Integration

> What I really like about this utility is its ability to integrate seemlessly with Active Directory using a Directory Object Picker Dialog object, which is displayed below.  This allows for the user to, when the button indicated below is pressed, to be able to search through active directory to find the user whose account they need to modify, rather than typing it out and hoping they spell it right.  When a user object is selected in the dialog picker, the user name for that user object is populated into the input box and is ready to be manipulated.
![Shot1](http://i.imgur.com/AW5ibs0.png)![shot 2](http://i.imgur.com/EaKhlOz.png)![shot3](http://i.imgur.com/EWukn40.png)

### User account manipulation
> And now the user can select any of the buttons in the GUI to perform an action against that username.  And this is where the salesforce integration is necessary.  The policy concerning password resets for the General Managers of the stores is to send the reset password to the Area Supervisor who oversees the region in which that store is located for authentication purposes, and the only place to find that information was in salesforce under the information for a particular store.  So when the "Email area supervisor" button is clicked, the script will first determine if the username which is in the input field is a person's user account or a store's user account (yes, the stores had user accounts too for some reason.....  I'm so glad I don't work there anymore.)  Then, if the account was a person's account, the script would obtain the store number to which that user is assigned by reading it from the appropriate Active Directory attribute.  Then, the script would query the salesforce database using the appropriate store number and would obtain the email address of the area supervisor in whose district that store is located.  The script would then (using the windows credentials) create and send an email to the area supervisor with the new password and using information queried from Active Directory to tailor the email with all information needed.

### Contacting store ISP
> The last function that I want to draw attention to is the "Contact store ISP" function (ISP meaning Internet Service Provider).  Arbys used a solarwinds product to monitor the internet connectivity of all of their store locations, and each store was provided with standard, non-comercial (equipment used by average users in their homes) networking equipment in order to save money I guess.  The problem with this was that, due to the fact that the equipment was not in a home environment and was the cheapest of the cheap, the equipment would break frequently.  So after a technician walked the end user through performing layer one troubleshooting on the networking equipment, if internet was not restored, we would need to email one of the three different ISPs that Arbys used to provide internet to the store, and the solarwinds tool was the only place where that information could be obtained (seriously, I am sooooo glad I do not still work there).  So, using the tool, the user could use the account of the store that was experiencing the issue, or the account of the general manager who ran the store, obtain the light status of the modem device (which lights were on, were they blinking or solid, what color, etc.) and input it into the edit box indicated below, and when the "Email store ISP" button is clicked, will obtain the applicable store number from whichever user account is in the input box, automatically login to the solarwinds tool using the general password that everyone used, access the page for the affected store, determine the ISP of that store using the text returned from the web page, and then craft and send an email to the appropriate ISP containing all relevant information concerning the store, its location, the problem, the light status of the modem, etc.

![Lights](http://i.imgur.com/POXzCQa.png)

### Other miscellaneous tasks
> There are a few other tasks that this tool performs, but they are relatively simple, highly repetative tasks that really aren't worth going into detail about, like reseting the technician Aloha account password to the correct default by automatically entering 5 random passwords, then the correct default password (since the system restricts you from using any password that was one of the last 5 used, etc).


> Feel free to send me a message, either here or on LinkedIn and let me know what you think.
