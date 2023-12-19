#####################################################################################################################
#Webgraphy
# https://learn-powershell.net/2015/01/24/creating-a-table-in-word-using-powershell/ - How to create tables in PS
# https://social.msdn.microsoft.com/Forums/en-US/18167a67-4bf7-4c1c-bc34-551ccbeecb97/powershell-word-table-cell-merge?forum=worddev - How to Merge Cells
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_hash_tables?view=powershell-7.3 - Hashes manual
# https://learn.microsoft.com/en-us/office/vba/api/word.wdcolorindex - list of colors
# https://stackoverflow.com/questions/75029186/powershell-insert-image-into-specific-position-in-a-word-document - how to insert an image in word using PS
# https://stackoverflow.com/questions/22925135/is-there-a-way-to-convert-a-powershell-string-into-a-hashtable - convert string to a hash
# https://stackoverflow.com/questions/4396771/how-to-append-to-powershell-hashtable-value - append data to a hash
# https://ss64.com/ps/substring.html - how to use substring
# https://theitbros.com/powershell-gui-for-scripts/ - how to create a PS GUI
# https://steemit.com/utopian-io/@cha0s0000/use-powershell-to-convert-word-file-to-pdf-file - how to export word to PDF
#####################################################################################################################
#This work is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
#To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/4.0/.
#####################################################################################################################
#File Created By Héctor Rodríguez Fusté
#SysAdmin / OS & Software Engineer
#####################################################################################################################
#Script to Create Sox Form and Access Form for Team Members
#####################################################################################################################
#Version: 1 - Created on 19/01/2017 in Python
#Forked to Powershell on 03/05/2023 by Héctor Rodríguez Fusté

#####################################################################################################################
#####################################################################################################################

#Functions
function WriteLog{ #Function to write log file for error debugging.
    Param ([string]$LogString) #Function Parameters
    
    $Stamp = (Get-Date).toString("dd/MM/yyyy HH:mm:ss") #TimeStamp for log file
    $LogMessage = "$Stamp $LogString" #Line added into log
    Add-content $LogFile -value $LogMessage #Writing errors into log file
}

function CreateFolder { #Function to create folders, it will check if the folder exists, if so it won't create it
    param ([string]$Folder)

    if(!(Test-path $Folder)){ #Checking if the folder  exists
        New-Item $Folder -ItemType Directory #Creating folder
    }
}

function Open-File { #Function to open a file in the machine, it will take only the file path but user will choose it visually.
    param ([string]$FileType)
    
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter = $FileType
    }
    $null = $FileBrowser.ShowDialog()

    return $FileBrowser.FileName
}

function Capitalize { #Function to capitalize an string as PS doesn't allow to do it in a easy way
    param ([string]$String)

    return $string.replace($string[0],$string[0].tostring().toupper())
}

function ReportLog { #Open the log file after the script have failed to show errors and let the user knows what happened
    # param ([string]$Logfile)
    
    start-process notepad.exe $Logfile
}

function SavingStat { #Saving PID from process given to know what was open in that moment, ex: if you have an excel file opened, after you open another excel process, you will be able to close only the new process, because you had the state saved
    param([string]$ProcessName)

   return get-process | where-object {$_.ProcessName -like "*$($ProcessName)"} | select-object -expandproperty id
}

function KillProcess { #Process to kill only the process opened between Saved States. Ex: This script avoid the closure of other apps that were

    param(
        [Parameter(position=0)][object] $StateBefore,
        [Parameter(position=1)][object] $StateNow
    )

    if($StateBefore.length -eq 0){
        taskkill /pid $StateNow /f | out-null
    }else{

        $id = compare-object $StateBefore $StateNow | select-object -ExpandProperty InputObject #comparing states, it gives you the process opened after the first state
        
        if($id.count -gt 1){
            foreach($x in $id){
                taskkill /pid $x /f | out-null
            }
        }else{
            taskkill /pid $id /f | out-null
        }
    }
}

function OpenWordFile { 
    param([string] $WordFile)

    #Opening Word File
    try{
        $StateBefore = SavingStat("Word")
        $Word = New-Object -ComObject Word.Application #Calling Word app
        $StateNow = SavingStat("Word")
        $Word.Visible = $False
        # $Word.Visible = $true # Visibility of Word app when script is running
        $DocW = $Word.Documents.add($WordFile)
        $Selection = $Word.Selection

        return $StateBefore,$StateNow,$Word,$DocW,$Selection
    }catch [System.Runtime.InteropServices.ComException] {
        WriteLog "ERROR;WORD_OPENING: $($PSItem)" #Error code
        KillProcess $StateBefore $StateNow
        ReportLog
    }   
}

function SaveAsPDF($WordObject,$WordFile,$WordPath){

    $WordFile.SaveAs([ref] $WordPath.replace(".docx",".pdf"), [ref] 17)
}

function SendEmail{
    param (
        [parameter(Mandatory=$true, position = 0)][string] $From,
        [parameter(Mandatory=$true, position = 1)][string] $To,
        [parameter(Mandatory=$false, position = 2)][string] $cc,
        [parameter(Mandatory=$true, position = 3)][string] $Subject,
        [parameter(Mandatory=$false, position = 4)][object] $Attach,
        [parameter(Mandatory=$true, position = 5)][object] $Body
    )

     $FromEmailList = $From
     $ToEmailList = $To
     $CCEmailList = $cc
     $EmailSubject = $Subject
    
    try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0)
            $Mail.From = $($FromEmailList)
            $Mail.To = $($ToEmailList)
            $Mail.CC = $($CCEmailList)
            $Mail.Subject = $EmailSubject
            $Mail.Attachments.add($Attach)
            $Mail.Body = $Body
            $Mail.save()

            $Inspector = $Mail.GetInspector
            $Inspector.display()

        }
        catch {
            WriteLog "ERROR;EMAIL_SEND: $($PSItem)" #Error code
            ReportLog
        }
}

function ConvertToHash{
    param([string]$string)

    $hash = @{}

    foreach($setting in $string.split(";")){
        $data = ConvertFrom-StringData -StringData $setting
        $hash.add("$($data.keys)","$($data.values)")
    }
    return $hash
}

function CreateUserName {
    param (
        [parameter(mandatory=$true, position = 0)][string]$name,
        [parameter(mandatory=$true, position = 0)][string]$lastname
    )
    $countLastName = $lastname.split(" ").count

    if($name.indexof(" ") -ne -1){

        $name = Capitalize("$($name.split(" ")[0])")
    }else{
        $name = Capitalize("$($name)")
    }
    
    if($lastname.indexof(" ") -ne -1){
        switch($lastname.split(" ").count){
            {$_ -eq 2}{
                $secondlastname = Capitalize("$($lastname.split(" ")[1])")
                $lastname = Capitalize("$($lastname.split(" ")[0])")
            }
            {$_ -eq 3}{
                $thirdlastname = Capitalize("$($lastname.split(" ")[2])")
                $secondlastname = Capitalize("$($lastname.split(" ")[1])")
                $lastname = Capitalize("$($lastname.split(" ")[0])")
            }
            Default {
                $lastname = Capitalize("$($lastname)")
            }
        }
    }else{
        $lastname = Capitalize("$($lastname)")
    }

    switch($countLastName){
        {$_ -eq 1}{
            $username = "{0}{1}" -f $name,$lastname

            if($username.Length -gt 15){
                $username = $username.substring(0,15)
            }
        }
        {$_ -eq 2}{
            $username = "{0}{1}" -f $name,$lastname

            if($username.Length -gt 15){
                $username = "{0}{1}" -f $name,$secondlastname

                if($username.Length -gt 15){
                    $username = "{0}{1}" -f $name,$lastname
                    $username = $username.substring(0,15)
                }
            }
        }
        {$_ -eq 3}{
            $username = "{0}{1}" -f $name,$lastname

            if($username.Length -gt 15){
                $username = "{0}{1}" -f $name,$secondlastname

                if($username.Length -gt 15){
                    $username = "{0}{1}" -f $name,$thirdlastname

                    if($username.Length -gt 15){
                        $username = "{0}{1}" -f $name,$lastname
                        $username = $username.substring(0,15)
                    }
                }
            }
        }
        Default{
            $username = "{0}{1}" -f $name,$secondlastname
            
            if($username.Length -gt 15){
                $username = $username.substring(0,15)
            }
        }
    } 

    return $username
    
}
#####################################################################################################################
#####################################################################################################################
#Variables

#Script Vars
$ScriptName = ($MyInvocation.MyCommand.Name).trim(".ps1")
$errorcode = 0
$Logfile = "C:\temp\APP_{0}_{1}.log" -f $ScriptName,(Get-Date).tostring("hhmmyyyyddMM")
$MonthName = Capitalize((Get-Date).tostring("MMMM"))
$DayName = Capitalize((Get-Date).tostring("dddd"))
# $MonthNumber = (Get-Date).tostring("MM")
$YearNumber = (Get-Date).tostring("yyy")
$DayNumber = (Get-Date).tostring("dd")

$Testing = $False

if(($Testing)){
    #Testing Variables; INSERT HERE YOUR VARIABLES FOR DEBUGGING
    $UserRunning = "UserName"

    #Forms Location Vars
    $RootFormsFolder = "Sox Files"
    $ResourcesPath = "$($PWD.path)\Resources"
    $UserListFormSavePath = "$($PWD.path)\$RootFormsFolder\Sox Form\$($YearNumber)\$($MonthName)"
    $AccessUserListPath = "$($PWD.path)\$RootFormsFolder\Access Form\TMP\userlist.txt"
    $TerminatedUserListPath = "$($PWD.path)\$RootFormsFolder\Access Form\TMP\terminateduserlist.txt"
    $AccessSettingsPath = "$($PWD.path)\$RootFormsFolder\Access Form\TMP\accesslist.txt"
    $AccessFormSavePath = "$($PWD.path)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)"
    $TerminatedFormSavePath = "$($PWD.path)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)\Terminated"
    $UserListFile = "$($PWD.path)\userlist.txt"

    CreateFolder("$($PWD.path)\$RootFormsFolder\Access Form\TMP")
    CreateFolder("$($PWD.path)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)\Terminated")
    CreateFolder("$($PWD.path)\$RootFormsFolder\Sox Form\$($YearNumber)\$($MonthName)")
}else{
    $UserRunning = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.split("\")[-1]

    #Forms Location Vars
    $RootFormsFolder = "Sox Files"
    $ResourcesPath = "$($PSScriptRoot)\Resources"
    $UserListFormSavePath = "$($PSScriptRoot)\$RootFormsFolder\Sox Form\$($YearNumber)\$($MonthName)"
    $AccessUserListPath = "$($PSScriptRoot)\$RootFormsFolder\Access Form\TMP\userlist.txt"
    $TerminatedUserListPath = "$($PSScriptRoot)\$RootFormsFolder\Access Form\TMP\terminateduserlist.txt"
    $AccessSettingsPath = "$($PSScriptRoot)\$RootFormsFolder\Access Form\TMP\accesslist.txt"
    $AccessFormSavePath = "$($PSScriptRoot)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)"
    $TerminatedFormSavePath = "$($PSScriptRoot)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)\Terminated"
    $UserListFile = "$($PSScriptRoot)\userlist.txt"

    #Folder Creation
    CreateFolder("$($PSScriptRoot)\$RootFormsFolder\Access Form\TMP")
    CreateFolder("$($PSScriptRoot)\$RootFormsFolder\Access Form\$($YearNumber)\$($MonthName)\Terminated")
    CreateFolder("$($PSScriptRoot)\$RootFormsFolder\Sox Form\$($YearNumber)\$($MonthName)")
}

#Resources Vars
$AccessFormFile = "$($ResourcesPath)\NetworkUserForm.docx"
$TerminatedFormFile = "$($ResourcesPath)\TerminatedForm.docx"
$SoxFormFile = "$($ResourcesPath)\UserListForm.docx"
$ItSign = "$($ResourcesPath)\SignatureExample.png"
$UserListEmailTemplate = "$($ResourcesPath)\DefaultSoxEmailTemplate.html"



if($UserRunning -notlike "User*"){
    if(!(Test-Path $UserlistFile)){
        copy-item -path "$($ResourcesPath)\Templateuserlist.txt" -Destination "$($PSScriptRoot)\userlist.txt"
    }else{
        $UserListToAdd = Get-Content($UserlistFile)
    }
}else{
    try{
        $AccessUserListToAdd = Get-Content($AccessUserListPath)
        $TerminatedUserList = Get-Content($TerminatedUserListPath)
        $AccessSettings = Get-Content($AccessSettingsPath)
    }catch{
        WriteLog "WARNING:DATA_FILES_NOT_AVAILABLES: $($PSItem)" #Error Code
        $errorcode = 1
    }
}
$UserListFormName = "$($UserListFormSavePath)\company Sox $($MonthName) - $($YearNumber) - $($DayNumber).docx"
$TerminatedFormName = "$($UserListFormSavePath)\company Sox $($MonthName) - $($YearNumber) - $($DayNumber).docx"

#Forms Vars
$SoxFormTableHeader = @("First Name","Last Name","Personal e-mail","Site Code","Time Attendance App User Number","PayRoll App User Number","Position","Department","Starting / Leaving (Alta / Baja)","Date")
$AccessFormTableHeader = @("Team Member Name:","Sequence Number:","Effective Date of Form:","Personal Email:")
$AccessFormTableHeader2 = @("Employee Num.","Position / Dept:","Site Code:") 
$TerminatedFormTableHeader = @("LAST NAME:","PAYROLL APP USER NUM:") 
$TerminatedFormTableHeader2 = @("FIRST NAME:","END DATE:") 
$AccessBodyTable = @("Network Access:","Intranet Access:","Program1 Access:","Program2 Access:","Program3 Access:","Program4 Access:","Program5 Access:","Company Email Access:","Terminal Server Access:","Payroll App Access:","Time Attendance Access:","Program6 Access:","Program7 Access:","Program8 Access:","Program9 Access:","Program10 Access:")
$TerminatedBodyTable = @("DEPARMENT1","EXAMPLE1","HUMAN RESOURCES","EXAMPLE2","PAYROLL SYSTEM","TIME ATTENDANCE SYSTEM","FINANCE","EXAMPLE3","DEPARTMENT2","EXAMPLE4","EXAMPLE5","INFORMATION TECHNOLOGY","OFFICE KEY","EXAMPLE6","EXAMPLE7","ACTIVE DIRECTORY","DEPARTMENT3","EXAMPLE8")
$TerminatedFooterTable = @("DEPARMENT1 SUPERVISOR / MANAGER SIGNATURE","HUMAN RESOURCES SUPERVISOR / MANAGER SIGNATURE","FINANCE SUPERVISOR / MANAGER SIGNATURE","DEPARTMENT2 SUPERVISOR / MANAGER SIGNATURE","IT SUPERVISOR / MANAGER SIGNATURE","DEPARTMENT3 SUPERVISOR / MANAGER SIGNATURE")
$AccessSettingsKeys = @("Network","Intranet", "Program1A","Program2", "Program3", "Program4", "Program5", "Email", "TerminalServer", "Payroll", "TimeAttendance", "Program6", "Program7", "Program8", "Program9", "Program10")
$Disclaimer = "`nThe Username allows the undersigned to access to system functions designed for his/her job profile. Passwords and security codes must not be disclosed to any employees or other individuals since it would enable them to have access to sensitive hotel data. By signing this form, I acknowledge the company requirement not to disclose data or procedures for accessing data to which I have no access or knowledge. Any breach of security or improper use of the systems will be ground for dismissal.`nI have read and understood the above and acknowledge the responsibility for my Username, password or any other security code. I will use them only for proper discharge of my responsibilities and for no other purpose.`nIf I forget my Username, password or security code, I will contact the Information Systems Manager (ISM) immediately. I will also contact the ISM if anyone learns of my password or security code."
$TeamsAlert = "`nAll Team Members who have been authorized to generate / make guest keys must do so while maintaining the integrity of the guest and the Company Corporation. Below is a list of guidelines to be followed:"
$SecuritySteps = "`t * Always verify the identity of the subject for who you are making a key. Examine the subjects ID and confirm room authorization to the room. When in doubt, call Security.`n`t * Never make / generate a key for a non-guest. When in doubt, call Security.`n`t * Never make more than the requested number of keys. Once a key/keys have been made, keep it/them secured.`n`t * Never divulge your access code to anyone. If you believe that your code has been compromised, contact the Information Systems Manager."
$GreenColor = 32768
$RedColor = 255
$CyanColor = -738131969 

#Mail Vars
$FromEmailList = "company_HRD@Company.com"
$ToEmailList = "company_ISM@Company.com"
$CCEmailList = "company_HRD@Company.com"
$EmailSubject = "Sox Form $($DayName) $($DayNumber) $($MonthName)"

$EmailBody = Get-Content("$($UserListEmailTemplate)")

#####################################################################################################################
#####################################################################################################################

#Script
try{
    if($UserRunning -notlike "User*"){
        #Opening User List Sox Form
        $SoxFormVariables = OpenWordFile($SoxFormFile)
        $Selection = $SoxFormVariables[-1]
        $DocW = $SoxFormVariables[-2]
        $Word = $SoxFormVariables[-3]
        $StateNow = $SoxFormVariables[-4]
        $StateBefore = $SoxFormVariables[-5]

        #Creating the table in the Word
        try{
            $Table = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                $Selection.Range,($UserListToAdd.Count+1),10, #Create a table with next size (Amount of users to add in the list + amount of columns from the form)
                [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, #Setting to create a table in word
                [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
            )

            ## Writing the Table's header
            for($x = 1;$x -le 10;$x++){
                $Table.Cell(1,$x).Range.text = $SoxFormTableHeader[$x-1] #writing the first line;0 is row & $x is column
                $Table.Cell(1,$x).VerticalAlignment = 1 #Vertical aligment
                $Table.Cell(1,$x).Range.Bold = 1 #Set string to Bold
                $Table.Cell(1,$x).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
            }
            
            for($y = 2;$y -le ($UserListToAdd.count + 1);$y++){ #Filling table with user data
                for($x = 1;$x -le 10;$x++){

                    if($UserListtoadd.count -ne 1){
                        $Table.Cell($y,$x).Range.text = $UserListToAdd[$y-2].split(";")[$x-1] #Cell with data
                        $Table.Cell($y,$x).Range.ParagraphFormat.Alignment = 1 #Aligment to center
                        $Table.Cell($y,$x).VerticalAlignment = 1 #Vertical aligment
                        if($UserListToAdd[$y-2].split(";")[$x-1] -eq "Starting"){
                            $Table.Cell($y,$x).range.shading.BackgroundPatternColor = $GreenColor
                            $Table.Cell($y,$x).Range.Bold = 1
                        }elseif ($UserListToAdd[$y-2].split(";")[$x-1] -eq "Leaving") {
                            $Table.Cell($y,$x).range.shading.BackgroundPatternColor = $RedColor
                            $Table.Cell($y,$x).Range.Bold = 1
                        }
                    }else{
                        $Table.Cell($y,$x).Range.text = $UserListToAdd.split(";")[$x-1] #Cell with data
                        $Table.Cell($y,$x).Range.ParagraphFormat.Alignment = 1 #Aligment to center
                        $Table.Cell($y,$x).VerticalAlignment = 1 #Vertical aligment
                        if($UserListToAdd.split(";")[$x-1] -eq "Starting"){
                            $Table.Cell($y,$x).range.shading.BackgroundPatternColor = $GreenColor
                            $Table.Cell($y,$x).Range.Bold = 1
                        }elseif ($UserListToAdd.split(";")[$x-1] -eq "Leaving") {
                            $Table.Cell($y,$x).range.shading.BackgroundPatternColor = $RedColor
                            $Table.Cell($y,$x).Range.Bold = 1
                        }
                    }
                }
            }

            $DocW.SaveAs($UserListFormName) #Saving the document to the dir given
            SaveAsPDF($Word,$DocW,$UserListFormName) #Saving the document to the dir given as a PDF
            $DocW.close() #Closing file
            $Word.quit() #Closing Word Application

        }catch{
            WriteLog "ERROR;USER_LIST_TABLE_CREATION: $($PSItem)" #Error code
            $errorcode = 1
            KillProcess -StateBefore $StateBefore -StateNow $StateNow
            
        }

        #Sending Email
        try{
            SendEmail -from $FromEmailList -to $ToEmailList -cc $CCEmailList -Subject $EmailSubject -Attach $UserListFormName -Body $EmailBody
        }catch{
            WriteLog "ERROR:SENDING_EMAIL: $($PSItem)"
            $errorcode = 1
        }

        #Transfering Data from UserList file to AccessList file
        try{
            foreach($line in $UserListToAdd){
                if($line.split(";")[-2] -eq "Starting"){
                    if(!(Test-path $AccessUserListPath)){
                        $line >> $AccessUserListPath
                    }else{
                        $AccessList = Get-Content($AccessUserListPath)
                        foreach($user in $AccessList){
                            $userData = "{0};{1}" -f $user.split(";")[0],$user.split(";")[1]
                            $linedata = "{0};{1}" -f $line.split(";")[0],$line.split(";")[1]
                            
                            if(!($linedata -eq $userData)){ #if line to be add is not same than the one we got from before, then add
                                $line >> $AccessUserListPath
                            }elseif(!($line -eq $user)){ #check if the whole line is same, if not, then add
                                $line >> $AccessUserListPath
                            }
                        }
                    }
                }elseif ($line.split(";")[-2] -eq "Leaving"){
                    if(!(Test-path $TerminatedUserListPath)){
                        $line >> $TerminatedUserListPath
                    }else{
                        $TerminatedList = Get-Content($TerminatedUserListPath)
                        foreach($user in $TerminatedList){
                            $userData = "{0};{1}" -f $user.split(";")[0],$user.split(";")[1]
                            $linedata = "{0};{1}" -f $line.split(";")[0],$line.split(";")[1]
                            
                            if(!($linedata -eq $userData)){ #if line to be add is not same than the one we got from before, then add
                                $line >> $TerminatedUserListPath
                            }elseif(!($line -eq $user)){ #check if the whole line is same, if not, then add
                                $line >> $TermintatedUserListPath
                            }
                        }
                    }
                }else{
                    WriteLog "Warning;USER_LIST_FORMAT_ERROR: User ($($line.split(";")[0]) $($line.split(";")[1])) Has Incorrect Data"
                }
            }
        }catch{
            WriteLog "ERROR;USER_LIST_FILE_TRANSFERING: $($PSItem)" #Error code

        }

    }else{
        #Creating Access Form using User List
        $AccessFormsDone = 1
        if(Test-Path $AccessUserListPath){
            foreach($user in $AccessUserListToAdd){
                #Opening Access Sox Form
                $AccessSoxFormVariables = OpenWordFile($AccessFormFile)
                $Selection = $AccessSoxFormVariables[-1]
                $DocW = $AccessSoxFormVariables[-2]
                $Word = $AccessSoxFormVariables[-3]
                $StateNow = $AccessSoxFormVariables[-4]
                $StateBefore = $AccessSoxFormVariables[-5]

                #Form Variables
                if($AccessSettings.count -eq 1){
                    $UserLine = $AccessSettings
                }else{
                    $UserLine = $AccessSettings[$AccessFormsDone - 1]
                }
                $UserSettings = ConvertToHash($UserLine)
                $Name = $User.split(";")[0]
                $Lastname = $User.split(";")[1]
                $OtherInncode = $User.split(";")[3]
                $EmployeeNum = $User.split(";")[4]
                $Position = $User.split(";")[6]
                $Dept = $User.split(";")[7]
                $StartDate = $User.split(";")[9]
                $username = CreateUserName -name "$($Name)" -lastname "$($Lastname)"
                
                Switch($Dept){
                    {$_ -like "EX1*" -or $_ -like "EX01*"}{
                        $KeyType = "KEY01"
                        $HoDName = "DEPT1 MANAGER"
                    }
                    {$_ -like "EX2*" -or $_ -like "EX02*"}{
                        $KeyType = "KEY02"
                        $HoDName = "DEPT2 MANAGER"
                    }
                    {$_ -like "EX3*" -or $_ -like "EX03*"}{
                        $KeyType = "KEY03"
                        $HoDName = "DEPT3 MANAGER"
                    }
                    {$_ -like "EX4*" -or $_ -like "FB*" -or $_ -like "EX04*"}{
                        $KeyType = "KEY04"
                        $HoDName = "DEPT4 MANAGER"
                    }
                    {$_ -like "EX5*" -or $_ -like "EX05*"}{
                        $KeyType = "KEY05"
                        $HoDName = "DEPT5 MANAGER"
                    }
                    Default{
                        $KeyType = "Not Assigned Yet"
                        $HoDName = "-"
                    }
                }

                #Creating the table in the Word
                try{

                    ####
                    # FIRST TABLE
                    ####

                    $Table = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,4,5, #Create a table with next size (Amount of users to add in the list + amount of columns from the form)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    ## Writing the Table's header
                    for($x = 1;$x -le 4;$x++){
                        $Table.Cell($x,1).Range.text = $AccessFormTableHeader[$x-1] #writing the first line;
                        $Table.Cell($x,1).VerticalAlignment = 1 #Vertical aligment
                        $Table.Cell($x,1).Range.Bold = 1 #Set string to Bold
                        $Table.Cell($x,1).Range.Font.Size = 9
                        
                        
                        $Table.Cell($x,3).Range.text = $AccessFormTableHeader2[$x-1] #writing the first line;n
                        $Table.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                        $Table.Cell($x,3).Range.Bold = 1 #Set string to Bold
                        $Table.Cell($x,3).Range.Font.Size = 9

                        if($x -eq 4){
                            $Table.Cell($x,3).Merge($Table.Cell($x,2))
                            $Table.Cell($x,3).Range.text = "Office Key:"
                            $Table.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                            $Table.Cell($x,3).Range.Bold = 1 #Set string to Bold
                            $Table.Cell($x,3).Range.Font.Size = 9
                        }
                    }
                    
                    #Filling Header Table with User Data
                    #TEAM MEMBER NAME
                    $Table.Cell(1,2).Range.text = "{0} {1}" -f $User.split(";")[0],$User.split(";")[1] #writing the first line;
                    $Table.Cell(1,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(1,2).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(1,2).Range.Font.Size = 9
                    
                    #SEQUENCE NUMBER
                    $Table.Cell(2,2).Range.text = "{0} - {1}" -f $MonthName,$AccessFormsDone
                    $Table.Cell(2,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(2,2).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(2,2).Range.Font.Size = 9
                    
                    #EFFECTIVE DATE FORM
                    $Table.Cell(3,2).Range.text = $User.split(";")[-1]
                    $Table.Cell(3,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(3,2).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(3,2).Range.Font.Size = 9
                    
                    #PERSONAL EMAIL
                    $Table.Cell(4,2).Range.text = $User.split(";")[2]
                    $Table.Cell(4,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(4,2).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(4,2).Range.Bold = 0 #Set string to Bold
                    $Table.Cell(4,2).Range.Font.Size = 9
                    
                    #EMPLOYEE NUM
                    $Table.Cell(1,4).Range.text = $User.split(";")[4]
                    $Table.Cell(1,4).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(1,4).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(1,4).Range.Font.Size = 9
                    
                    $Table.Cell(1,5).Range.text = $User.split(";")[5]
                    $Table.Cell(1,5).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(1,5).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(1,5).Range.Font.Size = 9
                    
                    #POSITION / DEPT
                    $Table.Cell(2,5).Merge($Table.Cell(2,4))
                    $Table.Cell(2,4).Range.text = "{0} / {1}" -f $Position,$Dept
                    $Table.Cell(2,4).VerticalAlignment = 0 #Vertical aligment
                    $Table.Cell(2,4).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(2,4).Range.Font.Size = 9
                    
                    #SITE CODE:
                    $Table.Cell(3,5).Merge($Table.Cell(3,4))
                    if($OtherInncode.length -eq 0){
                        $Table.Cell(3,4).Range.text = "-"
                    }else{
                        $Table.Cell(3,4).Range.text = "$($OtherInncode)"
                    }
                    $Table.Cell(3,4).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(3,4).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(3,4).Range.Font.Size = 9
                    
                    #OFFICE KEY
                    $Table.Cell(4,4).Range.text = $KeyType
                    $Table.Cell(4,4).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(4,4).Range.ParagraphFormat.Alignment = 1 #Center String in Cell
                    $Table.Cell(4,4).Range.Font.Size = 9

                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength


                    ####
                    # SECOND TABLE
                    ####

                    #Creating Body Table from Document
                    $TableB = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,($AccessBodyTable.count),5, #Create a table with next size (Amount of users to add in the list + amount of columns from the form)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    ## Writing the Body Table
                    for($x = 1;$x -le ($AccessBodyTable.count);$x++){
                        $TableB.Cell($x,1).Range.text = $AccessBodyTable[$x-1] #writing the first line;
                        $TableB.Cell($x,1).VerticalAlignment = 1 #Vertical aligment
                        $TableB.Cell($x,1).Range.Bold = 1 #Set string to Bold
                        $TableB.Cell($x,1).Range.Font.Size = 9
                        
                        if($UserSettings.$($AccessSettingsKeys[$x-1]) -eq "Yes"){
                            $TableB.Cell($x,2).Range.text = "T"  #writing the first line;
                            $TableB.Cell($x,2).Range.font.name = "Wingdings 2"  #writing the first line;
                        }
                        $TableB.Cell($x,2).VerticalAlignment = 1 #Vertical aligment
                        $TableB.Cell($x,2).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                        $TableB.Cell($x,2).Range.Font.Size = 15
                        # $TableB.Cell($x,2).Range.Bold = 1 #Set string to Bold

                        if($AccessBodyTable[$x-1] -like "Program1 A*"){
                            $TableB.Cell($x,3).Range.text = "Username / Rights:" #writing the first line;
                            $TableB.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                            $TableB.Cell($x,3).Range.Bold = 1 #Set string to Bold
                            $TableB.Cell($x,3).Range.Font.Size = 9
                            
                        }elseif($AccessBodyTable[$x-1] -match "Program6*"){
                            $TableB.Cell($x,3).Range.text = "Username / Card:" #writing the first line;
                            $TableB.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                            $TableB.Cell($x,3).Range.Bold = 1 #Set string to Bold
                            $TableB.Cell($x,3).Range.Font.Size = 9

                        }else{
                            $TableB.Cell($x,3).Range.text = "Username:" #writing the first line;
                            $TableB.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                            $TableB.Cell($x,3).Range.Bold = 1 #Set string to Bold
                            $TableB.Cell($x,3).Range.Font.Size = 9
                            $TableB.Cell($x,5).Merge($TableB.cell($x,4))
                        }

                        if($UserSettings.$($AccessSettingsKeys[$x-1]) -eq "Yes"){
                            switch ($AccessSettingsKeys[$x-1]){
                                {$_ -eq "Network"}{
                                    $TableB.cell($x,4).range.text = "DOMAIN\$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Intranet"}{
                                    $TableB.cell($x,4).range.text = $username
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program1A"}{
                                    $TableB.cell($x,4).range.text = $username.ToUpper()
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                    switch ($Dept) {
                                        {$_ -like "DEPT1*"}{
                                            $TableB.cell($x,5).range.text = "DEPARTMENT1 Staff"
                                            $TableB.Cell($x,5).VerticalAlignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.Font.Size = 9
                                        }
                                        {$_ -like "DEPT2*"}{
                                            $TableB.cell($x,5).range.text = "DEPARTMENT2 Staff"
                                            $TableB.Cell($x,5).VerticalAlignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.Font.Size = 9
                                        }
                                        {$_ -like "DEPT3*"}{
                                            $TableB.cell($x,5).range.text = "DEPARTMENT3 Staff"
                                            $TableB.Cell($x,5).VerticalAlignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                            $TableB.Cell($x,5).Range.Font.Size = 9
                                        }
                                    }
                                }
                                {$_ -eq "Program2"}{
                                    $TableB.cell($x,4).range.text = $username.ToUpper()
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program3"}{
                                    $TableB.cell($x,4).range.text = "DOMAIN\$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program4"}{
                                    $TableB.cell($x,4).range.text = "DOMAIN\$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program5"}{
                                    $TableB.cell($x,4).range.text = "{0}{1}" -f $username.toupper()[0],$username.substring($name.length,$username.length - $name.length).toupper() #getting first char from username, deleting name - first char and giving last name in upper. Doing it like this, because we need to user the correct username as sometimes user could have username different than firstname + lastname. ex: my account was HectorFuste instead of HectorRodriguez
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Email"}{
                                    $TableB.cell($x,4).range.text = "{0}.{1}@Company.com" -f $name,$username.substring($name.length,$username.length - $name.length)
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "TerminalServer"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Payroll"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "TimeAttendance"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program6"}{
                                    $TableB.cell($x,4).range.text = "$($Position)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9

                                    $TableB.cell($x,5).range.text = "Not Assigned yet"
                                    $TableB.Cell($x,5).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,5).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,5).Range.Font.Size = 9
                                }
                                {$_ -eq "Program7"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program8"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program9"}{
                                    $TableB.cell($x,4).range.text = "$($username)"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }
                                {$_ -eq "Program10"}{
                                    $TableB.cell($x,4).range.text = "No Assigned Yet"
                                    $TableB.Cell($x,4).VerticalAlignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1 #Vertical aligment
                                    $TableB.Cell($x,4).Range.Font.Size = 9
                                }

                            }
                        }
                    }

                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength

                    ####
                    # THIRD TABLE
                    ####

                    $TableC = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,(7),5, #Create a table with next size (Amount of users to add in the list + amount of columns from the form)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    $TableC.Cell(4,5).Merge($TableC.Cell(4,4))
                    $TableC.Cell(4,4).Merge($TableC.Cell(4,3))
                    $TableC.Cell(4,3).Merge($TableC.Cell(4,2))
                    $TableC.Cell(4,2).Merge($TableC.Cell(4,1))
                    
                    $TableC.Cell(5,5).Merge($TableC.Cell(5,4))
                    $TableC.Cell(5,4).Merge($TableC.Cell(5,3))
                    $TableC.Cell(5,3).Merge($TableC.Cell(5,2))
                    $TableC.Cell(5,2).Merge($TableC.Cell(5,1))
                    
                    $TableC.Cell(6,5).Merge($TableC.Cell(6,4))
                    $TableC.Cell(6,4).Merge($TableC.Cell(6,3))
                    $TableC.Cell(6,3).Merge($TableC.Cell(6,2))
                    $TableC.Cell(6,2).Merge($TableC.Cell(6,1))
                    
                    $TableC.Cell(7,5).Merge($TableC.Cell(7,4))
                    $TableC.Cell(7,4).Merge($TableC.Cell(7,3))
                    $TableC.Cell(7,3).Merge($TableC.Cell(7,2))
                    $TableC.Cell(7,2).Merge($TableC.Cell(7,1))

                    for($x = 1;$x -le 5;$x++){
                        switch ($x) {
                            {$_ -eq 1} {
                                $TableC.Cell(1,$x).Range.Text = "Department Head:"
                                $TableC.Cell(1,$x).Range.Font.Size = 8
                                $TableC.Cell(1,$x).Range.Italic = 1
                                
                                $TableC.Cell(2,$x).Range.Text = "$($HoDName)"
                                $TableC.Cell(2,$x).Range.Font.Size = 8
                                $TableC.Cell(2,$x).Range.Bold = 1

                                if($HoDName -eq "MAN01"){
                                    $ImgSelection = $TableC.Cell(3,$x).range.start
                                    $TableC.Cell(3,$x).Range.ParagraphFormat.Alignment = 0
                                    $Selection.start = $ImgSelection
                                    # $Selection.end = 1226
                                    
                                    
                                    $img = $Selection.InlineShapes.AddPicture("$($ItSign)")
                                    $img.Height = 40
                                    $img.Width = 75
                                }
                                
                                $TableC.Cell(4,$x).Range.Text = "$($Disclaimer)"
                                $TableC.Cell(4,$x).Range.Font.Size = 7
                                $TableC.Cell(4,$x).Range.Italic = 1
                                
                                $TableC.Cell(5,$x).Range.Text = "Key system access authorization"
                                $TableC.Cell(5,$x).Range.Font.Size = 9
                                $TableC.Cell(5,$x).Range.Bold = 1
                                
                                $TableC.Cell(6,$x).Range.Text = "$($TeamsAlert)"
                                $TableC.Cell(6,$x).Range.Font.Size = 7
                                $TableC.Cell(6,$x).Range.Italic = 1
                                
                                $TableC.Cell(7,$x).Range.Text = "$($SecuritySteps)"
                                $TableC.Cell(7,$x).Range.Font.Size = 7
                                $TableC.Cell(7,$x).Range.Italic = 1
                                
                            }
                            {$_ -eq 2}{
                                if($HoDName -eq "MAN01"){
                                    $TableC.Cell(3,$x).Range.text = $StartDate
                                    $TableC.Cell(3,$x).Range.Font.Size = 8
                                    $TableC.Cell(3,$x).Range.Italic = 1
                                    $TableC.Cell(3,$x).Range.ParagraphFormat.Alignment = 0
                                }
                            }
                            {$_ -eq 3}{
                                $TableC.Cell(1,$x).Range.Text = "Cross-Training Department Head:"
                                $TableC.Cell(1,$x).Range.Font.Size = 8
                                $TableC.Cell(1,$x).Range.Italic = 1
                                $TableC.Cell(1,$x).Range.ParagraphFormat.Alignment = 1
                                
                            }
                            {$_ -eq 4}{
                                $TableC.Cell(3,$x).Range.text = $StartDate
                                $TableC.Cell(3,$x).Range.Font.Size = 8
                                $TableC.Cell(3,$x).Range.Italic = 1
                                $TableC.Cell(3,$x).Range.ParagraphFormat.Alignment = 2
                            }
                            {$_ -eq 5}{
                                $TableC.Cell(1,$x).Range.Text = "Information System Manager:"
                                $TableC.Cell(1,$x).Range.Font.Size = 8
                                $TableC.Cell(1,$x).Range.Italic = 1
                                $TableC.Cell(1,$x).Range.ParagraphFormat.Alignment = 2
                                
                                $TableC.Cell(2,$x).Range.Text = "MANAGER 01"
                                $TableC.Cell(2,$x).Range.Font.Size = 8
                                $TableC.Cell(2,$x).Range.Bold = 1
                                $TableC.Cell(2,$x).Range.ParagraphFormat.Alignment = 2
                                
                                # $TableC.Cell(3,$x).Range.Text = "INSERT SIGNATURE HERE"
                                $ImgSelection = $TableC.Cell(3,$x).range.start
                                $TableC.Cell(3,$x).Range.ParagraphFormat.Alignment = 2

                                $Selection.start = $ImgSelection
                                # $Selection.end = 1226
                                
                                
                                $img = $Selection.InlineShapes.AddPicture("$($ItSign)")
                                $img.Height = 40
                                $img.Width = 75

                            }
                            Default {
                                $TableC.Cell(1,$x).Range.Text = ""
                            }
                        }
                    }
                    

                    for($x = 1;$x -le 7;$x++){
                        if($x -ne 1){
                            $TableC.Borders[$x].Visible = $False
                        }else{
                            $TableC.Borders[$x].LineWidth = 12
                            $TableC.Borders[$x].ColorIndex = 3
                            $TableC.Borders[$x].Color = $CyanColor
                        }
                    } 
                    
                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength

                    ####
                    # FOURTH TABLE
                    ####

                    $TableD = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,(1),4, #Create a table with next size (Amount of users to add in the list + amount of columns from the form)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    for($x = 1;$x -le 7;$x++){
                        if($x -ne 1){
                            $TableD.Borders[$x].Visible = $False
                        }else{
                            $TableD.Borders[$x].LineWidth = 12
                            $TableD.Borders[$x].ColorIndex = 3
                            $TableD.Borders[$x].Color = $CyanColor
                        }
                    }  
                    
                    $TableD.Cell(1,1).range.text = "Date:"
                    $TableD.Cell(1,1).Range.Font.Size = 8
                    $TableD.Cell(1,1).Range.Italic = 1
                    
                    $TableD.Cell(1,2).range.text = $User.split(";")[-1]
                    $TableD.Cell(1,2).Range.Font.Size = 8
                    
                    # $TableD.Cell(1,3).range.text = ""
                    
                    $TableD.Cell(1,3).range.text = "Employee:"
                    $TableD.Cell(1,3).Range.Font.Size = 8
                    $TableD.Cell(1,3).Range.Italic = 1
                    $TableD.Cell(1,3).Range.ParagraphFormat.Alignment = 0
                    
                    $TableD.Cell(1,4).range.text = "{0} {1}" -f $User.split(";")[0],$User.split(";")[1]
                    $TableD.Cell(1,4).Range.Font.Size = 8
                    $TableD.Cell(1,4).Range.ParagraphFormat.Alignment = 2

                    
                    $AccessFormName = "$($AccessFormSavePath)\company_$($User.split(";")[0])_$($User.split(";")[1])-$($User.split(";")[-1].replace("/","-")).docx"

                    $DocW.SaveAs($AccessFormName) #Saving the document to the dir given
                    SaveAsPDF($Word,$DocW,$AccessFormName)
                    $DocW.close() #Closing file
                    $Word.quit() #Closing Word Application
                }catch{
                    WriteLog "ERROR;ACCESS_FORM_WORD_TABLE_CREATION: $($PSItem)" #Error code
                    $errorcode = 1
                    KillProcess -StateBefore $StateBefore -StateNow $StateNow
                }
                $AccessFormsDone += 1
            }
        }else{
            WriteLog "ERROR;ACCESS_USER_LIST: $($PSitem)"
            $errorcode = 1
        }

        if(Test-Path $TerminatedUserListPath){
            foreach($user in $TerminatedUserList){
                $TerminatedSoxFormVariables = OpenWordFile($TerminatedFormFile)
                $Selection = $TerminatedSoxFormVariables[-1]
                $DocW = $TerminatedSoxFormVariables[-2]
                $Word = $TerminatedSoxFormVariables[-3]
                $StateNow = $TerminatedSoxFormVariables[-4]
                $StateBefore = $TerminatedSoxFormVariables[-5]

                #Form Variables
                $Name = $User.split(";")[0]
                $Lastname = $User.split(";")[1]
                $OtherInncode = $User.split(";")[3]
                $EmployeeNum = $User.split(";")[4]
                $Position = $User.split(";")[6]
                $Dept = $User.split(";")[7]
                $EndDate = $User.split(";")[9]

                try {
                    ####
                    # FIRST TABLE
                    ####

                    $Table = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,$TerminatedFormTableHeader.length,4, #Create a table with next size (row , columns)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    ## Writing the Table's header
                    for($x = 1;$x -le $TerminatedFormTableHeader.length;$x++){
                        $Table.Cell($x,1).Range.text = $TerminatedFormTableHeader[$x-1] #writing the first line;
                        $Table.Cell($x,1).VerticalAlignment = 1 #Vertical aligment
                        $Table.Cell($x,1).Range.ParagraphFormat.Alignment = 2 #Vertical aligment
                        $Table.Cell($x,1).Range.Bold = 1 #Set string to Bold
                        $Table.Cell($x,1).Range.Font.Size = 14
                        
                        
                        
                        $Table.Cell($x,3).Range.text = $TerminatedFormTableHeader2[$x-1] #writing the first line;n
                        $Table.Cell($x,3).VerticalAlignment = 1 #Vertical aligment
                        $Table.Cell($x,3).Range.ParagraphFormat.Alignment = 2 #Vertical aligment
                        $Table.Cell($x,3).Range.Bold = 1 #Set string to Bold
                        $Table.Cell($x,3).Range.Font.Size = 14
                        
                    }
                    
                    #Adding User Data
                    $Table.Cell(1,2).Range.text = $Lastname  #writing the first line;
                    $Table.Cell(1,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(1,2).Range.Font.Size = 11

                    $Table.Cell(2,2).Range.text = $EmployeeNum  #writing the first line;
                    $Table.Cell(2,2).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(2,2).Range.Font.Size = 11

                    $Table.Cell(1,4).Range.text = $name  #writing the first line;
                    $Table.Cell(1,4).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(1,4).Range.Font.Size = 11

                    $Table.Cell(2,4).Range.text = $EndDate  #writing the first line;
                    $Table.Cell(2,4).VerticalAlignment = 1 #Vertical aligment
                    $Table.Cell(2,4).Range.Font.Size = 11

                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength

                    ####
                    # SECOND TABLE
                    ####

                    $TableB = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,$TerminatedBodyTable.Length,5, #Create a table with next size (row , columns)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    #Writing Table Headers
                    for($x = 1;$x -le $TerminatedBodyTable.length;$x++){
                        switch($x){
                            {$_ -eq 1}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            {$_ -eq 3}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            {$_ -eq 7}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            {$_ -eq 9}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            {$_ -eq 12}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            {$_ -eq 17}{
                                $TableB.Cell($x,1).Merge($TableB.Cell($x,2))
                                $TableB.Cell($x,1).Range.Font.Size = 14
                                $TableB.Cell($x,1).Range.Bold = 1
                                $TableB.Cell($x,1).Range.Underline = 1
                            }
                            Default{
                                $TableB.Cell($x,1).Range.Font.Size = 13
                                $TableB.Cell($x,1).Range.Italic = 1
                                
                                $TableB.Cell($x,3).range.text = "YES"
                                $TableB.Cell($x,3).Range.ParagraphFormat.Alignment = 1
                                $TableB.Cell($x,3).Range.Font.Size = 13
                                
                                $TableB.Cell($x,4).range.text = "NO"
                                $TableB.Cell($x,4).Range.ParagraphFormat.Alignment = 1
                                $TableB.Cell($x,4).Range.Font.Size = 13
                                
                                $TableB.Cell($x,5).range.text = "N/A"
                                $TableB.Cell($x,5).Range.ParagraphFormat.Alignment = 1
                                $TableB.Cell($x,5).Range.Font.Size = 13
                                
                            }
                        }
                        
                        #Adding Data
                        $TableB.Cell($x,1).range.text = $TerminatedBodyTable[$x-1]
                        $TableB.Cell($x,1).VerticalAlignment = 1 #Vertical aligment
                        
                        foreach($n in (2,4,8,10,13,14,15,18)){
                            $TableB.Cell($n,2).range.text = "RETURNED"
                            $TableB.Cell($n,2).Range.ParagraphFormat.Alignment = 2
                            $TableB.Cell($n,2).Range.Font.Size = 13
                        }
                        
                        foreach($n in (5,6,11,16)){
                            $TableB.Cell($n,2).range.text = "TERMINATED"
                            $TableB.Cell($n,2).Range.ParagraphFormat.Alignment = 2
                            $TableB.Cell($n,2).Range.Font.Size = 13
                        }

                    }
                    
                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength

                    ####
                    # THIRD TABLE
                    ####

                    $TableC = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,2,$TerminatedFooterTable.Length, #Create a table with next size (row , columns)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    for($x = 1;$x -le $TerminatedFooterTable.length;$x++){
                        $TableC.Cell(1,$x).range.text = $TerminatedFooterTable[$x-1]
                        $TableC.Cell(1,$x).range.Bold = 1
                        $TableC.Cell(1,$x).VerticalAlignment = 0 
                        $TableC.Cell(1,$x).Range.ParagraphFormat.Alignment = 1
                        $TableC.Cell(1,$x).range.Font.Size = 12

                        $TableC.Cell(2,$x).range.text = "`n`n`n`n`n`n{0}" -f $EndDate
                        $TableC.Cell(2,$x).range.Italic = 1
                        $TableC.Cell(2,$x).VerticalAlignment = 1 
                        $TableC.Cell(2,$x).Range.ParagraphFormat.Alignment = 1
                        $TableC.Cell(2,$x).range.Font.Size = 11
                    }

                    $Selection.start = $Selection.StoryLength
                    $Selection.Paragraphs.Add() | Out-Null
                    $Selection.start = $Selection.StoryLength

                    ####
                    # FOURTH TABLE
                    ####

                    $TableD = $Selection.Tables.add( #Creating the table using the Selector pointer we made before
                        $Selection.Range,3,5, #Create a table with next size (row , columns)
                        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior#, #Setting to create a table in word
                        # [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent #setting cell size fit with content
                    )

                    $TableD.Cell(1,1).range.text = "TEAM MEMBER SIGNATURE"
                    $TableD.Cell(1,1).range.Bold = 1
                    $TableD.Cell(1,1).range.Font.Size = 12

                    $TableD.Cell(2,1).range.text = "{0} {1}" -f $name,$Lastname
                    $TableD.Cell(2,1).range.Italic = 1
                    $TableD.Cell(2,1).range.Font.Size = 11

                    $TableD.Cell(3,1).range.text = "`n`n`n`n`n`n{0}" -f $EndDate
                    $TableD.Cell(3,1).range.italic = 1
                    $TableD.Cell(3,1).range.Font.Size = 11


                    for($x = 1;$x -le 7;$x++){
                        $Table.Borders[$x].Visible = $False
                        $TableB.Borders[$x].Visible = $False
                        $TableC.Borders[$x].Visible = $False
                        $TableD.Borders[$x].Visible = $False
                    }
                    $TerminatedFormName  = "$($TerminatedFormSavePath)\company_Terminated_User_$($User.split(";")[0])_$($User.split(";")[1])-$($User.split(";")[-1].replace("/","-")).docx"
                    
                    $DocW.SaveAs($TerminatedFormName) #Saving the document to the dir given
                    SaveAsPDF($Word,$DocW,$TerminatedFormName)
                    $DocW.close() #Closing file
                    $Word.quit() #Closing Word Application
                }catch {
                    WriteLog "ERROR;TERMINATED_FORM_WORD_TABLE_CREATION: $($PSItem)" #Error code
                    KillProcess -StateBefore $StateBefore -StateNow $StateNow
                    $errorcode = 1
                }
            }
        }else{
            # WriteLog "ERROR;TERMINATED_USER_LIST: TERMINATED USER LIST DOES NOT EXIST;NO REPORTS WILL BE MADE"
            $errorcode = 1
        }
    }
}catch{
    if($errorcode -eq 1){
        WriteLog "ERROR;ERROR_IN_SCRIPT_EXECUTION: $($PSItem)" #Error code
        ReportLog
    }
    
}finally{
    write-output "The End"
}