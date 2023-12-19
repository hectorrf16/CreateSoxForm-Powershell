# Create Sox Form
##### _Automatically creates Sox forms for new Team Members_

###### Create Sox Form is the app to create, distribute the Sox Form for new incorporation of Team Members.
###### _A new version of the former `Create Sox Form` 2017 release_

## Features

- Fewer iterations to be made with
- Faster introduction of information
- Easier to debug and maintain
- More customizable as it only uses programs which are already on your machine, nothing to add

#### Features under development
- [X] PDF export after creating the form
- [ ] Visual Interface for specific parts if required
- [ ] '.exe' with icon for better feel & look
- [...]

----

## Pre-Requesites to run
1. Folder structure must be like next Folder Tree diagram
```
CreateSoxForm
├───Resources
├───Screenshots
└───Sox Files
    ├───Access Form
    │   ├───TMP
    │   └───Year
    │       └───Month
    │           └───Terminated
    └───Sox Form
        └───Year
            └───Month
```
2. Root Folder '_CreateSoxForm_' should have next list of files
```
- CreateSoxForm.cmd
- CreateSoxForm.ps1
- userlist.txt
```
>**Note:** `userlist.txt` may not be in the directory, the APP will recreate it if it is not detected.

3. '_Access Form\TMP_'  folder under '_Sox Files_' should have next files:
```
- userlist.txt
- accesslist.txt
- terminateduserlist.txt
```
> **Note:** these files should be saved automatically when Human Resources sent the User List Form, if for some reason there are no files, use the templates from '_CreateSoxForm\Resources_' folder. You just need to copy the '_templateaccesslist.txt_' because the '_userlist.txt_' can be copied from '_CreateSoxForm_' folder. 


----
## How To Run
1. Go to folder `CreateSoxForm` and run the app `CreateSoxForm.cmd`

----
## Let's Start
>**Note:** The program will detect the user running the app and thus open the required part
##### Introducing Data (Human Resource Part)

 1. Open `userlist.txt` file from '_CreateSoxForm_' folder to add user information.
2. Add your data as in the sample below, substituting them with your own. Save the document and lunch the app to generate forms.

**Example:**
```
FirstName;LastName;PersonalEmail;SiteCode;EmployeeNum;PayRollNumber;Position;Department;Starting/Leaving;StartDate
```
 3. Wait until the APP opens an Outlook screen for the file to be sent
 

#### Introducing Data (Information Technology Part)

1. Open files `userlist.txt`, `terminateduserlist.txt` and `accesslist.txt` in folder `.\CreateSoxForm\Sox Files\Access Form\TMP` to check data.You can add user data manually for Access Sox Form / Terminated Sox Form filling.
>**Note:** _'userlist.txt' and 'terminateduserlist.txt' will have the same structure than Human Resources 'userlist.txt' file in App root folder_
2. Add all access settings / options in the file `accesslist.txt` located in same folder mentioned above using next structure
**Example:**
```
Network = Yes;Program1 = No;Program2 = No;Program3 = No;Program4 = No;Program5 = No;Program6 = No;Program6 = No;Program7 = No;Payroll = No;TimeAttendance = No;Program8 = No;Program9 = No;Program10 = No;Program11 = No;Program11 = No
```
----
## Error Codes
For troubleshooting the errors from the App, just go to `C:\Temp\company_CreateSoxForm_hhmmyyyyddMM.log` log file and look for next error codes

- WORD_OPENING
> Error opening Word file, check if all templates are in the Resources App.
- EMAIL_SEND
> Outlook was not detected or could not be opened.
- USER_LIST_TABLE_CREATION
> An error happened when User List Form was being created. Check `CreateSoxForm\userlist.txt` if it has any error.
- SENDING_EMAIL
> The Email send function has failed, check with the user if a pop-up screen has appeared or not.
- USER_LIST_FILE_TRANSFERING
> As the `userlist.txt` has errors, the `Access Form\TMP\userlist.txt` was not be created. Just do it manually.
- ACCESS_FORM_WORD_TABLE_CREATION
> Error when Access Form file was being created, please check if something was not set correctly on `userlist.txt` or `accesslist.txt`.
- ACCESS_USER_LIST
> Error looking for data files in `Access Form\TMP`. Check if the data files are there or if they have any error in the data structure. 
- TERMINATED_FORM_WORD_TABLE_CREATION
> Something happened with Terminated Form creation, check `terminateduserlist.txt` has an error in the data structure.
- ERROR_IN_SCRIPT_EXECUTION
> General Error for the script, if something has failed and is not set as any of the other error codes, this will be the error code to group them.

----

## Warning Codes
There are warnings too, but they will not interrupt the App. If you want to follow up what happening to the app time to time, look for `WARNING:XXXXXX` in the log file
- DATA_FILES_NOT_AVAILABLES
> Team Members data files were not created or they can be opened, check on `Sox Files\Access Form\TMP` folder.
- USER_LIST_FORMAT_ERROR
> The user written in the log file has an error in its data, could be a bad parameter introduced or something else, but it is not following the default structure.
----
