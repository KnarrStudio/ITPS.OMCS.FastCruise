# ITPS.OMCS.FastCruise 
This is another tool that is being added to the "IT PowerShell Open Mind Common Sense" tools set.

## The Fast Cruise 
We have to perform a manual operation check for all of the systems we have.  The term that we use for this is a "Fast Cruise". In the Navy a fast cruise is a simulated underway period that prepares the crew for life at sea.  Where this started was on a Naval Base, so people were familiar with the term and it stuck.  When it started, the fast cruise checks were manual checks using a paper check sheet.  As with most manual processes we found that the results were not always accurate, not complete and sometimes duplicated.  The initial script was part of the check sheet.  Over time the manual checks morphed to only the script and it was automated, but it didn't solve the problem of testing the physical aspect of the mouse and keyboard or even seeing if they were still at the workstation.  

The premise now is that the person doing the operational test will login to the system, thereby testing the human interface devices.  They will run the script which prompts them through the "Fast Cruise" checks and the script captures the data we are interested in.  Now we don't have to worry about those human mistakes.  Plus the time to complete the entire "Fast Cruise" 

## The Scripts
-  **Start-Fastcruise.ps1**  - This is the workhorse of the project.  You need to identify some of the paths and filenames, all these are done at the top of the script.  It also uses the **computerlocation.json** file for the selection of the locations. It helps all of spaces be typed correctly, but not needed, and will work without it.  If the file does not exist, then the program will ask for write in values.  It is highly recommended that you Edit this file for your location

-  **Export-ComputerDescription.ps1**  - This takes the output file that was created by **Start-Fastcruise.ps1** and parses the data to create a new csv file that contains the Computer Hostname and AD Descriptions.  This only needs to be used if you want to create an updated list of computers and Descriptions (the location).  One of our reasons for running the Fast Cruise is to make sure that we know where things are located.   We run it so that everyone has a quick reference sheet. At some point in the near term the AD Description will be correct and will only have to be executed when there are changes.  

-  **Convert-LocationHashToJson.ps1**  - I have found that the JSON file is a little difficult to create, but a hash table is easy, so this takes a hash table and turns it into a JSON file. 

### The Information recorded 
Information is captured through two means. Automatically based on settings on the system itself, and then manually through input the user responses and saves it in a CSV for later evaluation and proof that it has been completed. 

#### Automatically Retreived Information 
  1. Username of person completing the test  
  1. Date & Time of the test  
  1. MAC Address 
  1. Serial Number of Machine  
  1. Host Name  
  1. WSUS Search Success  
  1. WSUS Install Success  

#### User Involved Information 
  1. Software versions of your selection  
  1. Opens applications of your selection  
  1. Department 
  1. Building             
  1. Room 
  1. Desk - Starting from "A"  
  1. Phone - Must be in the 607-738-7571 format 
  1. Notes - Any notes that you want to add to the spreadsheet for later  
  1. Facility Issues - Allows the person to record issues they find in a room (Lights out, A/C not working) and stores it in a file that can be sent to the facility manager. 

#### User Customizable 
Information that is customizable for the user are the following:
1. Software Version - Returns the versions of the software requested
   1. Mozilla Firefox Version
   1. McAfee Agent Version 
   1. Adobe Test               
1. Operational Test - Opens a specific file or just the program to ensure the basic configuration is good and completed.
   1. MS Office Test
   1. Acrobat
1. Desk - This is just a way to lable the desk locaitons
    1. We use "A,B,C..."
    1. You can use anything you want to put in an Array
1. Department > Building > Room
    1. This is in the computerlocation.json file




