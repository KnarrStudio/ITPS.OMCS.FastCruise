# ITPS.OMCS.FastCruise 

## The Fast Cruise 
We have to perform a manual operation check for all of the systems we have.  The term that we use for this is a "Fast Cruise". A fast cruise is a simulated underway period that prepares the crew for life at sea.  The fast cruise checks started out as a manual checks using a check sheet.  As with most manual processes we found that the results were not always acurate, not complete and sometimes duplicated.  The initial script was part of the check sheet.  The things morphed to only the script, but that didn't solve the problem of testing the mouse and keyboard or even seeing if they are there.  As it is written now, the premis is that the person doing the operational test will login to the system, thereby testing the human interface devices, and then runs the script it captures the data we are interested in with user intevension.  The primary example is the location of the device, and making sure that it is correct. 

## The Scripts
- **Start-Fastcruise.ps1** - This is the workhorse of the project. The fast cruise used to be a very long process that took months.  This has been able to reduce the time down, to where we hope to be able to complete these tests once a month in less than a week. 
- This script captures the following and saves it in a CSV for later evaulation and proof that it has been completed: 
  1. Username of person completing the test 
  1. Date & Time of the test 
  1. Serial Number of Machine 
  1. Host Name 
  1. Software versions of your selection 
  1. Opens applications of your selection 
  1. The Building, Room, Desk, and phone number of where the workstation is located. 
 
  - All the above information is saved in a CVS file that is recreated Monthly 
  
- **computerlocation.json** - The script looks for this file and uses it for the location.  If the file does not exist, then the program will ask for write in values.  It is highly recommended that you Edit this file for your location

- **Export-ComputerDescription.ps1** - Once the "Start-FastCruise" script has been run, you will be able to parse the data in that document and create a new csv file that contains the Computer Host Name and AD Description.  This does not need to be executed, but should be for the next few months or there has been a change in the inventory of the systems.  At some point in the near term the AD Description will be correct and will only have to be executed when there are changes.  

- **Convert-LocationHashToJson.ps1** - I have found that the JSON file is a little difficuat to create, but a hash table is easy, so this takes a hash table and turns it into a JSON file.






