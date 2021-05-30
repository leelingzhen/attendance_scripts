# attendance_scripts
A simple python script to manage google sheet attendances

-ensure you have installed the relevant python libraries by 'pip install -r requirements.txt'

-the only format the script can take in is provided in attendance_format.xlsx. 
-upload the file into a google drive and turn it into a google sheet. Paste the url of the google sheet into the ATTENDANCE_URL variable. ensure that the url is a string
-ensure that the sheet sharing options are 'Anyone with the link' in the settings when the script is being run


for player profiles:
- under the team column, any number of variations are accepted 
ie '1,2,3,A,B,C,Singapore,US,Australia' are all acceptable inputs with n number of teams
- under the gender column only "M" and "F" are accepted 
- you may turn the the csv sheet into a cloud sheet as well, the same rules apply as mentioned above. you can do so by replacing player_profiles.csv with the google sheet url
