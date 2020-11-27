# BedrockBackup
Script to help with automatic backups for bedrock servers.  Script will close the server copy the world files and then reopen the server.
The script is written in vbs and create a properties file to store information about locations of folders and files

This script is fully automatic all time of backup is set during initial launch of program.

The vbs file should be installed in the same directory as your server.
When first running the program you will be prompted to setup the script. 
The default for all the folders and files will be the same directory as the vbs file. You can change the location through the prompts.
All locations will be stored in backup.properties and can be edited there.

The script will check how many vbs scripts are running and will close itself if more than one are running to prevent the script from interfering with itself.

It will then check if bedrock_server.exe is running 
  If it is it will give a 5 & 1 minute and  30 & 10-1 second warning before stopping the server, copying the files, and restarting the         server. 
  If it is not running it will copy the files and open the server. 


I could not find a way to automate backups for my bedrock server that actually worked so I decided to try and create one of my own
I wrote this in vbs because when I was researching it appeared to be able to do what I needed.
If you have any problems or if there is something you think I could change to make it better, please let me know and I will try to help if I can or change it if it needs changed.
