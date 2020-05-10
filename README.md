# BedrockBackup
Script to help with automatic backups for bedrock servers.  Script will close the server copy the world files and then reopen the server.
The script is written in vbs and create a properties file to store information about locations of folders and files

This script is not fully automatic but instead can be ran with Task Schedular on Windows to automate the backup.

The vbs file should be installed in the same directory as your server.
When first running the program you will be prompted to setup the script. 
The default for all the folders and files will be the same directory as the vbs file. You can change the location through the prompts.
All locations will be stored in backup.properties and can be edited there.

The script will check how many vbs scripts are running and will close itself if more than one are running to prevent the script from interfering with itself.

It will then check if bedrock_server.exe is running 
  If it is it will give a 5 & 1 minute and  30 & 10-1 second warning before stoping the server, copying the files, and restartign the         server. 
If it is not running it will copy the files and open the server. 
