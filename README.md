# R OneDrive Updating Script

This R script updates the contents of a personal OneDrive. The targeted user is someone who uses OneDrive as static storage, not users who use it as active storage with documents they modify often.

## Update Definition

For purposes of this script, "update" means that new files that have been added locally that do not exist on the OneDrive will be uploaded. This script does not upload modified files (unless the name of the file was also modified).

## Instructions

### Azure Login Prompt

When using the script, a popup similar to the image below will occur. I recommend selecting "no" as this will allow you to change between multiple OneDrives on different Microsoft accounts.

![Azure Login Image](AzureLogin.png)

### OneDrive Directory Hierarchy

To satisfy some of the script requirements, your OneDrive must have a folder with at least one file in it. Additionally, the folder should match the hierarchy of your local drive. For instance, on my local hard drive, I keep all my files in a folder called "My Files", so I named the OneDrive folder "My Files".

### Local File List Spreadsheet

Add the full path to the folders you want to upload/update. In the notes column, you can enter "FilesIn". This will upload any files in the folder, but will not upload the contents of any folders in the directory. This is useful if you need to break up a folder or don't want to upload certain directories within a folder.
