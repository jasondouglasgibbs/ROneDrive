##Script that receives a spreadsheet with local directories to upload to OneDrive.##
##Allows the user to also specify "AllFiles" or "FilesIn" which allows the user##
##to specify if they only want to select certain portions (this would be useful)##
##for larger folders that may need to be split across OneDrives.##
##Author: Jason D. Gibbs##
##Email: jasondouglasgibbs@gmail.com##

##Limitations.##
##-This script is designed to update OneDrive for users that use it as "static"##
##storage. This means that you use OneDrive to upload files, but you don't typically##
##update those files.##
##-The script basically determines "updates" are present by comparing file names##
##and paths, not by "modified" time.##
##-The OneDrive your using should mirror the initial structure of the local files##
##you are uploading. You must at least have a top level folder that contains at##
##least one file.##

##Libraries##
library(Microsoft365R)
library(tidyverse)
library(readxl)
##Creates an operator for the logic "not in".##
`%!in%`<-Negate(`%in%`)

##Inputs##
##This is the spreadsheet that contains your directories you want to upload to##
##OneDrive/make sure are updated in OneDrive.##
FileSpreadsheet<-read_xlsx("LocalFileList.xlsx")


##Initial OneDrive object creation.##
##Select the link in the console, load the web page, and input the code listed##
##in the console.##
OneDrive<-get_personal_onedrive(auth_type="device_code")
##Lists the top level folders that exist on the OneDrive. This is necessary##
##because of structure of paths for top level folders.##
##
TopLevelList<-OneDrive$list_files(full_names=TRUE)
TopLevelFolderList<-filter(TopLevelList, size!=0)
TopLevelFolderList<-filter(TopLevelFolderList, isdir==TRUE)
TopLevelFolderNRow<-nrow(TopLevelFolderList)


FolderList<-filter(TopLevelList, size!=0)
FolderList<-filter(FolderList, isdir==TRUE)

##Builds a dataframe containing folders within the top level folders.##
##NOTE: Empty folders are not included.##
for (i in 1:nrow(FolderList)){
    Counter<-1
    SubCounter<-1
    FolderTest<-TRUE %in% FolderList[,"isdir"]
    
    if(FolderList[i,"name"]%in% TopLevelFolderList[,"name"]){
      NextFolder<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
      NextFolderBind<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
      NextFolderBind<-filter(NextFolderBind, size!=0)
      NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
      FolderList<-rbind(FolderList, NextFolderBind)
    }


}

##Creates a DF that does not contain any top level folders.## 
NotTopLevelFolderList<-FolderList[-c(1:TopLevelFolderNRow),]

##For loop that recursively cycles through the OneDrive to build a folder structure.##
for(i in 1:nrow(NotTopLevelFolderList)){
  if(i>1){
    break
  }
  j<-1
  BreakCheck<-FALSE
  FolderTest<-TRUE
  while(FolderTest){
    print(j)
    
    ##Stops the while loop once all folders have been found.##
    if(j==nrow(NotTopLevelFolderList)+1){
      FolderTest<-FALSE
      break
    }
    
    ##Trys to list OneDrive folders in a specified directory.##
    NextFolderBind<-try(OneDrive$list_files(NotTopLevelFolderList[j, "name"], full_names=TRUE))
    
    ##Error logic for OneDrive query.##
    if(class(NextFolderBind)=="try-error"){
      error_type <- attr(NextFolderBind,"condition")
      error_message<-error_type$message
      
      ##HTTP 403 error that seems to correspond to connection issues with OneDrive.##
      ##Untested.##
      if(grepl("Service Unavailable", error_message)){
        z<-1
        ##While loop that retries the connection up to 10 times.##
        while(z<10&class(NextFolderBind)=="try-error"){
          NextFolderBind<-try(OneDrive$list_files(NotTopLevelFolderList[j, "name"], full_names=TRUE))
          print(paste0("Unable to reach OneDrive, retry ", z, " of ten times."))
          error_type <- attr(NextFolderBind,"condition")
          error_message<-error_type$message
          z<-z+1
        }
        
        if(z==10){
          break
        }
        
        if(z!=10){
          next
        }
      }
      
      break  
    }
    
    ##Cycles to next iteration if the retrieved DF is empty.##
    if(nrow(NextFolderBind)==0){
      j<-j+1
      next
    }
    ##Filters out files and size zero folders.##
    NextFolderBind<-filter(NextFolderBind, size!=0)
    NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
    ##Cycles to next iteration if the retrieved DF is empty.##
    if(nrow(NextFolderBind)==0){
      j<-j+1
      next
    }
    ##Iterates if all entries in the current DF already exist in the overall##
    ##folder list.##
    if(all(NextFolderBind[,"name"]%in%FolderList[, "name"])){
      j<-j+1
      next
    }
    FolderList<-rbind(FolderList, NextFolderBind)
 
    NotTopLevelFolderList<-rbind(NotTopLevelFolderList, NextFolderBind)
    rownames(NotTopLevelFolderList)<-NULL
  }
  

}

##Tests to see if all folders are included.
# TestResults<-TopLevelFolderList
# TestResults<-TestResults[0,]
# for(i in 1:nrow(NotTopLevelFolderList)){
#   Test<-OneDrive$list_files(NotTopLevelFolderList[i, "name"], full_names=TRUE)
#   Test<-filter(Test, size!=0)
#   Test<-filter(Test, isdir==TRUE)
#   Test<-filter(Test, Test[,"name"]%!in%FolderList[, "name"])
#   if(nrow(Test)!=0){
#   TestResults<-rbind(TestResults,Test)
#   }
#   print(i)
# }


##Builds the initial OneDrive file list.##
FileList<-OneDrive$list_files(full_names=TRUE)
FileList<-filter(FileList, size!=0)
FileList<-filter(FileList, isdir==FALSE)

for(i in 1:nrow(FolderList)){
  print(i)
  if(i<=nrow(TopLevelFolderList)){
    FileListBind<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
    FileListBind<-filter(FileListBind, size!=0)
    FileListBind<-filter(FileListBind, isdir==FALSE)
    if(nrow(FileListBind)==0){
      next
    }else{
    FileList<-rbind(FileList, FileListBind)
    next
    }
  }
  
  FileListBind<-OneDrive$list_files(FolderList[i,"name"],full_names=TRUE)
  FileListBind<-filter(FileListBind, size!=0)
  FileListBind<-filter(FileListBind, isdir==FALSE)
  if(nrow(FileListBind)==0){
    next
  }
  FileList<-rbind(FileList, FileListBind)
  
}



##Local Folder/File portion.##
FoldersToCreate<-FileSpreadsheet

for (i in 1:nrow(FoldersToCreate)){
  if(as.character(FoldersToCreate[i,"Notes"])=="FilesIn"&!is.na(FoldersToCreate[i,"Notes"])){
    next
  }
  
  FoldersToBind<-as.data.frame(list.dirs(as.character(FoldersToCreate[i,"Path"])))
  if(nrow(FoldersToBind)>1){
  FoldersToBind<-as.data.frame(FoldersToBind[-c(1),])
  }
  names(FoldersToBind)<-"Path"
  NotesName<-"Notes"
  FoldersToBind<-add_column(FoldersToBind, !!NotesName:=NA_character_)
  FoldersToBind<-filter(FoldersToBind,FoldersToBind[,"Path"]%!in%FoldersToCreate[, "Path"] )
  FoldersToCreate<-rbind(FoldersToCreate, FoldersToBind)
  
  
}

OneDrivePath<-"OneDrivePath"
FoldersToCreate<-add_column(FoldersToCreate, !!OneDrivePath:=NA_character_)


for(i in 1:nrow(FoldersToCreate)){
  ##Changes local folder name and updates folder DF if the current folder##
  ##contains a "#" symbol in its name.##
  if(grepl("#",as.character(FoldersToCreate[i,"Path"]))){
    NewFileName<-str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="#","")
    file.rename(from=as.character(FoldersToCreate[i,"Path"]),
                str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="#",""))
    
    FoldersToCreate[i,"Path"]<-NewFileName
    print(paste0("Folder name changed at ", i, " row."))
  }
  
  ##Changes local folder name and updates folder DF if the current folder##
  ##contains a "%" symbol in its name.##
  if(grepl("%",as.character(FoldersToCreate[i,"Path"]))){
    NewFileName<-str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="%","")
    file.rename(from=as.character(FoldersToCreate[i,"Path"]),
                str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="%",""))
    
    FoldersToCreate[i,"Path"]<-NewFileName
    print(paste0("Folder name changed at ", i, " row."))
  }
  FoldersToCreate[i, "OneDrivePath"]<-substring(as.character(FoldersToCreate[i, "Path"]),4)
  
}


##Compares the local and OneDrive folder lists and removes any folders that exist##
##both in the OneDrive and the local drive. The resulting list will add folders that##
##exist locally but NOT on the OneDrive.##
##NOTE: Empty folders on the OneDrive are a special case. The script will attempt##
##to create them later on, but the error will be prevented through use of Try().##
FoldersToCreateTest<-FoldersToCreate
FoldersToActuallyCreate<-FoldersToCreate
RemovalList<-list()

for(i in 1:nrow(FoldersToCreateTest)){
  if(FoldersToCreateTest[i,"OneDrivePath"]%in%FolderList[, "name"]){
    RemovalListAdd<-i
    RemovalList<-rbind(RemovalList,RemovalListAdd)
  }
  

}
##Simplifies list down to a single dimension.##
RemovalList<-sapply(RemovalList,"[[",1)
FoldersToActuallyCreate<-FoldersToActuallyCreate[-RemovalList,]
FoldersToCreate<-FoldersToActuallyCreate

##Prevents a probable error if there are no folders to create.##
if(nrow(FoldersToCreate)!=0){
  for(i in 1:nrow(FoldersToCreate)){
    FolderPath<-as.character(FoldersToCreate[i,"OneDrivePath"])
    FolderCreateTry<-try(OneDrive$create_folder(FolderPath), silent=TRUE)
  
    if(class(FolderCreateTry)=="try-error"){
      error_type <- attr(FolderCreateTry,"condition")
      error_message<-error_type$message

      if(grepl("with the same name", error_message)){
        print("Folder already exists (typically an empty folder, skipping to next).")
        print(i)
        next
      }
    
      if(grepl("Service Unavailable", error_message)){
        z<-1
        while(z<10&class(FolderCreateTry)=="try-error"){
          FolderCreateTry<-try(OneDrive$create_folder(FolderPath), silent=TRUE)
          print(paste0("Unable to reach OneDrive, retry ", z, " of ten times."))
          error_type <- attr(FolderCreateTry,"condition")
          error_message<-error_type$message
          z<-z+1
        }
      
        if(z==10){
          break
        }
      
        if(z!=10){
          next
        }
      }
    
      break  
    }
  
   print(i)
  }
}else{
  print("No folders needed to be created.")
}




##Files to upload
FilesToUpload<-FoldersToCreate
FilesToUpload<-FilesToUpload[0,]

for(i in 1:nrow(FoldersToCreateTest)){
  if(as.character(FoldersToCreateTest[i,"Notes"])=="FilesIn"&!is.na(FoldersToCreateTest[i,"Notes"])){
    FilesToUploadBind<-as.data.frame(setdiff(list.files(as.character(FoldersToCreateTest[i,"Path"]), full.names = TRUE), list.dirs(as.character(FoldersToCreateTest[i,"Path"]),recursive = FALSE, full.names = TRUE)))
    names(FilesToUploadBind)<-"Path"
    NotesName<-"Notes"
    FilesToUploadBind<-add_column(FilesToUploadBind, !!NotesName:=NA_character_)
    FilesToUploadBind<-add_column(FilesToUploadBind, !!OneDrivePath:=NA_character_)
    FilesToUpload<-rbind(FilesToUpload, FilesToUploadBind)
    next
    }
  
  # if(as.character(FoldersToCreate[i,"Notes"])=="AllFiles"&!is.na(FoldersToCreate[i,"Notes"])){
  #   FilesToUploadBind<-as.data.frame(setdiff(list.files(as.character(FoldersToCreate[1,"Path"]), full.names = TRUE), list.dirs(as.character(FoldersToCreate[1,"Path"]),recursive = FALSE, full.names = TRUE)))
  #   names(FilesToUploadBind)<-"Path"
  #   NotesName<-"Notes"
  #   FilesToUploadBind<-add_column(FilesToUploadBind, !!NotesName:=NA_character_)
  #   FilesToUploadBind<-add_column(FilesToUploadBind, !!OneDrivePath:=NA_character_)
  #   FilesToUpload<-rbind(FoldersToCreate, FoldersToBind)
  #   next
  # }
  
  FilesToUploadBind<-as.data.frame(setdiff(list.files(as.character(FoldersToCreateTest[i,"Path"]), full.names = TRUE), list.dirs(as.character(FoldersToCreateTest[i,"Path"]),recursive = FALSE, full.names = TRUE)))
  names(FilesToUploadBind)<-"Path"
  NotesName<-"Notes"
  FilesToUploadBind<-add_column(FilesToUploadBind, !!NotesName:=NA_character_)
  FilesToUploadBind<-add_column(FilesToUploadBind, !!OneDrivePath:=NA_character_)
  
  FilesToUpload<-rbind(FilesToUpload, FilesToUploadBind)

}

for(i in 1:nrow(FilesToUpload)){
  FilesToUpload[i, "OneDrivePath"]<-substring(as.character(FilesToUpload[i, "Path"]),4)
  
}

##Compares the local and OneDrive file lists and removes any files from the DF that exist##
##in both locations. Files that exist locally but not on the OneDrive will be preserved.##
RemovalList<-list()

for(i in 1:nrow(FilesToUpload)){
  if(FilesToUpload[i,"OneDrivePath"]%in%FileList[, "name"]){
    RemovalListAdd<-i
    RemovalList<-rbind(RemovalList,RemovalListAdd)
  }
  
  
}

RemovalList<-sapply(RemovalList,"[[",1)
FilesToUpload<-FilesToUpload[-RemovalList,]

TotalFiles<-as.numeric(nrow(FilesToUpload))
if(TotalFiles!=0){
  for(i in 1:nrow(FilesToUpload)){
    if(grepl("#",as.character(FilesToUpload[i,"Path"]))){
      NewFileName<-str_replace_all(as.character(FilesToUpload[i,"Path"]), pattern="#","")
      file.rename(from=as.character(FilesToUpload[i,"Path"]),
                str_replace_all(as.character(FilesToUpload[i,"Path"]), pattern="#",""))
      FilesToUpload[i,"Path"]<-NewFileName
      FilesToUpload[i, "OneDrivePath"]<-substring(NewFileName,4)
      print(paste0("File name changed at ", i, " row."))
    }
  
    if(grepl("%",as.character(FilesToUpload[i,"Path"]))){
      NewFileName<-str_replace_all(as.character(FilesToUpload[i,"Path"]), pattern="%","")
      file.rename(from=as.character(FilesToUpload[i,"Path"]),
                str_replace_all(as.character(FilesToUpload[i,"Path"]), pattern="%",""))
    
      FilesToUpload[i,"Path"]<-NewFileName
      FilesToUpload[i, "OneDrivePath"]<-substring(NewFileName,4)
      print(paste0("File name changed at ", i, " row."))
    }
  
  FileUploadTry<-try(OneDrive$upload_file(src=as.character(FilesToUpload[i,"Path"]), dest=as.character(FilesToUpload[i,"OneDrivePath"])), silent=TRUE)
    ##Error logic for OneDrive query.##
    if(class(FileUploadTry)=="try-error"){
      error_type <- attr(FileUploadTry,"condition")
      error_message<-error_type$message
        
      if(grepl("with the same name", error_message)){
        print("File  already exists, skipping to next")
        print(i)
        next
        
      }
      
      
      if(grepl("Service Unavailable", error_message)){
        z<-1
        while(z<10&class(FileUploadTry)=="try-error"){
          FileUploadTry<-try(OneDrive$upload_file(src=as.character(FilesToUpload[i,"Path"]), dest=as.character(FilesToUpload[i,"OneDrivePath"])), silent=TRUE)
          print(paste0("Unable to reach OneDrive, retry ", z, " of ten times."))
          error_type <- attr(FileUploadTry,"condition")
          error_message<-error_type$message
          z<-z+1
        }
        
        if(z==10){
          break
        }
        
        if(z!=10){
          print(paste0(i, " files uploaded of ", TotalFiles, " total."))
          next
        }
      }
      
      break

    }
  
  print(paste0(i, " files uploaded of ", TotalFiles, " total."))
  }
}else{
  print("OneDrive is up-to-date with local drive.")
}

