##OneDrive updater.##

##Libraries##
library(Microsoft365R)
library(tidyverse)
library(readxl)
`%!in%`<-Negate(`%in%`)

##Inputs
FileSpreadsheet<-read_xlsx("LocalFileList.xlsx")


##Initial OneDrive object creation.##
OneDrive<-get_personal_onedrive(auth_type="device_code")

TopLevelList<-OneDrive$list_files(full_names=TRUE)
TopLevelFolderList<-filter(TopLevelList, size!=0)
TopLevelFolderList<-filter(TopLevelFolderList, isdir==TRUE)
TopLevelFolderNRow<-nrow(TopLevelFolderList)


FolderList<-filter(TopLevelList, size!=0)
FolderList<-filter(FolderList, isdir==TRUE)

##Builds a dataframe containing all folders in the OneDrive.##
##NOTE: Empty folders are not included.##
for (i in 1:nrow(FolderList)){
    Counter<-1
    SubCounter<-1
    FolderTest<-TRUE %in% FolderList[,"isdir"]
    print("TRUE")
    
    if(FolderList[i,"name"]%in% TopLevelFolderList[,"name"]){
      NextFolder<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
      NextFolderBind<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
      NextFolderBind<-filter(NextFolderBind, size!=0)
      NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
      FolderList<-rbind(FolderList, NextFolderBind)
    }


}


NotTopLevelFolderList<-FolderList[-c(1:TopLevelFolderNRow),]

for(i in 1:nrow(NotTopLevelFolderList)){
  if(i>1){
    break
  }
  j<-1
  BreakCheck<-FALSE
  FolderTest<-TRUE
  while(FolderTest){
    print(j)
    
    # if(j>nrow(NotTopLevelFolderList)){
    #   BreakCheck<-TRUE
    #   FolderTest<-FALSE
    # }
    # 
    if(j==nrow(NotTopLevelFolderList)+1){
      FolderTest<-FALSE
      break
    }
    
    if(BreakCheck){
      break
    }
    
    NextFolderBind<-try(OneDrive$list_files(NotTopLevelFolderList[j, "name"], full_names=TRUE))
    
    if(class(NextFolderBind)=="try-error"){
      error_type <- attr(result,"condition")
      
      print(class(error_type))
      print(error_type$message)
      
      if(grepl("Bad Gateway (HTTP 502).", error_type$message)){
        j<-j-1
        next
      }else if(grepl("Bad Gateway (HTTP 400).", error_type$message)){
          j<-j-1
          next
        
      }else{
        
        break
      }
    }
    
    if(nrow(NextFolderBind)==0){
      j<-j+1
      next
    }
    NextFolderBind<-filter(NextFolderBind, size!=0)
    NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
    if(nrow(NextFolderBind)==0){
      j<-j+1
      next
    }
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
TestResults<-TopLevelFolderList
TestResults<-TestResults[0,]
for(i in 1:nrow(NotTopLevelFolderList)){
  Test<-OneDrive$list_files(NotTopLevelFolderList[i, "name"], full_names=TRUE)
  Test<-filter(Test, size!=0)
  Test<-filter(Test, isdir==TRUE)
  Test<-filter(Test, Test[,"name"]%!in%FolderList[, "name"])
  if(nrow(Test)!=0){
  TestResults<-rbind(TestResults,Test)
  }
  print(i)
}


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
FileSpreadsheet<-read_xlsx("LocalFileList.xlsx")

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
  if(grepl("#",as.character(FoldersToCreate[i,"Path"]))){
    NewFileName<-str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="#","")
    file.rename(from=as.character(FoldersToCreate[i,"Path"]),
                str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="#",""))
    
    FoldersToCreate[i,"Path"]<-NewFileName
    print(paste0("Folder name changed at ", i, " row."))
  }
  
  if(grepl("%",as.character(FoldersToCreate[i,"Path"]))){
    NewFileName<-str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="%","")
    file.rename(from=as.character(FoldersToCreate[i,"Path"]),
                str_replace_all(as.character(FoldersToCreate[i,"Path"]), pattern="%",""))
    
    FoldersToCreate[i,"Path"]<-NewFileName
    print(paste0("Folder name changed at ", i, " row."))
  }
  FoldersToCreate[i, "OneDrivePath"]<-substring(as.character(FoldersToCreate[i, "Path"]),4)
  
}

##Add subset logic for folders below.##

FoldersToCreateTest<-FoldersToCreate
FoldersToActuallyCreate<-FoldersToCreate
RemovalList<-list()


for(i in 1:nrow(FoldersToCreateTest)){
  if(FoldersToCreateTest[i,"OneDrivePath"]%in%FolderList[, "name"]){
    RemovalListAdd<-i
    RemovalList<-rbind(RemovalList,RemovalListAdd)
  }
  

}

RemovalList<-sapply(RemovalList,"[[",1)
FoldersToActuallyCreate<-FoldersToActuallyCreate[-RemovalList,]
FoldersToCreate<-FoldersToActuallyCreate



for(i in 1:nrow(FoldersToCreate)){
  FolderPath<-as.character(FoldersToCreate[i,"OneDrivePath"])
  FolderCreateTry<-try(OneDrive$create_folder(FolderPath), silent=TRUE)
  
  if(class(FolderCreateTry)=="try-error"){
    error_type <- attr(FolderCreateTry,"condition")
    error_message<-error_type$message

    if(grepl("with the same name", error_message)){
      print("Folder already exists (typically an empty folder, skipping to next")
      print(i)
      next
      
    }
      
  }
  
  print(i)
}




##Files to upload
FilesToUpload<-FoldersToCreate
FilesToUpload<-FilesToUpload[0,]

for(i in 1:nrow(FoldersToCreate)){
  if(as.character(FoldersToCreate[i,"Notes"])=="FilesIn"&!is.na(FoldersToCreate[i,"Notes"])){
    FilesToUploadBind<-as.data.frame(setdiff(list.files(as.character(FoldersToCreate[i,"Path"]), full.names = TRUE), list.dirs(as.character(FoldersToCreate[i,"Path"]),recursive = FALSE, full.names = TRUE)))
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
  
  FilesToUploadBind<-as.data.frame(setdiff(list.files(as.character(FoldersToCreate[i,"Path"]), full.names = TRUE), list.dirs(as.character(FoldersToCreate[i,"Path"]),recursive = FALSE, full.names = TRUE)))
  names(FilesToUploadBind)<-"Path"
  NotesName<-"Notes"
  FilesToUploadBind<-add_column(FilesToUploadBind, !!NotesName:=NA_character_)
  FilesToUploadBind<-add_column(FilesToUploadBind, !!OneDrivePath:=NA_character_)
  
  FilesToUpload<-rbind(FilesToUpload, FilesToUploadBind)

}

for(i in 1:nrow(FilesToUpload)){
  FilesToUpload[i, "OneDrivePath"]<-substring(as.character(FilesToUpload[i, "Path"]),4)
  
}

TotalFiles<-as.numeric(nrow(FilesToUpload))
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
  
  
  FileUploadTry<-try(OneDrive$upload_file(src=as.character(FilesToUpload[i,"Path"]), dest=as.character(FilesToUpload[i,"OneDrivePath"])))
  
    if(class(FileUploadTry)=="try-error"){
      error_type <- attr(FileUploadTry,"condition")
      error_message<-error_type$message
      print(class(error_type))
      print(error_type$message)
      break

    }
  
  
  print(paste0(i, " files uploaded of ", TotalFiles, " total."))
}

