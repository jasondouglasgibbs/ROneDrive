##OneDrive updater.##

##Libraries##
library(Microsoft365R)
library(tidyverse)
`%!in%`<-Negate(`%in%`)

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

Test<-NextFolderBind[,"name"]%in%FolderList[, "name"]
Test
# while(FolderTest){
#   if(Counter==1){
#   NextFolder<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
#   NextFolderBind<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
#   NextFolderBind<-filter(NextFolderBind, size!=0)
#   NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
#   FolderList<-rbind(FolderList, NextFolderBind)
# 
#   FolderTest<-TRUE %in% NextFolderBind[,"isdir"]
#   Counter<-Counter+1
# 
# 
#   }
#   else{
#     NextFolder<-OneDrive$list_files(NextFolderBind[SubCounter, "name"], full_names=TRUE)
#     NextFolderBind<-OneDrive$list_files(NextFolderBind[SubCounter, "name"], full_names=TRUE)
#     NextFolderBind<-filter(NextFolderBind, size!=0)
#     NextFolderBind<-filter(NextFolderBind, isdir==TRUE)
#     FolderList<-rbind(FolderList, NextFolderBind)
# 
#     if(nrow(NextFolderBind)==0){
#       break
#     }
# 
#     if(TRUE %in% NextFolderBind[,"isdir"]){
#       FolderTest<-TRUE %in% NextFolderBind[,"isdir"]
#       next
#     }else{
#     SubCounter<-SubCounter+1
#     }
#   }
# # SubCounter<-1
# }





##Builds the initial OneDrive file list.##
FileList<-OneDrive$list_files(full_names=TRUE)
FileList<-filter(FileList, size!=0)
FileList<-filter(FileList, isdir==FALSE)

for(i in 1:nrow(FolderList)){
  if(i<=nrow(TopLevelFolderList)){
    FileListBind<-OneDrive$list_files(substring(FolderList[i, "name"],2), full_names=TRUE)
    FileListBind<-filter(FileListBind, size!=0)
    FileListBind<-filter(FileListBind, isdir==FALSE)
    if(nrow(FileListBind)==0){
      next
    }else{
    FileList<-rbind(FileList, FileListBind)
    print("TRUE")
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
  print("TRUE")
  
}

