# http://www.zendei.com/article/2409.html
#作用：讀取每個文件夾下的excel，並將其合併成一個文件。
#共有3級：第一級：文件夾名，第二級：文件夾中的xlsx文件名，第三級：xlsx文件中的每行
#代碼後面註釋中若有：【修改】字樣，則表示如果要在你機器上運行該段代碼時，需要進行相應的修改。

##########方法一：最終單獨保存在每個文件夾下
rm(list=ls())

# setwd("E:/cnblogs")                                    #設定工作目錄
setwd(getwd())
PathName = setwd(getwd())
FolderName = c("Metastasis_Breast_DN") 

# library(xlsx)

first_category_name = list.files(FolderName)            #list.files命令得到"OOOP”文件夾下所有文件夾的名稱
dir = paste("/",FolderName,"/",first_category_name,sep="")   #用paste命令構建路徑變數dir,第一級目錄的詳細路徑
n = length(dir)                                       #讀取dir長度，也就是：總共有多少個一級目錄                                                     

n_sub<-rep(0,n)
n_sub<-as.data.frame(n_sub)
n_sub<-t(n_sub)
head(n_sub)                                          #n_sub是每個一級目錄(文件夾)下有多少個文件，也就是：有多少個二級目錄，初始化為0，用於後面的操作

##########
merge_1 <- c()

for(i in 1:n){         #對於每個一級目錄(文件夾)
  b=list.files(dir[i]) #b是列出每個一級目錄(文件夾)中每個xlsx文件的名稱
  n_sub[i]=length(b)   #得到一級目錄(文件夾)下xlsx的文件個數:n_sub
  
  for(j in 1:n_sub[i]){     #對於每個一級目錄(文件夾)下的每個txt文件
    new_1 <- read.table(file=paste(PathName,dir[i],sep=""),fill = TRUE) #讀取xlsx文件
    #    names(new_1)<-c('序號','APP','2016-01-11','2016-01-12','2016-01-13','2016-01-14','2016-01-15','2016-01-16','2016-01-17') 
    new_1<-new_1[-1:-2,1]     #因為實際數據需要（根據實際讀取txt文件的情況進行修改）

    
  }
  merge_1<- c(merge_1,new_1)
}
merge_1 <- unique(data.frame(unlist(merge_1)))
write.table(merge_1,paste(PathName,'/',FolderName,'_merge.txt',sep=''),row.names = F,col.names= F)#單獨保存在每個文件夾下
write.table(merge_1,paste(PathName,'/',FolderName,'_merge.csv',sep=''),row.names = FALSE,col.names= FALSE)#單獨保存在每個文件夾下

