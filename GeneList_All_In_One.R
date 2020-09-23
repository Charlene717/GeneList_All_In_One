# http://www.zendei.com/article/2409.html
#�@�ΡGŪ���C�Ӥ�󧨤U��excel�A�ñN��X�֦��@�Ӥ��C
#�@��3�šG�Ĥ@�šG��󧨦W�A�ĤG�šG��󧨤���xlsx���W�A�ĤT�šGxlsx��󤤪��C��
#�N�X�᭱�������Y���G�i�ק�j�r�ˡA�h���ܦp�G�n�b�A�����W�B��Ӭq�N�X�ɡA�ݭn�i��������ק�C

##########��k�@�G�̲׳�W�O�s�b�C�Ӥ�󧨤U
rm(list=ls())

# setwd("E:/cnblogs")                                    #�]�w�u�@�ؿ��i�ק�j
setwd(getwd())
PathName = setwd(getwd())
FolderName = c("Metastasis") 

library(xlsx)

first_category_name = list.files(FolderName)            #list.files�R�O�o��"OOOP����󧨤U�Ҧ���󧨪��W�١i�ק�j
dir = paste("./Metastasis/",first_category_name,sep="")   #��paste�R�O�c�ظ��|�ܼ�dir,�Ĥ@�ťؿ����ԲӸ��|�i�ק�j
n = length(dir)                                       #Ū��dir���סA�]�N�O�G�`�@���h�֭Ӥ@�ťؿ�                                                     

n_sub<-rep(0,n)
n_sub<-as.data.frame(n_sub)
n_sub<-t(n_sub)
head(n_sub)                                          #n_sub�O�C�Ӥ@�ťؿ�(���)�U���h�֭Ӥ��A�]�N�O�G���h�֭ӤG�ťؿ��A��l�Ƭ�0�A�Ω�᭱���ާ@

##########
for(i in 1:n){         #���C�Ӥ@�ťؿ�(���)
  b=list.files(dir[i]) #b�O�C�X�C�Ӥ@�ťؿ�(���)���C��xlsx��󪺦W��
  n_sub[i]=length(b)   #�o��@�ťؿ�(���)�Uxlsx�����Ӽ�:n_sub
  
  merge_1<-read.xlsx(paste0(PathName,"/",FolderName,"Output_List.xlsx",sheetIndex=1,encoding='UTF-8'))#
  dim(merge_1)
  names(merge_1)<-c('�Ǹ�','APP','2016-01-11','2016-01-12','2016-01-13','2016-01-14','2016-01-15','2016-01-16','2016-01-17')#�ڪ����C�W�A�ھڧA���鱡�p�ק�i�ק�j
  merge_1$second_category<-'second_category'
  merge_1$first_category<-'first_category'
  merge_1<-merge_1[1,-1]    #�o�@�q���ت��OŪ���@��xlsx���˨ҡA�o��@�Ӫ�l��dataframe(���t�ƾ�)�A�K���F�᭱���s�إ�datafram���·СA�Ω�᭱���rbind����
  
  for(j in 1:n_sub[i]){     #���C�Ӥ@�ťؿ�(���)�U���C��xlsx���
    new_1<-read.xlsx(file=paste(dir[i],'/',b[j],sep=''),sheetIndex=1,encoding='UTF-8') #Ū��xlsx���
    names(new_1)<-c('�Ǹ�','APP','2016-01-11','2016-01-12','2016-01-13','2016-01-14','2016-01-15','2016-01-16','2016-01-17') #�i�ק�j
    new_1<-new_1[-1,-1]     #�]����ڼƾڻݭn�A�R���Ĥ@��M�Ĥ@�C�]�ھڹ��Ū��xlsx��󪺱��p�i��ק�^
    new_1$second_category<-substr(b[j],1,4)        #�G�ťؿ����W�٬Oxlsx�����W�C
    new_1$first_category<-first_category_name[i]   #�@�ťؿ����W�٬O����󧨦W��
    merge_1<-rbind(merge_1,new_1)
  }
  write.xlsx(merge_1,paste(dir[i],'/merge.xlsx',sep=''),row.names = F,col.names= F)#��W�O�s�b�C�Ӥ�󧨤U
}