rm(list=ls())
require(plyr)
require(reshape2)
require(ggplot2)
require(xlsx)
require(multcomp)
require(epitools)

dat_in <- read.csv("C:/Users/Child Study/kd/20141125scid2/data/dataCorrectedByEva.csv")
dat_res <- read.csv("C:/Users/Child Study/kd/20141125scid2/data/response.csv")
dat_a<-merge(dat_in[,-1],dat_res[,c(-1,-3,-4,-6)],by="StudyID")

#group symp into 1:avoidant or paran 2:other
dat_a$avopar.base<-rowSums(dat_a[,c(4,9)])
dat_a$other.s.base<-rowSums(dat_a[,c(5:8,10:16)])
dat_a$avopar.v14<-rowSums(dat_a[,c(17,22)])
dat_a$other.s.v14<-rowSums(dat_a[,c(18:21,23:29)])


dat_a$avoid_or_paran.base<-ifelse(dat_a[,31]>=1,1,dat_a[,31])
dat_a$other.base<-ifelse(dat_a[,32]>=1,1,dat_a[,32])
dat_a$exclu_other.base<-ifelse(dat_a[,35]==1,0,ifelse(dat_a[,36]==1,1,dat_a[,36]))
dat_a$avoid_or_paran.v14<-ifelse(dat_a[,33]>=1,1,dat_a[,33])
dat_a$other.v14<-ifelse(dat_a[,34]>=1,1,dat_a[,34])
dat_a$exclu_other.v14<-ifelse(dat_a[,38]==1,0,ifelse(dat_a[,39]==1,1,dat_a[,39]))

dat_a$overall.base<-ifelse(dat_a$exclu_other.base==1,"Dx.Other",
                           ifelse(dat_a$avoid_or_paran.base==1,"Avoid or paran",
                                  ifelse(dat_a$exclu_other.base==0&dat_a$avoid_or_paran.base==0,"None.Dx",NA)))

dat_a$overall.v14<-ifelse(dat_a$exclu_other.v14==1,"Dx.Other",
                           ifelse(dat_a$avoid_or_paran.v14==1,"Avoid or paran",
                                  ifelse(dat_a$exclu_other.v14==0&dat_a$avoid_or_paran.v14==0,"None.Dx",NA)))

#cluster if any dx
dat_a$anydx.base<-as.numeric(apply(dat_a[,4:16],1,function(x){any(x>=1)}))
dat_a$anydx.v14<-as.numeric(apply(dat_a[,17:29],1,function(x){any(x>=1)}))
#cluster group A by  Para, schizTy, schizId
dat_a$par_schTy_schID.base<-as.numeric(apply(dat_a[,c(9:11)],1,function(x){any(x>=1)}))
dat_a$par_schTy_schID.v14<-as.numeric(apply(dat_a[,c(22:24)],1,function(x){any(x>=1)}))
#cluster group b by  bord narc antisocial historic
dat_a$bor_nar_ant_his.base<-as.numeric(apply(dat_a[,c(12:15)],1,function(x){any(x>=1)}))
dat_a$bor_nar_ant_his.v14<-as.numeric(apply(dat_a[,c(25:28)],1,function(x){any(x>=1)}))
#cluster gourp c by avoidant dependent ocd
dat_a$avo_dep_ocd.base<-as.numeric(apply(dat_a[,c(4:6)],1,function(x){any(x>=1)}))
dat_a$avo_dep_ocd.v14<-as.numeric(apply(dat_a[,c(17:19)],1,function(x){any(x>=1)}))


#############################################
overtab<-table(dat_a$overall.base,dat_a$overall.v14)
fisher.p<-round(fisher.test(overtab)[[1]],4)
chisq.p<-round(chisq.test(overtab)[[3]],4)
overtab<-as.data.frame.matrix(overtab)
overtab$sum.base<-as.numeric(c(sum(overtab[1,]),sum(overtab[2,]),sum(overtab[3,])))
overtab$prop.dx.base<-paste(round(overtab$sum.base/sum(overtab[,1:3]),4)*100,"%",sep="")

overtab[4,]<-c(sum(overtab[1:3,1]),sum(overtab[1:3,2]),sum(overtab[1:3,3]),sum(overtab[1:3,1:3]),"100%")
rownames(overtab)[4]<-"Total"
overtab$p_value<-c("Fisher",fisher.p,"Chisq",chisq.p)

#over.base.res<-tab(dat_a$overall.base[!is.na(dat_a$overall.base)&!is.na(dat_a$overall.v14),],)
################################################

dat_sas<-dat_a
dat_sas[is.na(dat_a)]<-' '
dat_sas$overall.base<-as.factor(dat_sas$overall.base)
dat_sas$overall.v14<-as.factor(dat_sas$overall.v14)



write.csv(dat_sas,"C:/Users/Child Study/kd/20141125scid2/data/recoded.data.csv",row.names=F)
dat_s<-dat_a
#reorder and rename the data.
dat_a<-subset(dat_a,select=-c(avopar.base,other.s.base,avopar.v14,other.s.v14))
dat_a<-dat_a[,c(1:3,grep("base",colnames(dat_a)),grep(".v14",colnames(dat_a)),grep("CAPS",colnames(dat_a)))]
#colnames(dat_a)[c(17:19,33:35)]<-c("avoid_or_paran.base","other.base","exclu_other.base",
#                                   "avoid_or_paran.v14","other.v14","exclu_other.v14")

dat_a[,4:ncol(dat_a)]<-data.frame(sapply(dat_a[,4:ncol(dat_a)],function(x){
  factor(x,levels=0:1)
}))

dat_a<-dat_a[,-c(19,20,40,41)]
#levels(dat_at[,3])
#freq table with no missing
obstab<-function(dat){
  tab_woNA<-list()
  for(i in 4:22){
    tab_woNA[[i-3]]<-table(factor(dat[,i],levels=0:1),
                           factor(dat[,i+19],levels=0:1))
  }
  #table into data.frame
  
  dat_obs<-ldply(tab_woNA,function(x){
    cbind(x[1,1],x[2,1],x[1,2],x[2,2],
          sum(x[2,])/sum(x),sum(x[,2])/sum(x),x[2,1]/sum(x),
          ifelse(sum(x)==0,NA,round(fisher.test(x)[[1]],4)),
          ifelse(sum(x)==0,NA,round(chisq.test(x)[[3]],4)))
  })
  colnames(dat_obs)<-c("no.base_no.v14","dx.base_no.v14","no.base_dx.v14","dx.base_dx.v14",
                       "prop.dx.base","prop.dx.v14","lost.dx.v14","fisher.p-value","chisq.p-value")
  #caculate prop of dx @base/v14 and lost@v14
  #dat_obs$prop.dx.base<-(dat_obs[,2]+dat_obs[,4])/(rowSums(dat_obs[,1:4]))
  #dat_obs$prop.dx.v14<-(dat_obs[,3]+dat_obs[,4])/(rowSums(dat_obs[,1:4]))
  #dat_obs$lost.dx.v14<-dat_obs[,2]/(rowSums(dat_obs[,1:4]))
  #symptom
  dat_obs$Symptom<-substr(colnames(dat)[4:22],1,nchar(colnames(dat)[4:22])-5)
  
  #freq table of each symp & each time by response
  #redundent (with missing value)
  tab_res<-list()
  for(i in 4:41){
    tab_res[[i-3]]<-table(factor(dat[,i],levels=0:1),
                          factor(dat[,"CAPS_Response_V14"],levels=0:1))
  }
  #table into data.frame
  dat_obs_res<-ldply(tab_res,function(x){
    cbind(x[1,1],x[2,1],x[1,2],x[2,2],
          x[1,2]/sum(x[1,]),x[2,2]/sum(x[2,]),
          ifelse(sum(x)==0,NA,round(fisher.test(x)[[1]],4)))
  })
  dat_obs_res<-cbind(dat_obs_res[1:19,],dat_obs_res[20:38,])
  colnames(dat_obs_res)<-c("no.base_no.res.v14","dx.base_no.res.v14","no.base_res.v14","dx.base_res.v14",
                           "prop.res/no.base","prop.res/dx.base","fisher.p-value.base",
                           "no.v14_no.res.v14","dx.v14_no.res.v14","no.v14_res.v14","dx.v14_res.v14",
                           "prop.res/no.v14","prop.res/dx.v14","fisher.p_value.v14")
  #calculate prop of dx @ each time by response
  
  #freq table of each symp & each time by response(no missing value)
  dat_a_woNa<-subset(dat,!is.na(dat[,4])&!is.na(dat[,23]))
  tab_res_woNa<-list()
  for(i in 4:41){
    tab_res_woNa[[i-3]]<-table(factor(dat_a_woNa[,i],levels=0:1),
                               factor(dat_a_woNa[,"CAPS_Response_V14"],levels=0:1))
  }
  dat_obs_res_woNa<-ldply(tab_res_woNa,function(x){
    cbind(x[1,1],x[2,1],x[1,2],x[2,2],
          x[1,2]/sum(x[1,]),x[2,2]/sum(x[2,]),
          ifelse(sum(x)==0,NA,round(fisher.test(x)[[1]],4)),
          ifelse(sum(x)==0,NA,round(chisq.test(x)[[3]],4)))
  })
  dat_obs_res_woNa<-cbind(dat_obs_res_woNa[1:19,],dat_obs_res_woNa[20:38,])
  colnames(dat_obs_res_woNa)<-c("no.base_no.res.v14","dx.base_no.res.v14","no.base_res.v14","dx.base_res.v14",
                                "prop.res/no.base","prop.res/dx.base","fisher.p-value.base","chisq.p-value.base",
                                "no.v14_no.res.v14","dx.v14_no.res.v14","no.v14_res.v14","dx.v14_res.v14",
                                "prop.res/no.v14","prop.res/dx.v14","fisher.p_value.v14","chisq.p-value.v14")
  
  
  ##dx@base dx or not @v14 related to response
  #freq and test fisher test
  ft<-function(i){
     dat_t<-subset(dat,dat[,i]==1)
    tb<-table(factor(dat_t[,i+19],levels=0:1),factor(dat_t[,"CAPS_Response_V14"],levels=0:1))
    if(sum(tb)==0){
      fisher.pval<-NA
      odds.ratio<-NA
      fisher.ci<-NA
      chisq.pval<-NA
    }else{
      fisher.pval<-round(fisher.test(tb[,c(2,1)])[[1]],4)
      odds.ratio<-round(fisher.test(tb[,c(2,1)])[[3]],4)
      fisher.ci<-paste("(",round(fisher.test(tb[,c(2,1)])[[2]][1],4),
                       ",",round(fisher.test(tb[,c(2,1)])[[2]][2],4),")",sep="")
      chisq.pval<-round(chisq.test(tb)[[3]],4)
    }
    result<-cbind(substr(colnames(dat_t)[i],1,nchar(colnames(dat_t)[i])-5),tb[1,1],tb[1,2],tb[2,1],tb[2,2],
                  tb[1,2]/sum(tb[1,]),tb[2,2]/sum(tb[2,]),
                  fisher.pval,odds.ratio,fisher.ci,chisq.pval)
    return(result)
  }
  
  dat_basebyres<-as.data.frame(t(sapply(4:22,function(x){
    ft(x)
  })),stringsAsFactors=FALSE)
  colnames(dat_basebyres)<-c("Symptom","no.v14_no.resp","no.v14_resp","dx.v14_no.resp","dx.v14_resp",
                             "prop.res/base.no.v14","prop.res/base.dx.v14",
                             "fisher_p.value","Odds Ratio(res in lose/no.lose)","95% CI","chi-squar_p.value")
  
  dat_obs_res_woNa<-cbind(dat_obs_res_woNa,dat_basebyres)
  #dat_obs<-cbind(dat_obs,dat_obs_res_woNa)
  #dat_obs_res_woNa$Symptom<-dat_obs$Symptom
  tab_gain<-list()
  for(i in 4:22){
    dat_gain<-subset(dat_a_woNa,dat_a_woNa[,i]==0&dat_a_woNa[,i+19]==1)
    tab_gain[[i-3]]<-table(factor(dat_gain[,i+19],levels=1),
                           factor(dat_gain[,"CAPS_Response_V14"],levels=0:1))
  }
  #table into data.frame
  
  dat_gain<-ldply(tab_gain,function(x){
    cbind(x[1,1],x[1,2],
          sum(x[1,2])/sum(x)
          )
  })
  colnames(dat_gain)<-c("no.res_gain","res_gain","prop.res/gain")
  
  dat_obs[c(2,6),]<-dat_obs[c(6,2),]
  dat_obs[c(-1:-2,-14:-19),]<-dat_obs[order(-dat_obs[c(-1:-2,-14:-19),]$prop.dx.base)+2,]
  
  dat_obs_res_woNa<-dat_obs_res_woNa[match(dat_obs$Symptom,dat_obs_res_woNa$Symptom),]
  dat_gain<-cbind(Symptom=dat_obs$Symptom,dat_gain)
  result<-list(dat_obs,dat_obs_res_woNa,dat_gain)
  return(result)
}
overall<-obstab(dat_a)
dat_obs<-overall[[1]]
dat_obs_res_woNa<-overall[[2]]

tp<-melt(dat_obs[,c(5:7,10)],id.vars="Symptom")
tp1<-melt(dat_obs_res_woNa[,c(5,6,13,14,17,22,23)],id.vars="Symptom")
tp1$value<-as.numeric(tp1$value)
#h<-hist(as.matrix(dat_obs[,5]))
#h$density=h$counts/sum(h$counts)*100
#plot(h,freq=F)


trtr<-c(dat_obs$Symptom[1:2],"  ",dat_obs$Symptom[3:(length(dat_obs$Symptom)-6)])
barp<-function(dat,limlab,lim,name,subname,labangle,labsize){
  rp<-ggplot(dat,aes(x=Symptom,y=value,fill=variable))+ylim(0,lim)+
    geom_bar(stat="identity",position='dodge',
          alpha=0.8)+
    scale_x_discrete(limits=limlab)+theme_bw()+
    theme(axis.text.x=element_text(face = "bold",angle=labangle,size=labsize))+
    ggtitle(bquote(atop(.(name), atop(italic(.(subname)), "")))) 
  return(rp)
}

baselose<-subset(tp,variable%in%c("prop.dx.base","lost.dx.v14")&Symptom%in%dat_obs$Symptom[1:13])
base<-subset(tp,variable%in%c("prop.dx.base")&Symptom%in%dat_obs$Symptom[1:13])
lose<-subset(tp,variable%in%c("lost.dx.v14")&Symptom%in%dat_obs$Symptom[1:13])
baselose_c<-subset(tp,variable%in%c("prop.dx.base","lost.dx.v14")&Symptom%in%c("avoid_or_paran","other","anydx"))
base_c<-subset(tp,variable%in%c("prop.dx.base")&Symptom%in%c("avoid_or_paran","other","anydx"))
lose_c<-subset(tp,variable%in%c("lost.dx.v14")&Symptom%in%c("avoid_or_paran","other","anydx" ))

baselose_cl<-subset(tp,variable%in%c("prop.dx.base","lost.dx.v14")&Symptom%in%dat_obs$Symptom[17:19])
base_cl<-subset(tp,variable%in%c("prop.dx.base")&Symptom%in%dat_obs$Symptom[17:19])
lose_cl<-subset(tp,variable%in%c("lost.dx.v14")&Symptom%in%dat_obs$Symptom[17:19])



resbase<-subset(tp1,variable%in%c("prop.res/no.base","prop.res/dx.base","prop.res/no.v14","prop.res/dx.v14")&
                  Symptom%in%dat_obs$Symptom[1:13])

resbase_c<-subset(tp1,variable%in%c("prop.res/no.base","prop.res/dx.base","prop.res/no.v14","prop.res/dx.v14")&
                    Symptom%in%c("avoid_or_paran","other","anydx"))
resbase_cl<-subset(tp1,variable%in%c("prop.res/no.base","prop.res/dx.base","prop.res/no.v14","prop.res/dx.v14")&
                     Symptom%in%dat_obs$Symptom[17:19])




##losing dx related to trt
#sum(interaction(dat_a[,4],dat_a[,17],dat_a[,3])=="1.0.Relax",na.rm=T)

ft_trt<-function(method,i,l){
  if(method=="overall"){
    dat_t<-subset(dat_a,!is.na(dat_a[,4])&!is.na(dat_a[,23]))
    j<-i
    l<-2
    sym<-NA
  }else if(method=="lose"){
    dat_t<-subset(dat_a,!is.na(dat_a[,4])&!is.na(dat_a[,23])&dat_a[,i]==1)
    j<-i+19
    l<-1
    sym<-substr(colnames(dat_t)[i],1,nchar(colnames(dat_t)[i])-5)
  }
  tb<-table(factor(dat_t[,j],levels=0:1),factor(dat_t[,3],levels=c("IPT","PE","Relax")))
  if(sum(tb)==0){
    fisher.pval<-NA
    chisq.pval<-NA
    }else{
    fisher.pval<-round(fisher.test(tb)[[1]],4)
    chisq.pval<-round(chisq.test(tb)[[3]],4)
  }
  result<-cbind(sym,tb[1,1],tb[2,1],tb[1,2],tb[2,2],tb[1,3],tb[2,3],
                tb[l,1]/sum(tb[,1]),tb[l,2]/sum(tb[,2]),tb[l,3]/sum(tb[,3]),
                tb[l,1]/sum(tb[1,]),tb[l,2]/sum(tb[1,]),tb[l,3]/sum(tb[1,]),
                fisher.pval,chisq.pval)
  return(result)
}
dat_basebytrt<-as.data.frame(t(sapply(4:22,function(x){
  ft_trt("lose",x)
})),stringsAsFactors=FALSE)
colnames(dat_basebytrt)<-c("Symptom","lose.v14_IPT","not.lose.v14_IPT","lose.v14_PE",
                           "not.lose.v14_PE","lose.v14_Relax","not.lose.v14_Relax",
                           "prop.lose.v14_IPT","prop.lose.v14_PE","prop.lose.v14_Relax",
                           "prop.IPT/lose","prop.PE/lose","prop.Relax/lose",
                           "fisher_p.value","chi-square_p.value")

dat_basebytrt<-dat_basebytrt[match(dat_obs$Symptom,dat_basebytrt$Symptom),]

#plot of lose by treatment
tp_trt<-melt(dat_basebytrt[c(1,8:10)],id.vars="Symptom")
tp_trt$value<-as.numeric(tp_trt$value)

#Observe table for treatment=IPT
obs_res_ipt<-obstab(subset(dat_a,TherapyCode=="IPT"))
obs_res_pe<-obstab(subset(dat_a,TherapyCode=="PE"))
obs_res_relax<-obstab(subset(dat_a,TherapyCode=="Relax"))

dat_obs_ipt<-cbind(Symptom=obs_res_ipt[[2]]$Symptom,obs_res_ipt[[1]][,-10],obs_res_ipt[[2]][,-17])
dat_obs_pe<-cbind(Symptom=obs_res_pe[[2]]$Symptom,obs_res_pe[[1]][,-10],obs_res_pe[[2]][,-17])
dat_obs_relax<-cbind(Symptom=obs_res_relax[[2]]$Symptom,obs_res_relax[[1]][,-10],obs_res_relax[[2]][,-17])

pdf("C:/Users/Child Study/kd/20141125scid2/result/plot_orginal.pdf")
barp(base,trtr,0.6,"Dx at baseline","-Original data",35,10)
barp(lose,trtr,0.6,"Lose Dx at Week 14","-Original data",35,10)
barp(baselose,trtr,0.6,"Dx at baseline and Lose Dx at Week 14","-Original data",35,10)
barp(resbase,trtr,1,"Dx at baseline and week 14 by Response","-Original data",35,10)
barp(subset(tp1,variable%in%c("prop.res/base.no.v14","prop.res/base.dx.v14")&Symptom%in%dat_obs$Symptom[1:13]),trtr,1,
     "Lose Dx by Response","-Original data",35,10)
barp(subset(tp_trt,Symptom%in%dat_obs$Symptom[1:13]),trtr,1,
     "Lose Dx by Treatment","-Original data",35,10)
dev.off()

pdf("C:/Users/Child Study/kd/20141125scid2/result/plot_group.pdf")
barp(base_c,c("avoid_or_paran","other","anydx"),0.6,"Dx at baseline","-Group data",0,12)
barp(lose_c,c("avoid_or_paran","other","anydx"),0.6,"Lose Dx at Week 14","-Group data",0,12)
barp(baselose_c,c("avoid_or_paran","other","anydx"),0.6,"Dx at baseline and Lose Dx at Week 14","-Group data",0,12)
barp(resbase_c,c("avoid_or_paran","other","anydx"),1,"Dx at baseline and week 14 by Response","-Group data",0,12)
barp(subset(tp1,variable%in%c("prop.res/base.no.v14","prop.res/base.dx.v14")&Symptom%in%c("avoid_or_paran","other","anydx")),
     c("avoid_or_paran","other","anydx"),1,
     "Lose Dx by Response","-Group data",0,12)
barp(subset(tp_trt,Symptom%in%c("avoid_or_paran","other","anydx")),c("avoid_or_paran","other","anydx"),1,
     "Lose Dx by Treatment","-Group data",0,12)
dev.off()

pdf("C:/Users/Child Study/kd/20141125scid2/result/plot_cluster.pdf")
barp(base_cl,c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),0.6,"Dx at baseline","-Clusterd data",0,12)
barp(lose_cl,c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),0.6,"Lose Dx at Week 14","-Clusterd data",0,12)
barp(baselose_cl,c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),0.6,"Dx at baseline and Lose Dx at Week 14","-Clusterd data",0,12)
barp(resbase_cl,c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),1,"Dx at baseline and week 14 by Response","-Clusterd data",0,12)
barp(subset(tp1,variable%in%c("prop.res/base.no.v14","prop.res/base.dx.v14")&
              Symptom%in%c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd")),
     c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),1,
     "Lose Dx by Response","-Clusterd data",0,12)
barp(subset(tp_trt,Symptom%in%c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd")),
     c("par_schTy_schID","bor_nar_ant_his","avo_dep_ocd"),1,
     "Lose Dx by Treatment","-Clusterd data",0,12)
dev.off()

#colnames(dat_obs)
dat_obs<-cbind(Symptom=dat_obs[,10],dat_obs[,-10])
cldat<-function(dat,percent){
  for(i in percent){
    dat[,i]<-as.numeric(dat[,i])
    dat[,i]<-ifelse(is.na(dat[,i]),NA,paste(round(dat[,i],4)*100,"%",sep=""))
  }
  #dat[,rou]<-round(dat[,rou],4)
  result<-dat
  return(result)
}
dat_obs<-cldat(dat_obs,c(6:8))
dat_obs_res_woNa<-cbind(Symptom=dat_obs_res_woNa[,17],dat_obs_res_woNa[,-17])
dat_obs_res_woNa<-cldat(dat_obs_res_woNa,c(6:7,14,15,22,23))
dat_basebytrt<-cldat(dat_basebytrt,c(8:13))
dat_obs_ipt<-cldat(dat_obs_ipt,c(6:8,15,16,23,24,31,32))
dat_obs_pe<-cldat(dat_obs_pe,c(6:8,15,16,23,24,31,32))
dat_obs_relax<-cldat(dat_obs_relax,c(6:8,15,16,23,24,31,32))

dat_obs_gain<-cbind(overall[[3]]," "="",IPT=obs_res_ipt[[3]][,-1]," "="",PE=obs_res_pe[[3]][,-1],
                    " "="",Relax=obs_res_relax[[3]][,-1])
dat_obs_gain<-cldat(dat_obs_gain,c(4,8,12,16))
data_table<-list(dat_obs,dat_obs_res_woNa,dat_basebytrt,dat_obs_ipt,dat_obs_pe,dat_obs_relax,dat_obs_gain)
sheetname<-c("Freq and prop table(overall & by response)","freq,prop and test by response",
             "freq,prop and test by treatment","IPT group freq,prop by response","PE group freq,prop by response",
             "Relax group freq,prop by response","gain dx by response")
wb<-createWorkbook("xlsx")
for (i in 1:length(data_table)){
  s<-createSheet(wb,sheetName = sheetname[i])
  addDataFrame(data_table[[i]],s)
}
saveWorkbook(wb,"C:/Users/Child Study/kd/20141125scid2/result/table.xlsx")
#####################################
#MODELLING
dat_s<-dat_s[,c(1:3,grep("base",colnames(dat_s)),grep(".v14",colnames(dat_s)),grep("CAPS",colnames(dat_s)))]

dat_m<-dat_s[,c(-2,-17,-18,-21,-40,-41,-44)]
colnames(dat_m)[colnames(dat_m)=="CAPS_Response_V14"]<-"Response"
dat_m[,-c(1,2,18,38,43)]<-data.frame(sapply(dat_m[,-c(1,2,18,38,43)],function(x){
  ifelse(x==0,"No",ifelse(x==1,"Yes",x))
}))
dat_mb<-subset(dat_m,!is.na(avoid.base)&!is.na(Response))
dat_m<-subset(dat_m,!is.na(avoid.base)&!is.na(avoid.v14))
#dat_m<-cbind(dat_m[,1:15],dat_m[,c(30,31,34)],dat_m[,c(-1:-15,-29,-30,-31,-34)],"Response"=dat_m[,29])
mylgit<-function(var,dat,keep){
  dat[,var]<-factor(dat[,var])
  if(length(levels(dat[,var]))>1){
    l1<-glm(Response~dat[,var]+TherapyCode+TherapyCode:dat[,var],dat,family=binomial)
    l2<-glm(Response~dat[,var]+TherapyCode,dat,family=binomial)
    lt<-anova(l1,l2,test="F")
    t1<-cbind(" "=gsub("dat[, var]",paste(var,"=",sep=""),rownames(coef(summary(l1))),fixed = TRUE),
              as.data.frame(round(coef(summary(l1)),4),row.names=F)," "="")
    result<-as.data.frame(rbind(as.matrix(t1),c("Test","of","interaction","by","","F-test"),
                              colnames(as.data.frame(lt)),round(as.matrix(lt),4)),row.names=F)
    if(keep=="No"){
      if(as.matrix(lt)[2,6]>0.1){
      t2<-cbind(" "=gsub("dat[, var]",paste(var,"=",sep=""),rownames(coef(summary(l2))),fixed = TRUE),
                as.matrix(round(coef(summary(l2)),4),row.names=F)," "="")
      result<-as.matrix(result)
      result[nrow(result)-4,ncol(result)]<-round(summary(l1)$aic,2)
      t2[nrow(t2),ncol(t2)]<-round(summary(l2)$aic,2)
      row.names(t2)<-rep("",nrow(t2))
      result<-as.data.frame(rbind(as.matrix(result),c("","Model","without","Interaction","",""),t2),row.names=F)
      colnames(result)[ncol(result)]<-"AIC"
      }
    }
  }else if (length(levels(dat[,var]))==1){
    result<-data.frame(t(c(var,"has","only","1","level","")))
  }
return(result)
}
rs<-list()
for(i in colnames(dat_m)[c(-1:-2,-43)]){
  rs[[match(i,colnames(dat_m)[c(-1:-2,-43)])]]<-mylgit(i,dat_m,"Yes")
}

rs.f<-list()
for(i in 1:20){
  rs.f[[i]]<-data.frame(rbind(as.matrix(rs[[i]]),rep("",6),c(colnames(dat_m)[i+22],rep("",5)),as.matrix(rs[[i+20]])))
  colnames(rs.f[[i]])[1]<-colnames(dat_m)[i+2]
  colnames(rs.f[[i]])[ncol(rs.f[[i]])]<-" "

}
rsb<-list()
for(i in colnames(dat_mb)[c(-1:-2,-23:-43)]){
  rsb[[match(i,colnames(dat_mb)[c(-1:-2,-23:-43)])]]<-mylgit(i,dat_mb,"No")
}

sheetname_test<-substr(colnames(dat_m)[3:22],1,nchar(colnames(dat_m)[3:22])-5)
wb1<-createWorkbook("xlsx")
for (i in 1:length(rs.f)){
  s<-createSheet(wb1,sheetName = sheetname_test[i])
  addDataFrame(rs.f[[i]],s)
}
saveWorkbook(wb1,"C:/Users/Child Study/kd/20141125scid2/result/table_test.xlsx")

sheetname_test_base<-substr(colnames(dat_mb)[3:22],1,nchar(colnames(dat_mb)[3:22])-5)
wb3<-createWorkbook("xlsx")
for (i in 1:length(rsb)){
  s<-createSheet(wb3,sheetName = sheetname_test_base[i])
  addDataFrame(rsb[[i]],s)
}
saveWorkbook(wb3,"C:/Users/Child Study/kd/20141125scid2/result/table_test_base.xlsx")


#test<-melt(dat_m,id.vars=c("StudyID","TherapyCode","CAPS_Response_V14"))
#colnames(test)[3:4]<-c("Response","Symptom")
#ddply(test,.(TherapyCode,Symptom,value,Response),function(nrow)
group.dif.tx<-function(base,post){
group.dif<-function(model,linkm){
  log.or<-round(confint(glht(model, linfct = linkm))$confint[1,1],4)
  or<-round(exp(log.or),4)
  conin<-paste("(",round(exp(confint(glht(model, linfct = linkm))$confint[1,2]),2),",",
               round(exp(confint(glht(model, linfct = linkm))$confint[1,3]),2),")")
  stat<-round(summary(glht(model, linfct = linkm))$test$tstat[[1]],4)
  pval<-round(summary(glht(model, linfct = linkm))$test$pvalues[1],4)
  
  result<-t(c(log.or,or,conin,stat,pval))
  return(result)
}

avp.b.g<-rbind(cbind("IPT-PE",group.dif(base,matrix(c(0,0,-1,0,-1,0),1))),
               cbind("IPT-Relax",group.dif(base,matrix(c(0,0,0,-1,0,-1),1))),
               cbind("PE-Relax",group.dif(base,matrix(c(0,0,1,-1,1,-1),1))))
avp.v14.g<-rbind(cbind("IPT-PE",group.dif(post,matrix(c(0,0,-1,0,-1,0),1))),
                 cbind("IPT-Relax",group.dif(post,matrix(c(0,0,0,-1,0,-1),1))),
                 cbind("PE-Relax",group.dif(post,matrix(c(0,0,1,-1,1,-1),1))))

avp.g<-rbind("",avp.b.g,"",avp.v14.g)
colnames(avp.g)<-c("Contrast","Log OR","Odds Ratio","95% CI","Z-Statistic","P_value")
return(avp.g)
}

l_avo_par<-glm(Response~avoid_or_paran.base+TherapyCode+TherapyCode:avoid_or_paran.base,dat_m,family=binomial)
l_avo_par.v14<-glm(Response~avoid_or_paran.v14+TherapyCode+TherapyCode:avoid_or_paran.v14,dat_m,family=binomial)

l_any<-glm(Response~anydx.base+TherapyCode+TherapyCode:anydx.base,dat_m,family=binomial)
l_any.v14<-glm(Response~anydx.v14+TherapyCode+TherapyCode:anydx.v14,dat_m,family=binomial)

l_pss<-glm(Response~par_schTy_schID.base+TherapyCode+TherapyCode:par_schTy_schID.base,dat_m,family=binomial)
l_pss.v14<-glm(Response~par_schTy_schID.v14+TherapyCode+TherapyCode:par_schTy_schID.v14,dat_m,family=binomial)

l_bnah<-glm(Response~bor_nar_ant_his.base+TherapyCode+TherapyCode:bor_nar_ant_his.base,dat_m,family=binomial)
l_bnah.v14<-glm(Response~bor_nar_ant_his.v14+TherapyCode+TherapyCode:bor_nar_ant_his.v14,dat_m,family=binomial)

l_ado<-glm(Response~avo_dep_ocd.base+TherapyCode+TherapyCode:avo_dep_ocd.base,dat_m,family=binomial)
l_ado.v14<-glm(Response~avo_dep_ocd.v14+TherapyCode+TherapyCode:avo_dep_ocd.v14,dat_m,family=binomial)

avp.g<-group.dif.tx(l_avo_par,l_avo_par.v14)
avp.g[c(1,5),1]<-c("Avoid_or_Paran.base","Avoid_or_Paran.v14")
avp.g<-data.frame(avp.g)

any.g<-group.dif.tx(l_any,l_any.v14)
any.g[c(1,5),1]<-c("Anydx.base","Anydx.v14")
any.g<-data.frame(any.g)

pss.g<-group.dif.tx(l_pss,l_pss.v14)
pss.g[c(1,5),1]<-c("par_schTy_schID.base","par_schTy_schID.v14")
pss.g<-data.frame(pss.g)

bnah.g<-group.dif.tx(l_bnah,l_bnah.v14)
bnah.g[c(1,5),1]<-c("bor_nar_ant_his.base","bor_nar_ant_his.v14")
bnah.g<-data.frame(bnah.g)

ado.g<-group.dif.tx(l_ado,l_ado.v14)
ado.g[c(1,5),1]<-c("avo_dep_ocd.base","avo_dep_ocd.v14")
ado.g<-data.frame(ado.g)

rs.g<-list(avp.g,any.g,pss.g,bnah.g,ado.g)
sheetname_test_g<-c("Avoid_or_Paran","Anydx","par_schTy_schID","bor_nar_ant_his.base","avo_dep_ocd.base")
wb2<-createWorkbook("xlsx")
for (i in 1:length(rs.g)){
  s<-createSheet(wb2,sheetName = sheetname_test_g[i])
  addDataFrame(rs.g[[i]],s)
}

saveWorkbook(wb2,"C:/Users/Child Study/kd/20141125scid2/result/table_test_group.xlsx")
orTx<-function(symp){
 odds.dat<-subset(dat_m,symp=="Yes")
 odds.table<-table(odds.dat$TherapyCode,odds.dat$Response)[,c(2,1)]
 rest1<-oddsratio.fisher(odds.table[1:2,])
 rest2<-oddsratio.fisher(odds.table[-2,])
 rest3<-oddsratio.fisher(odds.table[-1,])
 result<-cbind(odds.table,c("IPT vs PE","PE vs Relax","IPT vs Relax"),
                  round(rbind(rest1[[2]][2,],rest3[[2]][2,],rest2[[2]][2,]),4),
               round(c(rest1[[3]][2,2],rest3[[3]][2,2],rest2[[3]][2,2]),4))
 return(result)
}
apor.base<-orTx(dat_m$avoid_or_paran.base)
apor.v14<-orTx(dat_m$avoid_or_paran.v14)

apor<-data.frame(rbind(cbind(c("IPT","PE","Relax"),apor.base),c("v14",rep("",7)),cbind(c("IPT","PE","Relax"),apor.v14)))
colnames(apor)[c(1:4,ncol(apor))]<-c("base","Response:Yes","Response:No","contrast","P-value")

orDx<-function(trt,symp){
  odds.dat<-subset(dat_m,TherapyCode==trt)
  odds.table<-table(odds.dat[,symp],odds.dat$Response)[c(2,1),c(2,1)]
  result<-cbind(c("Dx","no Dx"),odds.table,rbind(c(trt,"","",""),c(round(oddsratio.fisher(odds.table)[[2]][2,],4),
                                                                   round(oddsratio.fisher(odds.table)[[3]][2,2],4))))
  colnames(result)<-c(symp,"Response:Yes","Response:No",colnames(oddsratio.fisher(odds.table)[[2]]),"P_value")
  return(result)
}

avod.dx<-data.frame(rbind(orDx("IPT","avoid_or_paran.base"),"",orDx("PE","avoid_or_paran.base"),"",
                             orDx("Relax","avoid_or_paran.base"),""))
avod.v14<-rbind(orDx("IPT","avoid_or_paran.v14"),"",orDx("PE","avoid_or_paran.v14"),"",
                orDx("Relax","avoid_or_paran.v14"),"")
avod.v14<-data.frame(rbind(colnames(avod.v14),avod.v14))
colnames(avod.v14)<-colnames(avod.dx)
avod.or<-rbind(avod.dx,"","",avod.v14)

odds.dat<-subset(dat_m,dat_m$antisoc.base=="Yes")
anti<-data.frame(table(odds.dat$TherapyCode,factor(odds.dat$Response,levels=c(0,1)))[,c(2,1)])
colnames(anti)<-c("Response:Yes","No")

rs.or<-list(apor,avod.or,anti)
sheetname_test_or<-c("avo_par_byTx","avo_par_byDx","antisocial")
wb4<-createWorkbook("xlsx")
for (i in 1:length(rs.or)){
  s<-createSheet(wb4,sheetName = sheetname_test_or[i])
  addDataFrame(rs.or[[i]],s)
}

saveWorkbook(wb4,"C:/Users/Child Study/kd/20141125scid2/result/group_or.xlsx")