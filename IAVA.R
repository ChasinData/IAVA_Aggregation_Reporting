
# remove all R objects in the working directory
rm(list=ls(all=TRUE))

#Increase memory size (java out of memory and must be before pkg load)
options(java.parameters = "-Xmx15000m")

# id packages
packages <- c("plyr","dplyr","readxl", "zoo",  "stringi", "ggplot2","tidyr","stringr")

#install and library
lapply(packages, require, character.only = TRUE)

#set folders
proj_name = "IAVA Scorecard"
base_path="C:/Users/Stephen.M.Trask/Documents"
severity_data="S:/CISO/ACAS/ACAS IAVA Reports/IAVA Severity Summary Reports CSV Format"
summary_data="S:/CISO/ACAS/ACAS IAVA Reports/IAVA Summary Reports CSV Format"
archive_data="S:/CISO/ACAS/ACAS IAVA Reports/IAVA SCORECARD ARCHIVE DATA"
report="S:/CISO/ACAS/ACAS IAVA Reports"
code.plc='H:/Code Source/IAVA_Aggregation_Reporting'

#set project path
proj_path = paste(base_path,"/",proj_name,sep = "")
out_path = file.path(proj_path, "Output")
cleaned_path = file.path(proj_path, "CleanData")
data_path = file.path(proj_path, "RawData")
Outlier_Path  = file.path(proj_path, "Outliers")
Asset_path="S:/CISO/ACAS/ACAS IAVA Reports/Current Asset Numbers/Total Assets by HRP.xlsx"

dir.create(file.path(proj_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(out_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(cleaned_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(data_path), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(Outlier_Path), showWarnings = FALSE, recursive = FALSE, mode = "0777")




#Get Date
age= 100
date <- Sys.Date()

#Get File names
severity.files <- list.files(severity_data, pattern="*.csv", recursive = F, full.names=F, all.files = T, include.dirs = T)
summary.files <- list.files(summary_data, pattern="*.csv", recursive = F, full.names=F, all.files = T, include.dirs = T)
#asset.file <- list.files(Asset_path, pattern="*.xlsx", recursive = F, full.names=F, all.files = T, include.dirs = T)
#EMPTY DATAFRAMES
severity=as.data.frame(matrix( ,0,3))
summary=as.data.frame(matrix( ,0,4))
#nvl=as.data.frame(matrix( ,0,2))
#no.Vulnerabilities =matrix( ,0,2)
no.Vulnerabilities.lst=as.list(NULL)

#BIND ALL FILES

for (i in  severity.files){
  #i=severity.files[1]
  df<- read.csv(file.path(severity_data, i), header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  #df<- read.csv(i, header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  df$HRP.name <- sub(" IAVA Severity for Alerts.csv","",i)
  #df$HRP.name <- sub(" Severity Summary for Alerts.csv","",i)
  df$As.of.Date=file.mtime(file.path(severity_data, i)); df$As.of.Date=(format(df$As.of.Date,"%d-%b-%Y"))
  severity=as.data.frame(rbind(severity,df))
}
severity$HRP.name=gsub("Regional Health Command - Atlantic", "RHC-A (P) HQ", severity$HRP.name, ignore.case = FALSE, perl = FALSE,
                       fixed = FALSE, useBytes = FALSE) 
#severity$HRP.name=gsub("RHC -A", "RHC-A (P) HQ", severity$HRP.name, ignore.case = FALSE, perl = FALSE,
#                       fixed = FALSE, useBytes = FALSE) 
severity$HRP.name=gsub("Ft. Myers Severity Summary for Alerts.csv", "Ft. Myers", severity$HRP.name, ignore.case = FALSE, perl = FALSE,
                       fixed = FALSE, useBytes = FALSE) 

severity$As.of.Date=as.POSIXct(severity$As.of.Date, format = "%d-%b-%Y")

sev.max.date=max(severity$As.of.Date)

for (i in  summary.files){
  #i=summary.files[1]
  df<- read.csv(file.path(summary_data, i), header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  #df<- read.csv(i, header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  if(nrow(df)>0) {
    df$HRP.name <- sub(" IAVA Summary for Alerts.csv","",i)
    df$As.of.Date=file.mtime(file.path(summary_data, i))
    df$As.of.Date=(format(df$As.of.Date,"%d-%b-%Y"))
    df$IAVM.Count=length(unique(df$IAVM))
  } else {
    no.Vulnerabilities.lst= c(sub(" IAVA Summary for Alerts.csv","",i), no.Vulnerabilities.lst)
    print("no data")
  }
  summary=rbind(summary,df)
  #no.Vulnerabilities =rbind(nvl,no.Vulnerabilities )
}
summary$HRP.name=gsub("Regional Health Command - Atlantic", "RHC-A (P) HQ", summary$HRP.name, ignore.case = FALSE, perl = FALSE,
                      fixed = FALSE, useBytes = FALSE) 
#summary$HRP.name=gsub("RHC -A (P) HQ", "RHC-A (P) HQ", summary$HRP.name, ignore.case = FALSE, perl = FALSE,
#                      fixed = FALSE, useBytes = FALSE) 
summary$HRP.name=gsub("Ft. Myers Severity Summary for Alerts.csv", "Ft. Myers", summary$HRP.name, ignore.case = FALSE, perl = FALSE,
                      fixed = FALSE, useBytes = FALSE) 
summary$As.of.Date=as.POSIXct(summary$As.of.Date, format = "%d-%b-%Y")
sum.max.date=max(summary$As.of.Date)


#CREATE BACKUPS
write.csv(severity, file.path(archive_data,paste("Severity Data as of ", date, ".csv",sep="")))
write.csv(summary, file.path(archive_data,paste("Summary Data as of ", date, ".csv",sep="")))
nam=paste("HRPs with NO Vulnerabilities greater 21 days as of ", date,".txt",sep="")
dput(no.Vulnerabilities.lst, file = file.path(archive_data,"Current HRP with No Vulnerabilities post 21 days.txt"))


save(summary, file = file.path(report,"summary.rds"))
save(severity, file = file.path(report,"severity.rds"))

summary.bu=summary
severity.bu=severity

#summary=summary[1:1505, ]
#summary=summary.bu
#severity=severity.bu

#Summary Charts

#Collect Data for Calcs
######ADD REGION ROLL UP
Site<-c("GAHC","IRACH","KACC","KACH","KAHC","MCAHC","RHC-A (P) HQ","WAMC","BACH","EAMC","FAHC","LAHC","MFACH","WINN","DAHC","BMACH","ARHC","FTDTL","KUSAHC","ALL")
HRP.name<-c("Ft. Drum","Ft. Knox","Ft. Meade","West Point","Ft. Lee","Ft. Eustis","RHC -A (P) HQ","Ft. Bragg","Ft. Campbell","Ft. Gordon","Redstone","Ft. Rucker","Ft. Jackson","Ft. Stewart","Carlisle Barracks","Ft. Benning","Ft. Myers","FTDTL","Kirk APG","Region (ALL HRPs)")
Facility<-c("CLINIC","HOSP","CLINIC","HOSP","CLINIC","CLINIC","HQ","MEDCEN","HOSP","MEDCEN","CLINIC","CLINIC","HOSP","HOSP","","HOSP","","","","ALL")
UIC<-c("W4U2AA","W2LAAA","W6F2AA","W2H8AA","W2LMAA","W2K1AA","W07TAA","W2L6AA","W2L8AA","W3QMAA","W2FLAA","W2MQAA","W2MJAA","W2MSAA","","W2L3AA","","","","ROLL UP")
Short<-c("DRUM","KNOX","MEADE","WEST POINT","LEE","EUSTIS","RHC-A","BRAGG","CAMPBELL","GORDON","REDSTONE","RUCKER" ,"JACKSON","STEWART","CARLISLE","BENNING","MYERS","FTDTL","KIRK","REGION")
SiteCode<-c("N03","N04","N05","N13","N12","N11","N18","N10","S05","S03","S07","S08","S06","S09","N14", "S04", "N05 / N18", "N17","N06","NN")
Rollup <- as.data.frame(cbind(Site,HRP.name,Facility,UIC,Short,SiteCode))


##Asset data 
assets<-read_excel(Asset_path, sheet = 1, col_names = TRUE, col_types = NULL, na = "",skip = 0)
assets=assets[ ,1:2] #Ensure no extraneuos data
colnames(assets)[1]='ComputerName';colnames(assets)[2]='SiteCode'
assets=assets[!is.na(assets["ComputerName"]), ] #Complete Cases only
assets$SiteCode=as.factor(assets$SiteCode)

#######################
Asset.Data = select(assets, everything() ) %>%
  group_by(SiteCode) %>%
  summarise(Assets=length(SiteCode))
Asset.Total = select(assets, everything() ) %>%
  summarise(Assets=length(SiteCode))
Asset.Total$SiteCode = "NN"; Asset.Total=Asset.Total[ ,c(2,1)]
Assets.Complete=rbind(Asset.Total,Asset.Data)

Rollup=join(Rollup, Assets.Complete, by = "SiteCode")

#Total assets for region
Region=length(assets$SiteCode); 
#total iava for region
IAVA=sum(summary$Total)
#unique IAVM
Region.IAVM.Count=length(unique(summary$IAVM))
#######################################################
#IAVM data
#load(file = file.path(archive_data,"summary.rds"))
IAVM<-summary#.bu
IAVM=IAVM[IAVM$IAVM != "NotNeeded", ]
IAVM$Year = sapply(strsplit(as.character(IAVM$IAVM), "\\-"), "[", 1)
IAVM.Year = select(IAVM, -Host.Total, -As.of.Date) %>%
  group_by(HRP.name,Year,Severity) %>%
  summarise(IAVM.Count=length(unique(IAVM)))

IAVM.cht <- ggplot(IAVM.Year, aes(x = Severity, y = IAVM.Count, fill = HRP.name  )) +
  geom_bar(stat='identity') +
  #geom_text(aes(label = IAVM.Count, y = IAVM.Count*1.05), size = 3) +
  facet_grid(~ Year ) + #scales = 'free'
  scale_colour_hue(c=4, l = 80) + #higher l is darker
  #scale_y_continuous(labels = comma) +
  xlab('') +
  ylab('Amount of IAVMs') +
  ggtitle(paste('All Open IAVMs in Region as of ', sum.max.date, sep="")) +
  theme(legend.position = "bottom", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottomt 
ggsave(file.path(proj_path,"Open IAVM.png"),width=10, height=7)


#################ADD REGION DATA to summary

Region.Data = select(summary, -Host.Total, -IAVM,  -As.of.Date) %>%
  group_by(Severity) %>%
  summarise(Total=sum(Total))
Region.Data$As.of.Date=sum.max.date
Region.Data$HRP.name="Region (ALL HRPs)"
Region.Data$IAVM="NotNeeded"
Region.Data$Host.Total=Region.Data$Total
Region.Data=Region.Data[ ,c(5,1,6,2,4,3)]
Region.Data$IAVM.Count=Region.IAVM.Count
summary=rbind(summary,Region.Data)
chk=unique(summary$HRP.name)
#####

#summary=summary[-c(1575:1577), ]

Sum.Data=select(summary, -IAVM, -Total,-As.of.Date) %>%
  group_by(HRP.name) %>%
  summarise(Total=sum(Host.Total)) %>%
  arrange(-Total)

#Order data for charts
x=Sum.Data
x$HRP.name <- factor(x$HRP.name, levels = x$HRP.name[order(x$Total)])
Sum.Data=x

#Score calc
Sum.Data <- join(Sum.Data, Rollup, by = "HRP.name")
Sum.Data$Score <- as.numeric(format(round(Sum.Data$Total/Sum.Data$Assets,2)))
Sum.Data$Score.Name <-paste(Sum.Data$Short,Sum.Data$Score,sep = "_")

meansub=Sum.Data[ ! Sum.Data$Site %in% c("ALL"), ] 
meancalc=format(round(mean(meansub$Score),2))


Sum.Data.bySeverity=select(summary, -IAVM, -Total,-As.of.Date) %>%
  group_by(HRP.name,Severity) %>%
  summarise(Total=sum(Host.Total)) %>%
  arrange(-Total)

#Cast DF to do rowsums
Sum.Data.bySeverity=spread(Sum.Data.bySeverity,Severity,Total,fill=0)
slide.sum = Sum.Data.bySeverity
assets4merge=Rollup[ ,c(2,7)]
slide.sum = merge(assets4merge,slide.sum) 
#Examine Levels
Sum.Data.bySeverity$HRP.name=as.factor(Sum.Data.bySeverity$HRP.name);
levels(Sum.Data.bySeverity$HRP.name)
#Reorder the HRP Column by  sum of Vulnerabilities  
Sum.Data.bySeverity$HRP.name<-reorder(Sum.Data.bySeverity$HRP.name, rowSums(Sum.Data.bySeverity[-1]))
#Pull it back together
Sum.Data.bySeverity=gather(Sum.Data.bySeverity,key=Severity,value=Total, -HRP.name)

####BRing in score data
temp=Sum.Data[ ,c(1,6:9)]
Sum.Data.bySeverity <- merge(Sum.Data.bySeverity, temp, by.all = "HRP.name")

severity = filter(severity, severity$Severity == "High" |
                    severity$Severity == "Critical")

##################### Get severity data
Sev.Data=select(severity, -As.of.Date) %>%
  group_by(HRP.name) %>%
  summarise(Total=sum(Count))%>%
  arrange(-Total)
Sev.Data$Date=sev.max.date

load(file = file.path(report,"Historical Data BACKUP.RDS"))
#load(file = file.path(archive_data,"Historical Data MOST RECENT.RDS"))#Historical Data MOST RECENT.RDS
#colnames(bck)[3]="HRP.name"
###Remove last merge
dt=stri_sub(as.character(sev.max.date), 1, 10)
#bck$Date=as.character(bck$As.of.Date)
bck=filter(bck, Date!=dt); bck=bck[ ,c(3,1,2)]


#bck$Date=as.POSIXct(bck$Date, format="%d-%b-%Y")
#### SHOW CODE WHEN NOW Data avail
bck=rbind(Sev.Data[ ,1:3],bck)
bck.mr=bck


save(bck.mr,file=file.path(report,"Historical Data MOST RECENT.RDS"))
save(bck,file=file.path(report,"Historical Data BACKUP.RDS"))

#Graph Visuals
Sum.Data$Score.Name<-reorder(Sum.Data$Score.Name, Sum.Data$Total)
tt=Sum.Data[-1, ]
sum(tt$Total)
#Reorder the HRP Column by  sum of Vulnerabilities  
sum.cht <- ggplot(Sum.Data, aes(x = Score.Name, y = Total )) +
  geom_bar(aes(fill = HRP.name), stat="identity") +
  geom_text(aes(label = format(round(Total, 2)), y = Total+250), size = 3) +
  xlab('') +
  ylab('Total Vulnerabilities ') +
  ggtitle(paste('Total Vulnerabilities, by HRP as of ', sum.max.date, sep="")) +
  theme(legend.position = "none", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottom
sum.cht
ggsave(file.path(proj_path,"Total Assets with Vulnerabilities, by HRP.png"),width=10, height=7)
dev.off()

Sum.Data$Score <- as.numeric(format(round(Sum.Data$Total/Sum.Data$Assets,2)))
Sum.Data$HRP.name<-reorder(Sum.Data$HRP.name, Sum.Data$Score)
write.csv(Sum.Data, file.path(archive_data,paste("Summary Charting Data on ", date, ".csv",sep="")))
score.cht <- ggplot(Sum.Data, aes(x = HRP.name, y = Score )) +
  geom_bar(aes(fill = HRP.name), stat="identity") +
  geom_text(aes(label = Score, y = Score +.08), size = 3) +
  xlab('') +
  geom_hline(aes(yintercept = as.numeric(format(round(mean(meansub$Score),2))))) +
  annotate("text", min(Sum.Data$Score) + 2, as.numeric(meancalc) + .51, label = paste("Avg. Score is ",meancalc,sep = "")) +
  #geom_hline(data = Sum.Data,aes(yintercept=meanc)) +
  ylab('Score (Vulnerabilities/Assets)') +
  ggtitle(paste('Score by HRP as of ', sum.max.date, sep="")) +
  theme(legend.position = "none", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottom
score.cht
ggsave(file.path(proj_path,"Score, by HRP.png"),width=10, height=7)
dev.off()

Sum.Data.bySeverity$Severity <- factor(Sum.Data.bySeverity$Severity, levels = c("Critical","High","Medium"))
library(scales)

SevbyHRP.cht <- ggplot(Sum.Data.bySeverity, aes(x = HRP.name, y = Total, fill = factor(Severity) )) +
  geom_bar(stat="identity") +
  scale_fill_manual(values=c("red", "orange", "yellow")) +
  #geom_text(aes(label = format(round(Total, 2)), y = Total+250), size = 3) +
  xlab('') +
  ylab('Total Vulnerabilities (000s)') +
  scale_y_continuous(labels = comma) +
  #coord_flip() +
  ggtitle(paste('Total Vulnerabilities, by HRP  as of ', sum.max.date, sep="")) +
  theme(
    legend.title=element_blank(),
    legend.position=c(.25,.7),
    axis.text.x=element_text(angle=20, vjust = 1,hjust=1),
    axis.text.y=element_text(angle=20, vjust = 1,hjust=1)
  )
#theme(legend.position = "bottom", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottom
SevbyHRP.cht
ggsave(file.path(proj_path,"Vulnerabilities by Severity for each HRP.png"),width=10, height=7)
dev.off()

tmp2=Sum.Data[ ,c(1,10)];Sum.Data.bySeverity=join(Sum.Data.bySeverity, tmp2, by = "HRP.name")

write.csv(Sum.Data.bySeverity, file.path(archive_data,paste("Severity Charting Data on ", date, ".csv",sep="")))

wrap.cht <- ggplot(Sum.Data.bySeverity, aes(x = Severity, y = Total )) +
  geom_bar(aes(fill = Severity), stat="identity") +
  facet_wrap(~Score.Name, scales = "free") +
  scale_fill_manual(values=c("red", "orange", "yellow")) +
  geom_text(aes(label = format(round(Total, 2)), y = Total*1.2), size = 3) +
  xlab('') +
  ylab('Total Vulnerabilities ') +
  ggtitle('Total Vulnerabilities by HRP') +
  theme(legend.position = "none", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottom
wrap.cht
ggsave(file.path(proj_path,"Total by Severity and HRP.png"),width=10, height=7)
dev.off()


bck.d <- join(bck, temp, by = "HRP.name")
#Get region data
historical = select(bck.d, everything()) %>%
  group_by(Date) %>%
  summarise(Total=sum(Total))
historical$HRP.name="Region (ALL HRPs)"
historical=historical[ ,c(3,2,1)]
historical$Short = "REGION"
bck.2=bck.d[ ,1:4]
historical=rbind(historical,bck.2)


#bck.d=na.omit(bck.d)
#Order Dates
historical$Date=as.Date(historical$Date, "%d-%b-%Y")

#When is last time file updated
historical$Age=difftime(Sys.Date(), historical$Date, units = "days") 

#Filter for recent files
historical=historical[historical$Age < age, ]
write.csv(historical, file.path(archive_data,paste("Historical Charting IAVA beyond 21 days as of ", date, ".csv",sep="")))
historical2<-filter(historical,historical$HRP.name !="Region (ALL HRPs)")

history <- ggplot(historical2, aes(x = Date, y = Total, fill=HRP.name )) +
  #geom_line(size = 1) +
  geom_smooth() +
  facet_wrap(~HRP.name, scales = "free") +
  #geom_text(aes(label = HRP.name, x=max(Date),y = Total+250), size = 3) +
  # geom_text(aes(label = Total, x=Date,y = Total*1.1), size = 3) +
  scale_colour_hue(c=45, l = 45) + #higher l is darker
  xlab('') +
  ylab('Total Critical and High Vulnerabilities ') +
  ggtitle(paste(age, " Day ",'History of Vulnerabilities (Critical and High) beyond 21 Days, as of ', sev.max.date, sep="")) +
  theme(legend.position = "none", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottomt 
history
ggsave(file.path(proj_path,"Total Critical and High Severity and HRP.png"),width=10, height=7)
dev.off()

historical3<-filter(historical,historical$HRP.name =="Region (ALL HRPs)")

history <- ggplot(historical3, aes(x = Date, y = Total, colour=HRP.name )) +
  #geom_line(size = 1) +
  geom_smooth(size = 1) +
  #facet_wrap(~HRP.name ) +
  geom_text(aes(label = Total, x=Date,y = Total*1.1), size = 3) +
  scale_colour_hue(c=45, l = 45) + #higher l is darker
  xlab('') +
  ylab('Total Critical and High Vulnerabilities ') +
  ggtitle(paste(age, " Day ",'History of Vulnerabilities (Critical and High) beyond 21 Days for RHC-A Region, as of ', sev.max.date, sep="")) +
  theme(legend.position = "none", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottomt 
history
ggsave(file.path(proj_path,"Total Critical and High Severity in RHC-A.png"),width=10, height=7)
dev.off()




code.plc='H:/Code Source/IAVA_Aggregation_Reporting'
save(bck,file=file.path(report,"Historical Data BACKUP.RDS"))
save(bck,file=file.path(report,"Historical Data MOST RECENT.RDS"))

#unique(historical$Date)

##Add region over time chart

####################################################
#Get File names
summary.files2 <- list.files(summary_data, pattern="*.csv", recursive = T, full.names=T, all.files = T, include.dirs = T)
#EMPTY DATAFRAMES
summary2=as.data.frame(matrix( ,0,4))
#nvl=as.data.frame(matrix( ,0,2))
#no.Vulnerabilities =matrix( ,0,2)
no.Vulnerabilities.lst2=as.list(NULL)

#BIND ALL FILES
for (i in  summary.files2){
  #i=summary.files[1]
  df<- read.csv( i, header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  if(nrow(df)>0) {
    #df$HRP.name <- sub(" IAVA Summary for Alerts.csv","",i)
    #df$HRP.name <- sub("S:/CISO/ACAS/ACAS IAVA Reports","",i)#S:/CISO/ACAS/ACAS IAVA Reports
    df$As.of.Date=file.mtime(i)
    df$As.of.Date=(format(df$As.of.Date,"%d-%b-%Y"))
    df$IAVM.Count=length(unique(df$IAVM))
  } else {
    no.Vulnerabilities.lst= c(sub(" IAVA Summary for Alerts.csv","",i), no.Vulnerabilities.lst)
    print("no data")
  }
  summary2=rbind(summary2,df)
  #no.Vulnerabilities =rbind(nvl,no.Vulnerabilities )
}
summary2$As.of.Date=as.POSIXct(summary2$As.of.Date, format = "%d-%b-%Y")
sum.min.date2=min(summary2$As.of.Date)
# returns string w/o leading whitespace
trim.leading <- function (x)  sub("^\\s+", "", x)

# returns string w/o trailing whitespace
trim.trailing <- function (x) sub("\\s+$", "", x)

# returns string w/o leading or trailing whitespace
trim <- function (x) gsub("^\\s+|\\s+$", "", x)

summary2$IAVM = trim(summary2$IAVM)
First.DF = filter(summary2,summary2$As.of.Date==sum.min.date2 )#sum.min.date
First.DF$IAVM=trim(First.DF$IAVM)

First.IAVM=as.list(unique(First.DF$IAVM))
Current.IAVM=as.list(unique(summary$IAVM))

#Subset DF based on  LIST
Difference.IAVM= summary2[is.na(match(summary2$IAVM, First.DF$IAVM)), ] 

Week.Calc=select(Difference.IAVM, everything()) %>%
  group_by(Severity,IAVM) %>%
  summarise(days=difftime(max(As.of.Date),min(As.of.Date),units = "days"))

Avg.by.Sev=select(Week.Calc, everything()) %>%
  group_by(Severity) %>%
  summarise(Avg.Days=as.numeric(round(mean(days),2)))

Days.cht <- ggplot(Avg.by.Sev, aes(x = Severity, y = Avg.Days, fill = Severity  )) +
  geom_bar(stat='identity') +
  geom_text(aes(label = Avg.Days, y = Avg.Days*1.05), size = 3) +
  #geom_text(aes(label = "Calculated by averaging time for all IAVMs that emerged during 2016", y = 2), size = 3) +
  #facet_grid(~ Year ) + #scales = 'free'
  scale_colour_hue(c=4, l = 80) + #higher l is darker
  #scale_y_continuous(labels = comma) +
  xlab('') +
  ylab('Average Days to Removal of Vulnerability') +
  ggtitle(paste('Average Time to Remove Vulnerabilies in 2016', sep="")) +
  theme(legend.position = "bottom", axis.text.x=element_text(angle=20, vjust = 1,hjust=1)) #legend could be bottomt 
ggsave(file.path(proj_path,"Vulnerability Eradication.png"),width=10, height=7)


save.image(file.path(report,"IAVA_WORKSPACE.RData"))


###########################
library(ReporteRs)
#set global font size
options( "ReporteRs-fontsize" = 11.5)
mydoc = pptx(template = file.path(archive_data, "IAVA Scorecard Template2.pptx"))

MyFTable = vanilla.table( data = slide.sum )
MyFTable = setZebraStyle( MyFTable, odd = '#eeeeee', even = 'white' )
###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'Summary and Ratio Data')

body4 = Region #Total ASSETS
body3 = sum(summary$Total)/2 #Total IAVA
body2 = format(round(body3/body4,2))
body1 = sum.max.date
body5 = body1

mydoc = addSlide( mydoc, "IAVA by HRP Summary" )
mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = as.character(body1) ) #1
mydoc = addParagraph( mydoc, value = as.character(body1) )#2
mydoc = addFlexTable(mydoc,  MyFTable)
#mydoc = addSlide( mydoc, "Status" )
#writeDoc( mydoc, "test.pptx")

###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'Summary and Ratio Data')

body4 = Region #Total ASSETS
body3 = sum(summary$Total)/2 #Total IAVA
body2 = format(round(body3/body4,2))
body1 = as.character(sum.max.date)
body5 = body1

mydoc = addSlide( mydoc, "Summary and Ratio Data" )
#mydoc = addTitle( mydoc, as.character(x$Functional.Area))
mydoc = addParagraph( mydoc, value = body1 ) #1
mydoc = addParagraph( mydoc, value = body2 )#2
mydoc = addParagraph( mydoc, value = as.character(body3) )#3
mydoc = addParagraph( mydoc, value = as.character(body1)) #4
mydoc = addImage(mydoc, file.path(proj_path, "Score, by HRP.png"))
mydoc = addParagraph( mydoc, value = as.character(body4)) #6
mydoc = addParagraph( mydoc, value = as.character(Region.IAVM.Count)) #7
#mydoc = addSlide( mydoc, "Status" )
# writeDoc( mydoc, "test.pptx")
#########################################################################



# mydoc = pptx(template = file.path(archive_data, "IAVA Scorecard Template2.pptx"))

###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'IAVA by HRP Summary')

body4 = Region #Total ASSETS
body3 = sum(summary$Total)/2 #Total IAVA
body2 = format(round(body3/body4,2))
body1 = as.character(sum.max.date)
body5 = body1

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1

mydoc = addImage(mydoc, file.path(proj_path, "Vulnerabilities by Severity for each HRP.png"))
#mydoc = addSlide( mydoc, "Status" )
# writeDoc( mydoc, "test.pptx")

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1

mydoc = addImage(mydoc, file.path(proj_path, "Open IAVM.png"))
#mydoc = addSlide( mydoc, "Status" )
# writeDoc( mydoc, "test.pptx")
###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'Chart_Only')

body4 = Region #Total ASSETS
body3 = sum(summary$Total) #Total IAVA
body2 = format(round(body3/body4,2))
body1 = as.character(sum.max.date)
body5 = body1

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1
#  mydoc = addParagraph( mydoc, value = body1 )#2
mydoc = addImage(mydoc, file.path(proj_path, "Total by Severity and HRP.png"))
#mydoc = addSlide( mydoc, "Status" )
# writeDoc( mydoc, "test.pptx")


# mydoc = pptx(template = file.path(archive_data, "IAVA Scorecard Template2.pptx"))
###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'Chart_Only')

body4 = Region #Total ASSETS
body3 = sum(summary$Total) #Total IAVA
body2 = format(round(body3/body4,2))
body1 = as.character(sum.max.date)
body5 = body1

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1
#  mydoc = addParagraph( mydoc, value = body1 )#2
mydoc = addImage(mydoc, file.path(proj_path, "Vulnerability Eradication.png"))
#mydoc = addSlide( mydoc, "Status" )
# writeDoc( mydoc, "test.pptx")


###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, '21 Day data')

body4 = Region #Total ASSETS
body3 = sum(summary$Total)/2 #Total IAVA
body2 = format(round(body3/body4,2))
body5 = body1

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1
#  mydoc = addParagraph( mydoc, value = body1 )#2
mydoc = addImage(mydoc, file.path(proj_path, "Total Critical and High Severity and HRP.png"))
#mydoc = addSlide( mydoc, "Status" )

###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, '21 Day data')

body4 = Region #Total ASSETS
body3 = sum(summary$Total)/2 #Total IAVA
body2 = format(round(body3/body4,2))
body5 = body1

mydoc = addSlide( mydoc, "Chart_Only" )
#  mydoc = addTitle( mydoc, "G-6: IAVA Compliance Status")
mydoc = addParagraph( mydoc, value = body1 ) #1
#  mydoc = addParagraph( mydoc, value = body1 )#2
mydoc = addImage(mydoc, file.path(proj_path, "Total Critical and High Severity in RHC-A.png"))
#mydoc = addSlide( mydoc, "Status" )

###############################################
# check my layout names:
slide.layouts(mydoc)
slide.layouts(mydoc, 'Text')

body = paste("Total IAVA (", body3,") is calculated by taking the Sum of all IAVA in the Summary reports provided RHC-A (P) Cyber Security.
             Total Assets (",body4,") is calculated by taking the Sum of all 'ComputerName' in the Total Assets by HRP report, as pulled DHA. 
             Score is calculated by dividing the total IAVAs by the total Assets for each HRP.
             Some charts (with Total Vulnerabilities) have a modified HRP name, for which the format is the HRP _ SCORE. (eg.", Sum.Data[1,10], "),
             Average Score (",meancalc, ") is calculated by taking the mean score of all HRPs, excluding the Region score, Reference: Slide 2. 
             21 Day reporting is derived from the Severity reports pulled by RHC-A (P) Cyber Security.
             Open IAVMs (", Region.IAVM.Count, ") are the IAVMs from summary data, grouped by severity and year released.
             Vulnerability Eradication is calculated by averaging the time an IAVM 'issued' in 2016 took to 'leave' the reporting system.",
             sep="")


mydoc = addSlide( mydoc, "Text" )
mydoc = addTitle( mydoc, "Data Definitions")
mydoc = addParagraph( mydoc, value = body ) #1

filename = paste("IAVA Brief " ,sum.max.date, ".pptx",sep="")
filenamea = paste("S:/CISO/ACAS/ACAS IAVA Reports/IAVA Brief draft " ,sum.max.date, ".pptx",sep="")

writeDoc( mydoc, filename)
writeDoc( mydoc, file.path(report, filename))
writeDoc( mydoc, file.path(archive_data, filename))

Metric<-c('Total.IAVA', 'Total.Assets', 'Avg.IAVAperAsset', 'Open.Critical.IAVA', 'Open.Medium.IAVA',
          'Open.Low.IAVA', 'IAVA.Beyond.21', 'Vulnerability.Eradication.Days')
Critical<-select(summary, everything()) %>%
  filter(Severity=="Critical") %>% 
  summarise(Value=sum(Total))
Medium<-select(summary, everything()) %>%
  filter( Severity=="Medium") %>% 
  summarise(Value=sum(Total))
Low<-select(summary, everything()) %>%
  filter( Severity=="Low") %>% 
  summarise(Value=sum(Total))

Value<-c(sum(summary$Total)/2 , Region, body2, Critical, Medium, Low, 
         sum(Sev.Data$Total), mean(as.numeric(Week.Calc$days)))

#tmp.Calc<-as.data.frame(cbind(Value,Metric))
#tmp.Calc$Parent<-"Automatic"
#tmp.Calc$Section<-"ISSM"
#loc<-paste("S:/PMB/ISSM/2016/", as.numeric(toupper(format(Sys.Date(), format="%m"))), " ", toupper(format(Sys.Date(), format="%b")),"/Master M__AUTO.xlsx", sep="")
#write.xlsx(tmp.Calc, loc, row.name=F )
fldr<-'Y:/IAVA Reporting'

file.copy(file.path(report, filename), file.path(fldr),overwrite = T,copy.mode = TRUE, copy.date = TRUE)
dir.create(file.path(severity_data, sev.max.date), showWarnings = FALSE, recursive = FALSE, mode = "0777")
dir.create(file.path(summary_data, sum.max.date), showWarnings = FALSE, recursive = FALSE, mode = "0777")

newsumfldr=file.path(summary_data, sum.max.date)
for (i in summary.files){
  file.copy(file.path(summary_data,i), file.path(newsumfldr, i),overwrite = T,copy.mode = TRUE, copy.date = TRUE)
  file.remove(file.path(summary_data,i))
}
newsevfldr=file.path(severity_data, sev.max.date)
for (i in severity.files){
  file.copy(file.path(severity_data,i), file.path(newsevfldr, i),overwrite = T,copy.mode = TRUE, copy.date = TRUE)
  file.remove(file.path(severity_data,i))
}

