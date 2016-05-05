
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
#folders.sev=list.dirs(path = severity_data, full.names = TRUE, recursive = TRUE)
#folders.sum=list.dirs(path = summary_data, full.names = TRUE, recursive = TRUE)


#Get File names
severity.files <- list.files(severity_data, pattern="*.csv", recursive = T, full.names=T, all.files = T, include.dirs = T)
#EMPTY DATAFRAMES
severity=as.data.frame(matrix( ,0,3))

#nvl=as.data.frame(matrix( ,0,2))
#no.Vulnerabilities =matrix( ,0,2)
no.Vulnerabilities.lst=as.list(NULL)

#BIND ALL FILES

for (i in  severity.files){
  #i=severity.files[1]
  df<- read.csv(i, header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  #df<- read.csv(i, header = TRUE, sep = ",", quote = "\"", fill = TRUE, comment.char = "")
  df$HRP.name <- i
  #df$HRP.name <- sub(" Severity Summary for Alerts.csv","",i)
  df$As.of.Date=file.mtime(i); df$As.of.Date=(format(df$As.of.Date,"%d-%b-%Y"))
  severity=as.data.frame(rbind(severity,df))
}

severity[grepl("Barracks", severity$HRP.name, ignore.case=T), "HRP.Name"] <- "Carlisle Barracks"
severity[grepl("benning", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Benning"
severity[grepl("brag", severity$HRP.name, ignore.case=TRUE), "HRP.Name"] <- "Ft. Bragg"
severity[grepl("campbell", severity$HRP.name, ignore.case=TRUE), "HRP.Name"] <- "Ft. Campbell"
severity[grepl("drum", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Drum"
severity[grepl("Eustis", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Eustis"
severity[grepl("Gordon", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Gordon"
severity[grepl("Jackson", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Jackson"
severity[grepl("Knox", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Knox"
severity[grepl("Lee", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Lee"
severity[grepl("Mead", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Meade"
severity[grepl("myers", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Myers"
severity[grepl("ruck", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Rucker"
severity[grepl("stewart", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Ft. Stewart"
severity[grepl("ftdtl", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "FTDTL"
severity[grepl("Kirk", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Kirk"
severity[grepl("redstone", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "Redstone"
severity[grepl("Regional", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "RHC-A"
severity[grepl("West Point", severity$HRP.name,ignore.case=TRUE ), "HRP.Name"] <- "West Point"



severity$As.of.Date=as.POSIXct(severity$As.of.Date, format = "%d-%b-%Y")

sev.max.date=max(severity$As.of.Date)

severity$HRP.name=severity$HRP.Name




severity = filter(severity, severity$Severity == "High" |
                    severity$Severity == "Critical")

##################### Get severity data
bck=select(severity, everything()) %>%
  group_by(As.of.Date, HRP.name) %>%
  summarise(Total=sum(Count))%>%
  arrange(-Total)

code.plc='H:/Code Source/IAVA_Aggregation_Reporting'
save(bck,file=file.path(code.plc,"Historical Data BACKUP.RDS"))
save(bck,file=file.path(archive_data,"Historical Data BACKUP.RDS"))

