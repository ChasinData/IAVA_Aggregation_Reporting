
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


load(file.path(code.plc,"IAVA_WORKSPACE.RData"))


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

fldr<-'Y:/IAVA Reporting'

file.copy(file.path(report, filename), file.path(fldr),overwrite = T,copy.mode = TRUE, copy.date = TRUE)

