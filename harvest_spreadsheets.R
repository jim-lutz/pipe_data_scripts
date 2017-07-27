# harvest_spreadsheets.R
# Copies pipe test data out of spreadsheets and assembles into one large data file.
# Jim Lutz "Fri Jun 23 10:21:23 2017"

# make sure all packages loaded and start logging
source("setup.R")

# set the working directory names 
source("setup_wd.R")

# get a list of all the xls files in the wd_data subdirectories
xls_files <- list.files(path = wd_data, pattern = ".+.xls", 
                        full.names = TRUE, recursive = TRUE,
                        ignore.case = TRUE)
str(xls_files)
# 387 files

# try reading 1st xls file
xls_files[1]
install.packages("readxl")
library(readxl)

list.files()

# try sample file
read_excel("12_CPVCIN_2G_T1.XLS")
# Not an excel file
# Error in read_fun(path = path, sheet = sheet, limits = limits, shim = shim,  : 
#   Failed to open 12_CPVCIN_2G_T1.XLS
                  
read_xls("12_CPVCIN_2G_T1.XLS")
# Not an excel file
# Error in read_fun(path = path, sheet = sheet, limits = limits, shim = shim,  : 
#   Failed to open 12_CPVCIN_2G_T1.XLS

# however libreoffice calc can open the file.

# try gdata
install.packages("gdata")
library(gdata)
read.xls("12_CPVCIN_2G_T1.XLS", verbose = TRUE)
# uses perl. perl choked on it.

# try XLConnect
# needs rJava
install.packages("rJava")
library(rJava)

# configure: error: Cannot compile a simple JNI program. See config.log for details.
# 
# Make sure you have Java Development Kit installed and correctly registered in R.
# If in doubt, re-run "R CMD javareconf" as root.

# sudo R CMD javareconf
# conftest.c:1:17: fatal error: jni.h: No such file or directory
# compilation terminated.
 
# use Synaptic to install r-cran-rjava
library(rJava)

install.packages("XLConnect")
# * installing *source* package ‘XLConnectJars’ ...

# ** testing if installed package can be loaded
# Segmentation fault (core dumped)
# ERROR: loading failed

# try xlsReadWrite
install.packages("xlsReadWrite")
# Warning in install.packages :
#   package ‘xlsReadWrite’ is not available (for R version 3.4.0)

# try gnumeric
install.packages("gnumeric")
library(gnumeric)
read.gnumeric.sheet("12_CPVCIN_2G_T1.XLS",head = TRUE, quiet = FALSE)
  # LANG=C /usr/bin/ssconvert --export-type=Gnumeric_stf:stf_assistant  -O "locale=C format=automatic separator=, eol=unix sheet='Sheet1'" "12_CPVCIN_2G_T1.XLS" fd://1  
  # Unknown BIFF type in BOF 9
  # Unknown BOF (6)
  # Loading file:///home/jiml/HotWaterResearch/projects/Pipe%20Test%20Data/pipe_data_scripts/12_CPVCIN_2G_T1.XLS failed
  # Error in read.table(file = file, header = header, sep = sep, quote = quote,  : 
  #                       no lines available in input

# from https://msdn.microsoft.com/en-us/library/dd906793(v=office.12).aspx
# the first 2 bytes must bee
# An unsigned integer that specifies the BIFF version of the file. The value MUST be 0x0600. 
# used wxHexEditor to look at it. 1st 2 bytes were 09 00
# test changing 1st 2 bytes to 06 00
read.gnumeric.sheet("12_CPVCIN_2G_T1.3.XLS",head = TRUE, quiet = FALSE)
  # E Unsupported file format.
  # Error in read.table(file = file, header = header, sep = sep, quote = quote,  : 
  #   no lines available in input
# that didn't work

# tried ssconvert from gnumeric with various importers but couldn't get it to work
# see if localc has command line macros, since it can open the files.

# see soffice --help 
# looks like 
#   soffice --headless --convert-to csv 12_CPVCIN_2G_T1.XLS 
# does the trick!
# not anymore?

# was able to get it wiht unoconv
# with some problem on first try.
# $ unoconv --format csv 12_CPVCIN_2G_T1.XLS 
# Error: Unable to connect or start own listener. Aborting.
# $ unoconv --version
# unoconv 0.7
# Written by Dag Wieers <dag@wieers.com>
# Homepage at http://dag.wieers.com/home-made/unoconv/
#   
#   platform posix/linux
# python 3.5.2 (default, Nov 17 2016, 17:05:23) 
# [GCC 5.4.0 20160609]
# LibreOffice 5.1.6.2
# $ unoconv --format csv 12_CPVCIN_2G_T1.XLS 
# $ 


#======



load(file=paste0(wd_data,"DT.Rdata"))

tables()

# check reading correct Dhw and DhwBu
names(DT_HPWH)
DT_HPWH[, list(HPWH.ElecDhw=sum(HPWH.ElecDhw), HPWH.ElecDhwBU=sum(HPWH.ElecDhwBU))]
names(DT_HPWH01)
DT_HPWH01[, list(Dhw=sum(Dhw), DhwBU=sum(DhwBU))]

names(DT_ER)
DT_ER[, list(ER.ElecDhw=sum(ER.ElecDhw), ER.ElecDhwBU=sum(ER.ElecDhwBU))]
names(DT_ER01)
DT_ER01[, list(Dhw=sum(Dhw), DhwBU=sum(DhwBU))]

# keep DT_HPWH & DT_ER data.tables
l_tables <- tables()$NAME
rm(list = l_tables[grepl("*0[1-5]", l_tables)])

# merge table of electricity use & hot water for HPWH and ER
DT_EHW <-merge(DT_ER[,list(HoY,JDay,Mon,Day,Hr,sDOWH,nPeople,WH.Total=ER.WH.Total,ER.ElecDhw,ER.ElecDhwBU)],
               DT_HPWH[,list(HoY,HPWH.ElecDhw, HPWH.ElecDhwBU)], 
               by = "HoY")

# assume ElecTot and ElecDhwBU are same units, and are separate
DT_EHW[ , ER.Elec:= ER.ElecDhw + ER.ElecDhwBU]
DT_EHW[ , HPWH.Elec:= HPWH.ElecDhw + HPWH.ElecDhwBU]

# rearrange data for boxplots of hourly _use
DT_EHW[ , list(HoY,Hr, GPH=WH.Total, ER=ER.Elec, HPWH=HPWH.Elec)]

# melt into long format
DT_mEHW <- melt(DT_EHW[ , list(HoY,Hr, GPH=WH.Total, ER=ER.Elec, HPWH=HPWH.Elec)], id.vars = c("HoY","Hr"))
DT_mEHW[, list(variable=unique(variable))]
#    variable
# 1:      GPH
# 2:       ER
# 3:     HPWH

# change names
setnames(DT_mEHW, c("value", "variable"), c("hourly.use", "type.use"))

# boxplots of hourly use by type of use and hour of day
p <- ggplot(data = DT_mEHW )
p <- p + geom_boxplot( aes(y = hourly.use, x = as.factor(Hr),
                           fill = factor(type.use), 
                           color = factor(type.use),
                           dodge = type.use),
                       position = position_dodge(width = .7),
                       varwidth = TRUE)
p <- p + scale_fill_manual(values=c("#DDDDFF", "#FFDDDD", "#DDFFDD"),name="use")
p <- p + scale_color_manual(values=c("#0000FF", "#FF0000", "#00FF00"),name="use")
p <- p + ggtitle("Hot Water and Electricity Use") + labs(x = "hour", y = "Hourly Use")
p <- p + scale_x_discrete(breaks=1:24,labels=1:24)
p

ggsave(filename = paste0(wd_charts,"/Use_by_hour.png"), plot = p)

# now plot one day, high and low
DT_EHW[, list(GPD=sum(WH.Total)), by = JDay][order(GPD)]

# heavy day, JDay==138
DT_EHW[JDay==138,]
DT_heavy <- DT_EHW[JDay==138,list(Hr,GPH=WH.Total,ER=ER.Elec,HPWH=HPWH.Elec)]
str(DT_heavy)
p <- ggplot(data = DT_heavy )
p <- p + geom_line(aes(x=Hr, y=GPH), color="blue", size=2)
p <- p + geom_line(aes(x=Hr,y=ER), color="red")
p <- p + geom_line(aes(x=Hr,y=HPWH), color="green")
p <- p + ggtitle("Hot Water and Electricity Use (Mon, 5/18, 6 people)") + labs(x = "hour", y = "Hourly Use")
p <- p + scale_x_continuous(breaks=1:24,labels=1:24)
p

ggsave(filename = paste0(wd_charts,"/HW_elec_heavy.png"), plot = p)

# light day, JDay==15
DT_EHW[JDay==15,]
DT_light <- DT_EHW[JDay==15,list(Hr,GPH=WH.Total,ER=ER.Elec,HPWH=HPWH.Elec)]
str(DT_light)
p <- ggplot(data = DT_light )
p <- p + geom_line(aes(x=Hr, y=GPH), color="blue", size=2)
p <- p + geom_line(aes(x=Hr,y=ER), color="red")
p <- p + geom_line(aes(x=Hr,y=HPWH), color="green")
p <- p + ggtitle("Hot Water and Electricity Use (Thurs, 1/15, 2 people)") + labs(x = "hour", y = "Hourly Use")
p <- p + scale_x_continuous(breaks=1:24,labels=1:24)
p

ggsave(filename = paste0(wd_charts,"/HW_elec_light.png"), plot = p)
