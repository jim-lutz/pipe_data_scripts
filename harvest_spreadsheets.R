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

# split path and file name 
path <- dirname(xls_files)
filename <- basename(xls_files)

head (path)
length(path)
head(filename)
length(filename)

DT_fn <- data.table(path,filename)

write.csv(DT_fn, file="list_of_xls.csv", row.names = FALSE)


# how many .XLS and how many .xls
sum(str_detect(xls_files, ".XLS$"))
sum(str_detect(xls_files, ".xls$"))
sum(str_detect(xls_files, ".xlsx$"))
# 328+1+58 = 387
# it's either .XLS or .xls or .xlsx

# libreoffice command and args for converting xls files to csv
syscommand <- "/opt/libreoffice5.3/program/soffice.bin"
sysargs <- c("--convert-to",
             "csv","")

# initialize a data.table and a counter
DT_all<-data.table()
n=0


# a loop to process all the xls_files
for( fn in xls_files) { 
  n = n+1
  cat(paste0(n," ",fn)) 
  
  # put the filename in single quotes
  sysargs[3] <- paste0("'",fn,"'")
  
  # a system call to convert xls files to a csv file in the working directory
  system2(command = syscommand, args = sysargs)
  
  # split path and file name 
  path <- dirname(fn)
  filename <- basename(fn)
  
  # make the csv file name, could be either XLS or xls
  csvfn <- str_replace(filename, "(.XLS)|(.xls)", ".csv")
  
  # looks like the file structure is not as clean as hoped.
  # read in the first 2 lines of the csv file as a data.table
  DT_csv <- data.table(read.csv(csvfn, header = FALSE, nrows = 2 ))

  # add path and filename as fields
  DT_csv[,path:=path][,filename:=filename]
  
  DT_all <- rbind(DT_all, DT_csv)
  
  # remove the DT_csv file
  if (file.exists(csvfn)) file.remove(csvfn)

}

  
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
