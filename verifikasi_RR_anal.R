#!/usr/bin/Rscript
library(openxlsx)
library(readxl)
# library(RDCOMClient)
# install.packages("RDCOMClient")

# install.packages("remotes")
# remotes::install_github("KWB-R/kwb.geosalz")
# library("KWB-R/kwb.geosalz")
rm(list = ls())

is.integer0 <- function(x)
{
  is.integer(x) && length(x) == 0L
}

is.character0 <- function(x)
{
  is.character(x) && length(x) == 0L
}


mm = as.numeric(substr(Sys.Date(), 6, 7))-1
yy = (substr(Sys.Date(), 1, 4))
char_mon = c("JAN","FEB", "MAR", "APR",
             "MEI", "JUN", "JUL", "AGU",
             "SEP", "OKT", "NOV", "DEC")
mm_PO = as.character(mm+1)
mm_NE = as.character(mm-1)

if(nchar(mm_NE) < 2){
  mm_NE = paste0("0", mm_NE)
}
if(nchar(mm_PO) < 2){
  mm_PO = paste0("0", mm_PO)
}

list_pred = list.files("input/", pattern = mm_NE)
list_pred = list_pred[grep(list_pred, pattern = ".xlsx")]
list_anal = list.files("input/", pattern = mm_PO)
list_anal = list_anal[grep(list_anal, pattern = ".xlsx")]
# setwd("/mnt/f/WORK/Verifikasi/")



# ------------------------------- PREDIKSI DATA ---------------------------------- #
WD = loadWorkbook(file = paste0("input/",list_pred))
WDNAME = WD$sheet_names[grep(WD$sheet_names, pattern = "PRA")] 
WDNAME = WDNAME[grep(WDNAME, pattern = yy)] 
WDNAME = WDNAME[grep(WDNAME, pattern = char_mon[mm])] 
nsl = read_excel(paste0("input/",list_pred ), col_names = F, sheet = WDNAME)
RR = data.frame(Rcpt = as.numeric(as.matrix(nsl[8:47,14])),
                Recmwf = as.numeric(as.matrix(nsl[8:47,18])),
                Rreg = as.numeric(as.matrix(nsl[8:47,22])),
                Rrata2 = as.numeric(as.matrix(nsl[8:47,12])))
RR = RR[!is.na(RR$Rcpt),]

Sifat = data.frame(Scpt = as.numeric(as.matrix(nsl[8:47,16])),
                   Secmwf = as.numeric(as.matrix(nsl[8:47,20])),
                   Sreg = as.numeric(as.matrix(nsl[8:47,24])),
                   Srata2 = as.numeric(as.matrix(nsl[8:47,9])))
Sifat = Sifat[!is.na(Sifat$Scpt),]

Sifat2 = data.frame(S2cpt = substr(as.character(as.matrix(nsl[8:47,15])),1,1),
                    S2ecmwf = substr(as.character(as.matrix(nsl[8:47,19])),1,1),
                    S2reg = substr(as.character(as.matrix(nsl[8:47,23])),1,1),
                    S2rata2 = substr(as.character(as.matrix(nsl[8:47,8])),1,1)
)
Sifat2 = Sifat2[!is.na(Sifat2$S2cpt),]
Sta = as.character(as.matrix(nsl[8:47,2]))
Wil = as.character(as.matrix(nsl[8:47,1]))
ii = which(!is.na(Wil))
satu = rep(Wil[ii[1]],ii[2]-ii[1])
for(i in 1:(length(ii)-1)){
  if(i == (length(ii)-1)){
    satu = c(satu, rep(Wil[ii[i+1]], length(Wil)+1-ii[i+1]))
  }else{
    satu = c(satu, rep(Wil[ii[i+1]],ii[i+2]-ii[i+1]))
  }
}

Wil = satu[!is.na(Sta)]
Sta = Sta[!is.na(Sta)]


All_PRE = data.frame(Wil, Sta, RR, Sifat, Sifat2)

# -------------------------------------------------------------------------------- #

# ------------------------------- ANALISIS DATA ---------------------------------- #
WD = loadWorkbook(file = paste0("input/",list_anal))
WDNAME = WD$sheet_names[grep(WD$sheet_names, pattern = "ANAL")] 
WDNAME = WDNAME[grep(WDNAME, pattern = yy)] 
WDNAME = WDNAME[grep(WDNAME, pattern = char_mon[mm])] 
nsl = read_excel(paste0("input/",list_anal ), col_names = F, sheet = WDNAME)
RR = as.numeric(as.matrix(nsl[8:47,7]))
RR = RR[!is.na(RR)]
Sifat = as.numeric(as.matrix(nsl[8:47,9]))
Sifat = Sifat[!is.na(Sifat)]
Sifat2 = substr(as.character(as.matrix(nsl[8:47,8])),1,1)
Sifat2 = Sifat2[!is.na(Sifat2)]
All = data.frame(Wil, Sta, RR, Sifat, Sifat2)
# -------------------------------------------------------------------------------- #
# ------------------------------- Parsing Data Prediksi ---------------------------------- #
dtlist = list.files("input/2019/")
dtlist = dtlist[-grep(dtlist, pattern = ".xls")]

# parsing_xls_pred = function(x, ix){
#   # x = dtlist[1]; m = char_mon[mm]; ix = 1
#   wb <- loadWorkbook(file = paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx") )
#   iyang = which(wb$sheet_names == char_mon[mm])
#   writeData(wb, sheet = iyang,startCol = 3, startRow = 3, x = as.numeric(as.matrix(All_PRE[,ix+2])) )
#   # writeData(wb, sheet = iyang,startCol = 5, startRow = 3, x = as.numeric(as.matrix(All$RR)) )
#   saveWorkbook(file = paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx"), overwrite = T, wb = wb)
# }
# 
# for(i in 1:length(dtlist)){
#   parsing_xls_pred(x = dtlist[i], ix = i)
# }
# library(XLConnect)
# wb <- XLConnect::loadWorkbook(paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019_1.xlsx"))
# # appendWorksheet(wb, as.numeric(as.matrix(All$RR)), sheet = "SEP" )
# writeWorksheet(wb, as.numeric(as.matrix(All$RR)), sheet = "SEP", startCol = 5, startRow = 3, header = F)
# # startRow=(nrow(iris)+2)
# XLConnect::saveWorkbook(wb)
# ?XLConnect::appendWorksheet()

parsing_xls = function(x){
  # x = dtlist[3]
  # system(paste0("cp 'input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx' 'input/2019/",x,"/Curah Hujan/curah hujan.2019_1.xlsx'"))
  wb <- loadWorkbook(file = paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx") )
  iyang = which(wb$sheet_names == char_mon[mm])
  wb1 <- XLConnect::loadWorkbook(paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx"))
  XLConnect::writeWorksheet(wb1, as.numeric(as.matrix(All$RR)), sheet = wb$sheet_names[iyang], startCol = 5, startRow = 3, header = F)
  # hh = read_excel(paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019.xlsx") , sheet = wb$sheet_names[iyang], col_names = F)
  XLConnect::saveWorkbook(wb1)
  
  # writeData(wb, sheet = iyang,startCol = 5, startRow = 3, x = as.numeric(as.matrix(All$RR))  )
  # saveWorkbook(file = paste0("input/2019/",x,"/Curah Hujan/curah hujan.2019_2.xlsx"), overwrite = T, wb = wb)
  
}

for(i in 1:length(dtlist)){
  parsing_xls(dtlist[i] )
}

cat("ANALISIS OKE")