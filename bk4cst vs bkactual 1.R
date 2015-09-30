

#A separate function just to allow you to check whether a package is installed
EnsurePackage <- function(packageName)
{
  x <- as.character(packageName)
  if(!require(x, character.only = TRUE))
  {
    install.packages(pkgs = x, 
                     repos = "http://cran.r-project.org")
    require(x, character.only = TRUE)
  }
} 

#or the following (used this one instead)
EnsurePackage('xlsx')
library('xlsx')

#read the files
path <- '~/cisco/projects/suman WOS/data/actual booking csv/'
#allFiles <- list.files(path = path, pattern = "FY2015.csv")

alldf <- list()
i <- 1

#we need an outer loop since files are not in order of years (but by Quarter), we'll have to read the file in order of years
for(j in 3:5){
  print(paste0("j=", j))
  #j<-3
  allFiles <- list.files(path = path, pattern = paste0("FY201",j,".csv"))
  for(file in allFiles)
  {
    perpos <- which( strsplit( file, "")[[1]]==".")
    assign(
      gsub(" ","",substr(file, 1, perpos-1)), 
      
      # alldf[[i]] <- read.xlsx2(file = paste0(path, file), sheetIndex = 1, startRow = 4, colClasses = 'character')
      alldf[[i]] <- read.csv(header=TRUE, skip=3, paste(path, file, sep=""), stringsAsFactors = F)
      #skipped the first 3 rows cuz they are nothing but random stuff

    )
    #the first row is empty; so let's get rid of that
    alldf[[i]] <- alldf[[i]][, -1]
    
    i <- i+1
    print(paste0("i=",i))
  }
}

i <- 1 #reinitialize i just for personal convenience; no real meaning
test <- alldf[[i]]


# #read the files
# path <- '~/cisco/projects/suman WOS/data/actual booking/'
# allFiles <- list.files(path = path, pattern = "FY2015.csv")
# 
# alldf <- list()
# i <- 1
# 
# 
# for(file in allFiles)
# {
#   perpos <- which( strsplit( file, "")[[1]]==".")
#   assign(
#     gsub(" ","",substr(file, 1, perpos-1)), 
#     
#     # alldf[[i]] <- read.xlsx2(file = paste0(path, file), sheetIndex = 1, startRow = 4, colClasses = 'character')
#     alldf[[i]] <- read.csv(header=TRUE, skip=3, paste(path, file, sep=""), stringsAsFactors = F)
#     #skipped the first 3 rows cuz they are nothing but random stuff
#   )
#   i <- i+1
#   print(i)
# }

#nfile stores how many files were read
nfile <- length(alldf)

#1) take out the quarter name column AND PF column from each file
#2) adjourn the quarterly booking actual column from each quarter file by PID
alldfR <- list() #this will store reduced versions of the data frames
for(i in 1:nfile){
  print(i)
  #you'll get n data frames with the PID and booking amount for each quater within the fiscal year; n = number of files read
  temp <- alldf[[i]]
  qtrName <- unique(temp$Fiscal.Quarter.Name)
  #qtrName <- qtrName[-length(qtrName)] #get rid of the last character, which is ""
  names(temp)[length(temp)] <- qtrName #replace the last column name from "bookings.booked.amount" to the quarter name
  alldfR[[i]] <- temp[, c("PID", qtrName)]
  
}

#remove all unnecessary data
#rm(allFiles)
#remove all the individual data created by above code; the individual files created are usually in number date format; 
rm(perpos)
rm(file)
rm(temp)



#merge all quarterly data into one table
qtrlyActual <- Reduce(function(...) merge(..., by = 'PID', all = TRUE ), alldfR)

# #convert all NA's to zeros
# qtrlyActual[is.na(qtrlyActual)] <- 0
# #there is this symbole " - " that the file used to indicate zero, also need to replace that
# qtrlyActual <- gsub()
# 
# tobedeleted <- qtrlyActual$`Q4 FY2015`[2]
# qtrlyActual[qtrlyActual == tobedeleted] <- 0
#finally, rid of spaces by replacing it w NA
qtrlyActual[qtrlyActual == ""] <- NA

#rid of rows that has actual bookings all quarters; those are not of problem so no worries about them
bkngOK <- subset(qtrlyActual, rowSums(qtrlyActual[, -1], na.rm = T) >= 50)
# bkngOK <- subset(qtrlyActual, (rowSums(is.na(qtrlyActual)) + rowSums(qtrlyActual == 0)) == 0 )
#qtrActFocus = quarterly actual booking for focus (the problematic ones): the ones that has total 3 yrs of booking less than 50
qtrActFocus <- qtrlyActual[!(qtrlyActual$PID %in% bkngOK$PID),]

# 
# #all following data frames, will need to compare with FY16Q1 data to see if the discontinued PID is in fact really discontinued (no more booking)
# #only Q4 missing
# missingQ4only <- subset(qtrActFocus, rowSums(qtrActFocus == 0) == 1 & qtrActFocus$`Q4 FY2015` == 0)
# #Q3 & Q4 missing
# missingQ3Q4 <- subset(qtrActFocus, rowSums(qtrActFocus == 0) == 2 & qtrActFocus$`Q4 FY2015` == 0 & qtrActFocus$`Q3 FY2015` == 0)
# #Q2 - Q4 missing
# missingQ234 <- subset(qtrActFocus, rowSums(qtrActFocus == 0) == 3 & qtrActFocus$`Q1 FY2015` != 0)
# #now let's find the ones where Q1 & Q4 not missing, but Q2Q3 may both or either be missing
# MissingNotQ1Q4 <- subset(qtrActFocus, rowSums(qtrActFocus == 0) > 0 & qtrActFocus$`Q4 FY2015` != 0 & qtrActFocus$`Q1 FY2015` != 0)



#next, we gotta read in the actual current EoL file and compare with the above file
path <- '~/cisco/projects/suman WOS/data/'
EoLnow <- read.csv(file = paste0(path, "9-25 Item_GLO_EOL_Data.csv"), stringsAsFactors = F)

# by default, the ID number is treated as numeric value; so change that
EoLnow[1] <- sapply(EoLnow[1], function(x) as.character(x))
#rid of extra lines
EoLnow <- EoLnow[, !(names(EoLnow) %in% c("ORG", "ITEM_TYPE"))]
names(EoLnow)


#the file you'll write is:
# resultFileName <- "C:/Users/ali3/Documents/cisco/projects/suman WOS/analysis results/3yr 4cast w bk 50.csv"
resultDir <- "C:/Users/ali3/Documents/cisco/projects/suman WOS/analysis results/3 yr analysis/"

#now we combine the qtrActFocus to the EoL file to see which ones are marked as EoL in these problematic ones; then those are safe (sorta)
noBkwAllEoL <- merge(qtrActFocus, EoLnow, by = "PID", all.x = T, all.y = F)

#and then get rid of the ones that DOES have EoL lable; these are the OK ones; but still nice to know
noBkwEoL <- subset(noBkwAllEoL, !is.na(INVENTORY_ITEM_ID))
#write the file
#write.xlsx(noBkwEoL, file = resultFileName, sheetName = 'noBk wEoL', col.names = TRUE, row.names = FALSE, append = TRUE)
write.csv(noBkwEoL, file = paste0(resultDir, 'noBkwEoL.csv'), row.names = F)


#now filter out the OK ones, and you are left w the problematic ones; the ones that has no booking, and yet still NOT identified as EoL
noBknoEoL <- subset(noBkwAllEoL, !(PID %in% EoLnoBk$PID))
# write.xlsx(NoBknoEoL, 
#            file = resultFileName, 
#            sheetName = 'noBk noEoL', col.names = TRUE, row.names = FALSE, append = TRUE)
write.csv(noBknoEoL, file = paste0(resultDir, 'noBknoEoL.csv'), row.names = F)


#now let's work on the forecast file
path <- '~/cisco/projects/suman WOS/data/booking forecast/'
#qtrlyForecast <- read.xlsx2(file = paste0(path, "CDO Raw Fcst $ Details PID Level.xlsx"), sheetIndex = 1, startRow = 2)
qtrlyForecast <- read.csv(file = paste0(path, "CDO Raw Fcst $ Details PID Level.csv"), stringsAsFactors = F, skip = 1)


# #convert all NA's to zeros
# qtrlyForecast[is.na(qtrlyForecast)] <- 0
# #there is this symbole " - " that the file used to indicate zero, also need to replace that
# qtrlyForecast[qtrlyForecast == tobedeleted] <- 0
#finally, rid of spaces by replacing it w NA
qtrlyForecast[qtrlyForecast == ""] <- NA

# #convert all columns, except the PID column (which is the first one), into numeric values
# qtrlyForecast[, -1] <- sapply(qtrlyForecast[, -1], function(x) as.numeric(as.character(x)))



#finally, we have to map the ones w/o booking, no EoL lable, to the forecast; these are the ones that are supposed to be marked as EoL but not, so they are still getting forecasted
noBknoEoLwAll4cast <- merge(noBknoEoL, qtrlyForecast, by = "PID", all.x = T, all.y = F)

#there are some without booking, not labled as EoL, but has no forecast; take those out
#so you are left with rows that has no booking, not labled as EoL, but still has forecast
noBknoEolw4cast <- subset(noBknoEoLwAll4cast, 
                           rowSums(is.na(noBknoEoLwAll4cast[,(1+nfile+2+1):length(noBknoEoLwAll4cast)])) < (length(qtrlyForecast)-1) )
# write.xlsx(noBknoEolw4cast, file = resultFileName, 
#            sheetName = 'noBk noEoL w4cast', col.names = TRUE, row.names = FALSE, append = TRUE)
write.csv(noBknoEolw4cast, file = paste0(resultDir, 'noBknoEolw4cast.csv'), row.names = F)

#filter out the above ones, and you'll get a list of PIDS with no booking, not labled as EoL, but also no forecast as well
#could be a problem?
noBknoEoLno4cast <- subset(noBknoEoLwAll4cast, !(PID %in% noBknoEolw4cast$PID))
# write.xlsx(noBknoEolno4cast, file = resultFileName, 
#            sheetName = 'noBk noEoL no4cast', col.names = TRUE, row.names = FALSE, append = TRUE)
write.csv(noBknoEoLno4cast, file = paste0(resultDir, 'noBknoEolno4cast.csv'), row.names = F)


#should also find out the ones with no booking, identified as EoL, but still has forecast
noBkwEoLwAll4cast <- merge(noBkwEoL, qtrlyForecast, by = "PID", all.x = T, all.y = F)

#need to filter out the ones that actually has no forecast
noBkwEoLw4castIncludeZero <- subset(noBkwEoLwAll4cast, 
                          rowSums(is.na(noBkwEoLwAll4cast[,(1+nfile+2+1):length(noBkwEoLwAll4cast)])) < (length(qtrlyForecast)-2) )
write.csv(noBkwEoLw4castIncludeZero, file = paste0(resultDir, 'noBkwEoLw4castIncludeZero.csv'), row.names = F)

#also find out a subset of above dataframe that contains no booking, labeled as EoL, but still have positive large (>50) booking
noBkwEoLw4cast <- subset(noBkwEoLw4castIncludeZero, 
                         rowSums((noBkwEoLw4castIncludeZero[,(1+nfile+2+1):length(noBkwEoLw4castIncludeZero)]), na.rm = T) > 50 )
write.csv(noBkwEoLw4cast, file = paste0(resultDir, 'noBkwEoLw4cast.csv'), row.names = F)


#we are also intersted in finding which bookings has actual bookings >=50 is actually listed as EoL with forecast
bkwEoL <- merge(bkngOK, EoLnow, by = "PID")
#do these ones have forecast as well?
bkWEoLw4cast <- merge(bkwEoL, qtrlyForecast, by = "PID", all.x = T)
#filter out the ones that actually has forecast
bkWEoLw4cast <- bkWEoLw4cast <- subset(bkWEoLw4cast, 
                                       rowSums(is.na(bkWEoLw4cast[,(1+nfile+2+1):length(bkWEoLw4cast)])) < (length(qtrlyForecast)-2) )

#also need to find out booking >= 50, not labeled as EoL


#next, we are goign to create a result table to show the numbers; otherwise just tables of data is not really helpful...
result <- data.table(1,2,3,4,5,6,7,8)
#then rename all the columns with following code
#names(result) <- c("PF", "TAN", "as-is LAS", "to-be LAS", 'change in LAS', 'as-is OTS', 'to-be OTS', 'change in OTS')
setnames(result, c("V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8"), c("PF", "TAN", "as-is LAS", "to-be LAS", 'change in LAS', 'as-is OTS', 'to-be OTS', 'change in OTS'))




