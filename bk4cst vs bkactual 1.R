

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


#1) take out the quarter name column AND PF column from each file
#2) adjourn the quarterly booking actual column from each quarter file by PID
alldfR <- list() #this will store reduced versions of the data frames
for(i in 1:length(alldf)){
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

#convert all NA's to zeros
qtrlyActual[is.na(qtrlyActual)] <- 0
#there is this symbole " - " that the file used to indicate zero, also need to replace that
tobedeleted <- qtrlyActual$`Q4 FY2015`[2]
qtrlyActual[qtrlyActual == tobedeleted] <- 0
#finally, rid of spaces by replacing it w zeros
qtrlyActual[qtrlyActual == ""] <- 0

#rid of rows that has actual bookings all 4 quarters; those are not of problem so no worries about them
bkngOK <- subset(qtrlyActual, rowSums(qtrlyActual == 0) == 0)
# bkngOK <- subset(qtrlyActual, (rowSums(is.na(qtrlyActual)) + rowSums(qtrlyActual == 0)) == 0 )
#qtrActFocus = quarterly actual booking for focus (the problematic ones)
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
#finally, we want to see which ones have not been getting booking at all for the year
missingAll <- subset(qtrActFocus, rowSums(qtrActFocus[, -1]) < 20)



#now let's work on the forecast file
path <- '~/cisco/projects/suman WOS/data/booking forecast/'
qtrlyForecast <- read.xlsx2(file = paste0(path, "CDO Raw Fcst $ Details PID Level.xlsx"), sheetIndex = 1, startRow = 2)

#convert all NA's to zeros
qtrlyForecast[is.na(qtrlyForecast)] <- 0
#there is this symbole " - " that the file used to indicate zero, also need to replace that
qtrlyForecast[qtrlyForecast == tobedeleted] <- 0
#finally, rid of spaces by replacing it w zeros
qtrlyForecast[qtrlyForecast == ""] <- 0

#convert all columns, except the PID column (which is the first one), into numeric values
qtrlyForecast[, -1] <- sapply(qtrlyForecast[, -1], function(x) as.numeric(as.character(x)))



#now we combine the missingAll to the forecast file to see which ones have NOT been booked but still being forecasted
noBking <- merge(missingAll, qtrlyForecast, by = "PID", all.x = T)
noBkw4cast <- subset(noBking, rowSums(noBking[, 6:length(noBking)], na.rm = T) != 0)


write.xlsx(noBkw4cast, file = "C:/Users/ali3/Documents/cisco/projects/suman WOS/analysis results/forecast without booking.xlsx", sheetName = 'noBkw4cast < 20', col.names = TRUE, row.names = FALSE, append = TRUE)



#next, we gotta read in the actual current EoL file and compare with the above file
path <- '~/cisco/projects/suman WOS/data/'
EoLnow <- read.csv(file = paste0(path, "9-25 Item_GLO_EOL_Data.csv"), stringsAsFactors = F)

# by default, the ID number is treated as numeric value; so change that
EoLnow[1] <- sapply(EoLnow[1], function(x) as.character(x))


#merge the EoL file with the forecast file w no booking
EoLw4cast <- merge(noBkw4cast, EoLnow, by = "PID", all.x = T)

#now find out which are the ones with no booking but actually are correctly identified as EoL
EoLnoBk <- subset(EoLw4cast, !is.na(INVENTORY_ITEM_ID))
write.xlsx(EoLnoBk, file = "C:/Users/ali3/Documents/cisco/projects/suman WOS/analysis results/forecast without booking.xlsx", sheetName = 'noBK is EoL', col.names = TRUE, row.names = FALSE, append = TRUE)


#now, remove above sections from the rest; the left over ones are the ones w no booking, not identified as EoL, and still getting forecast
noBKnotEoLw4cast <- subset(EoLw4cast, !(EoLw4cast$PID %in% EoLnoBk$PID))
write.xlsx(noBKnotEoLw4cast, file = "C:/Users/ali3/Documents/cisco/projects/suman WOS/analysis results/forecast without booking.xlsx", sheetName = 'noBkw4cast notEoL', col.names = TRUE, row.names = FALSE, append = TRUE)




