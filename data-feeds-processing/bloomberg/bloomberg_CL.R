#########################
# Bloomberg Data Arrays #
#  Draft: 8/24/2020     #
#########################
# This sheet will tranpose a list of ticker 
# symbols to a csv file you can paste into
# the Bloomberg Excel Add-In to retrieve 
# array data.

# Thereafter, you can read in the Bloomberg 
# output formatted as separate tables to 
# produce a data frame.

# clear workspace
rm(list=ls())

# libraries used
library(openxlsx)
library(stringr)
library(plyr)
library(dplyr)
library(tidyr)
library(readr)
library(chron)


##############
### Inputs ###
##############

# read in ticker symbol file
tickers <- read.csv("C://Users/clj585/Downloads/tickers.csv", 
                    encoding="UTF-8", header=TRUE)

# Provide the Bloomberg value or field you'd like to retrieve
field <-	'EARN_ANN_DT_TIME_HIST_WITH_EPS'

# Provide the number of columns of array output
cols <- 6

# Provide the column labels
names <- c("Period", "Announcement_Date", "Announcement_Time",
           "Actual_EPS", "Comparable_EPS", "Estimated_EPS")


###############################
### Function to Format Data ###
###############################

format_table <- function(tickers_subset){
  table_input <- tickers_subset #tickers
  table_input <- as.data.frame(table_input)
  colnames(table_input)[1] <- "tickers"
  table_input$tickers <- toupper(table_input$tickers)
  table_input$equity_name <- paste(table_input$tickers, "US EQUITY", sep=" ")
  table_input$formula <- paste('\\=@BDS("', table_input$equity_name, '","', field, '")', sep="")
  #table_input$formula <- paste("'", table_input$formula, sep="")
  table_input$number <- seq(1,nrow(table_input))
  table2 <- table_input
  table2$tickers <- ""
  table2$equity_name <- ""
  table2$formula <- ""
  table_new <- data.frame()
  for (i in 1:cols+1){
    print(i)
    table_new <- rbind(table_new, table2)
  }
  rm(table2)
  table_input <- rbind(table_input, table_new)
  rm(table_new)
  table_input <- table_input[order(table_input$number),]
  table_input <- table_input[,c(1,3)]
  table_input <- as.data.frame(t(table_input))
  return(as.data.frame(table_input))
}

#########################################
### Reformat and Save Bloomberg Input ###
#########################################

# To be done 50+ times... 
for (i in 1:54){
  # Get correct indexes to splice data frame
  start_row_ind <- 2000*(i-1)+1
  if (i == 54){
    end_row_ind = dim(tickers)[1]
  }
  else{
    end_row_ind = start_row_ind + 2000-1
  }
  
  print(paste(start_row_ind, end_row_ind, sep=", "))
  
  # Subset the ticker data frame
  subset <- data.frame(tickers[start_row_ind:end_row_ind,])
  names(subset)[1] <- 'Tickers'
  print(dim(subset)[1])
  
  # reformat tickers with formulas
  subset <- format_table(subset)
  
  # save results
  filePath <- paste0("C://Users/clj585/Downloads/spliced_tic/B_input", "_", i, ".xlsx")
  print(tickers$Tickers[end_row_ind])
  write.xlsx(subset, filePath, col.names=FALSE)
}




# reformat tickers with formulas
tickers <- format_table(tickers)

# save results
write.xlsx(tickers, "C://Users/clj585/Downloads/B_input.xlsx", col.names=FALSE)

###########################################################
## Run all worksheet cells and then come back to program ##
###########################################################


###########################################################
## Extra verification: check for inactivated Excel cells ##
###########################################################


# read in ticker symbol file
tickers <- read.csv("C://Users/clj585/Downloads/tickers.csv", 
                    encoding="UTF-8", header=TRUE)

# Provide the Bloomberg value or field you'd like to retrieve
field <-	'EARN_ANN_DT_TIME_HIST_WITH_EPS'

# Provide the number of columns of array output
cols <- 6

# Provide the column labels
names <- c("Period", "Announcement_Date", "Announcement_Time",
           "Actual_EPS", "Comparable_EPS", "Estimated_EPS")

# Empty list to store problematic ones 
check_wrksht_list = c()

for (j in 5:54){
  # read in Bloomberg Input
  names <- c("Period", "Announcement_Date", "Announcement_Time",
             "Actual_EPS", "Comparable_EPS", "Estimated_EPS")
  fpth <- paste0("C://Users/clj585/Downloads/spliced_tic/B_input", "_", j, ".xlsx")
  bloomberg<-read.xlsx(fpth,#"C://Users/clj585/Downloads/spliced_tic/B_input_.xlsx", 
                       sheet = 1, startRow = 1, colNames = FALSE, rowNames = FALSE, 
                       detectDates = TRUE, skipEmptyRows = TRUE, skipEmptyCols = FALSE)
  
  # rbind extra columns if not divisible by 6
  if(dim(bloomberg)[2]%%(cols+1)>0){
    diff = 14000-dim(bloomberg)[2]
    bloomberg <- cbind(bloomberg, data.frame(matrix(NA, nrow = dim(bloomberg)[1], 
                                                    ncol = diff-1)))
  } 
  
  # add a final empty column if it was removed
  bloomberg$last <- rep('NA', nrow(bloomberg)) 
  
  # split file into separate tables
  reps_no <- dim(bloomberg)[2]/(cols+1)
  testing <- split.default(bloomberg, rep(1:reps_no, each = cols+1))
  bloomberg <- split.default(bloomberg, rep(1:reps_no, each = cols+1))
  
  # column names
  names <- c(names, "cticker")
  
  # combine each table into a dataframe
  bloomberg_df <- data.frame()
  for (i in 1:length(bloomberg)) {
    table <- bloomberg[[i]]
    table <- table[rowSums(is.na(table)) != ncol(table),]
    table$company <- table[1,1]
    table <- table[-1,-(cols+1)]
    colnames(table)<- names
    bloomberg_df <- rbind(bloomberg_df, table)
  }
  
  tempdf <- bloomberg_df
  
  # The check - create temp datafratme
  sav<-tempdf %>% filter(grepl(pattern = "=@BDS", x = tempdf$Period, ignore.case = TRUE))
  print(dim(sav)[1])
  if (dim(sav)[1]>0){
    print(c("check table ",j))
    check_wrksht_list <- c(check_wrksht_list, sav)
  }
}

#################################
### Reformat Bloomberg output ###
#################################

# Just work with what we have right now 
complete_num = c(1:10, 12:15)

# Saved data frames
save_bloomberg_dfs <- data.frame()

for (j in complete_num){
  # Tickers we want in each of the spliced excel sheets
  want_ind <- c()
  
  # Total list of indices in each spliced excel sheet
  total_ind <- c(1:2000)
  
  # read in Bloomberg Input
  names <- c("Period", "Announcement_Date", "Announcement_Time",
            "Actual_EPS", "Comparable_EPS", "Estimated_EPS")
  fpth <- paste0("C://Users/clj585/Downloads/spliced_tic/B_input", "_", 
                 j, ".xlsx")
  bloomberg<-read.xlsx(fpth,#"C://Users/clj585/Downloads/spliced_tic/B_input_.xlsx", 
                     sheet = 1, startRow = 1, colNames = FALSE, rowNames = FALSE, 
                     detectDates = TRUE, skipEmptyRows = TRUE, skipEmptyCols = FALSE)

  # rbind extra columns if not divisible by 6
  if(dim(bloomberg)[2]%%(cols+1)>0){
    diff = 14000-dim(bloomberg)[2]
    bloomberg <- cbind(bloomberg, data.frame(matrix(NA, nrow = dim(bloomberg)[1], 
                                                  ncol = diff-1)))
  } 

  # add a final empty column if it was removed
  bloomberg$last <- rep('NA', nrow(bloomberg)) 

  # split file into separate tables
  reps_no <- dim(bloomberg)[2]/(cols+1)
  testing <- split.default(bloomberg, rep(1:reps_no, each = cols+1))
  bloomberg <- split.default(bloomberg, rep(1:reps_no, each = cols+1))

  # column names
  names <- c(names, "cticker")
  
  for (k in 1:length(bloomberg)){
    checkdf <- bloomberg[[k]]
    if (!str_detect(checkdf[2,1], "#N/A")){
      want_ind <- c(want_ind, k)
    }
  }
  
  # Compare two lists and subset bloomberg which is a list 
  newlist <- bloomberg[match(want_ind, total_ind)]

  # combine each table into a dataframe
  bloomberg_df <- data.frame()
  for (i in 1:length(newlist)) {#:length(bloomberg)) {
    table <- newlist[[i]]#bloomberg[[i]]
    table <- table[rowSums(is.na(table)) != ncol(table),]
    table$company <- table[1,1]
    table <- table[-1,-(cols+1)]
    colnames(table)<- names
    bloomberg_df <- rbind(bloomberg_df, table)
  }
  
  print(paste("j is: ", j,"# kept tickers from excel sheet is ", 
              length(newlist)))
  print(paste("# rows bloomberg_df: ", dim(bloomberg_df)[1]))
  save_bloomberg_dfs <- rbind(save_bloomberg_dfs, bloomberg_df)
}


#################
### Fix Times ###
#################
convert_to_time <- function(value_input){
  value <- value_input
  if(is.na(value)==TRUE | grepl("[[:alpha:]]",value)==TRUE){
    print(value)
    return(value)
  }
  else {
    value <- as.numeric(value)
    value <- times(value)
    value <- paste("time:", value, sep=" ")
    return(value)
  }
}

times_list <- vector()
for (i in save_bloomberg_dfs$Announcement_Time){
  tim <- convert_to_time(i)
  print(tim)
  times_list <- c(times_list, tim)
}

save_bloomberg_dfs$Announcement_Time <- times_list


times_list <- vector()
for (i in bloomberg_df$Announcement_Time){
  time <- convert_to_time(i)
  print(time)
  
  times_list <- c(times_list, time)
}

bloomberg_df$Announcement_Time <- times_list

####################
### Save Results ###
####################
write.xlsx(save_bloomberg_dfs, 
           "C://Users/clj585/OneDrive - Northwestern University/bbg_final_test.xlsx", 
           col.names=TRUE)

write.xlsx(bloomberg_df, "C://Users/clj585/Downloads/Bloomberg_final_1.xlsx", 
           col.names=TRUE)
