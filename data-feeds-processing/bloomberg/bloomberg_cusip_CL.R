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
cusips <- read.csv("C://Users/clj585/Downloads/cusips.csv", 
                    encoding="UTF-8", header=TRUE)

# Provide the Bloomberg value or field you'd like to retrieve
field <-	'EARN_ANN_DT_TIME_HIST_WITH_EPS'

# Provide the number of columns of array output
cols <- 6

# Provide the column labels
names <- c("Period", "Announcement_Date", "Announcement_Time",
           "Actual_EPS", "Comparable_EPS", "Estimated_EPS")
cusips$cusips<-gsub("/cusip/","",as.character(cusips$cusips))



###############################
### Function to Format Data ###
###############################

format_table <- function(table_input){
  table_input <- cusips
  table_input <- as.data.frame(table_input)
  colnames(table_input)[1] <- "cusips"
  table_input$cusips <- toupper(table_input$cusips)
  table_input$equity_name <- paste(table_input$cusips, "EQUITY", sep=" ")
  table_input$formula <- paste('=@BDS("', table_input$equity_name, '","', field, '")', sep="")
  #table_input$formula <- paste("'", table_input$formula, sep="")
  table_input$number <- seq(1,nrow(table_input))
  table2 <- table_input
  table2$cusips <- ""
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
# reformat tickers with formulas
cusips <- format_table(cusips)

# save results
write.xlsx(cusips, "C://Users/clj585/Downloads/B_input_cusip.xlsx", col.names=FALSE)

###########################################################
## Run all worksheet cells and then come back to program ##
###########################################################

#################################
### Reformat Bloomberg output ###
################################# 

# read in Bloomberg Input
bloomberg<-read.xlsx("C://Users/clj585/Downloads/B_input.xlsx", 
                     sheet = 1, startRow = 1, colNames = FALSE, rowNames = FALSE, 
                     detectDates = TRUE, skipEmptyRows = TRUE, skipEmptyCols = FALSE)

# add a final empty column if it was removed
bloomberg$last <- rep('NA', nrow(bloomberg)) 

# split file into separate tables
reps_no <- dim(bloomberg)[2]/(cols+1)
bloomberg <- split.default(bloomberg, rep(1:reps_no, each = cols+1))

# column names
names <- c(names, "Company")

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
for (i in bloomberg_df$Announcement_Time){
  time <- convert_to_time(i)
  print(time)
  
  times_list <- c(times_list, time)
}

bloomberg_df$Announcement_Time <- times_list

####################
### Save Results ###
####################
write.xlsx(bloomberg_df, "C://Users/clj585/Downloads/Bloomberg_cusip_final.xlsx", 
           col.names=TRUE)
