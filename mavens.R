## Sets the path for the project folder and loads the necessary libraries
setwd("~/Desktop/Mavens_Assessment_SearchData")
library(XLConnect)
library(dplyr)

## The environment is cleared of any existing objects
rm(list=ls())

## The Excel workbook is loaded and the Google search data is read into a data 
## frame called 'cleaning'.
wb <- loadWorkbook("Mavens_Assessment_SearchData.xlsx")
cleaning <- readWorksheet(wb, sheet = "UK Cleaning Keywords")

## Checks if there are any additonal sheets from previous executions of the 
## script and removes them if they are present.
if (existsSheet(wb, "Sorted By Searches") | existsSheet(wb, "Sorted by Categories")){
    removeSheet(wb, sheet = "Sorted By Searches")
    removeSheet(wb, sheet = "Sorted By Categories")
}

## Simplifies the column headers by removing excess periods and the word 
## 'Searches' from the headers.
colnames(cleaning)[5:28] <- gsub("Searches\\.","", colnames(cleaning)[5:28])

## This section searches the Keyword column for phrases that match with each of
## the assigned identifiers associated with the Category group and stories their 
## location in a vector.
appliance.ind <- grep("machine|vacuum|steam|hoover", cleaning$Keyword)
methods.ind <- grep("^cleaning|how to", cleaning$Keyword)
kitchen.ind <- grep("kitchen|oven", cleaning$Keyword)
bathroom.ind <- grep("bathroom|shower|tub|toilet", cleaning$Keyword)
stain.ind <- grep("stain|removal|blood|wine|tea|ink", cleaning$Keyword)
carpet.ind <- grep("carpet|rug", cleaning$Keyword)
supplies.ind <- grep("mop|feather|dust|soap|spray|wipes", cleaning$Keyword)
auto.ind <- grep("\\<car\\>|motorcycle", cleaning$Keyword)

## Each of the Keyword phrases pulled from the previous section gets assigned 
## the appropriate Category title.
cleaning$Category[stain.ind] <- "Stain Removal"
cleaning$Category[supplies.ind] <- "Supplies"
cleaning$Category[appliance.ind] <- "Appliances"
cleaning$Category[methods.ind] <- "Cleaning Methods"
cleaning$Category[kitchen.ind] <- "Kitchen"
cleaning$Category[bathroom.ind] <- "Bathroom"
cleaning$Category[carpet.ind] <- "Carpet"
cleaning$Category[auto.ind] <- "Auto"

## Keywords that did not match one of the prescribed categories get assigned as 
## 'Miscellaneous' as their category.
misc.ind <- is.na(cleaning$Category)
cleaning$Category[misc.ind] <- "Miscellaneous"

## The data frame is sorted to show the keywords with the highest average monthly
## searches in decreaseing order. A new sheet is added to the Excel file called 
## 'Sorted By Searches' that contains this sorted data frame. 
cleaning <- cleaning[order(cleaning$Avg.Monthly.Searches, decreasing = TRUE),]
createSheet(wb, name = "Sorted By Searches")
writeWorksheet(wb, cleaning, sheet = "Sorted By Searches")
saveWorkbook(wb)

## The data frame is sorted to show the keywords grouped by their categories in 
## alphabetical order and then sorted by their highest average monthly searches
## in decreasing order. The total search values for each category are collected 
## and added as a new column in the data frame. 
cleaning <- cleaning[order(cleaning$Category,-cleaning$Avg.Monthly.Searches),]
categories <- unique(cleaning$Category)
cat.total <- rep(NA, nrow(cleaning))
for (i in 1:length(categories)){
    cat.total[i] <- filter(cleaning, Category == categories[i]) %>% select(contains("Avg.Monthly.Searches")) %>% sum()
}
cleaning <- cbind(cleaning, cat.total)

## A new sheet is added to the Excel file called 'Sorted By Categories' that 
## contains this sorted data frame.
createSheet(wb, name = "Sorted By Categories")
writeWorksheet(wb, cleaning, sheet = "Sorted By Categories")
saveWorkbook(wb)