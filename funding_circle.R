#Load required packages
install.packages("xlsx")
library(xlsx)
require(xlsx)
install.packages("readxl")
library("readxl")
install.packages("XLConnect")
library(XLConnect)

#Set directory
setwd("<input working directory here>")

#Read all files in the directory and merge them
file_list <- list.files()

for (file in file_list){
  
  # if the merged dataset doesn't exist, create it
  if (!exists("dataset")){
    dataset <- read_excel(file, sheet = "Loan Book", range = cell_cols("B:R"),col_names = TRUE, col_types = c("text","text","text","date","text","text","text","numeric","text","text","numeric","text","text","text","text","text","numeric"))
  }
  
  # if the merged dataset does exist, append to it
  if (exists("dataset")){
    temp_dataset <-read_excel(file, sheet = "Loan Book", range = cell_cols("B:R"),col_names = TRUE, col_types = c("text","text","text","date","text","text","text","numeric","text","text","numeric","text","text","text","text","text","numeric"))
    dataset<-rbind(dataset, temp_dataset)
    rm(temp_dataset)
  }
  
}

#Remove NA rows and column headers of each read file
dataset<-dataset[!(dataset$X__1=="Loan name"),]
dataset <- dataset[!is.na(dataset$X__1),]

#Rename column headers
library(plyr)
dataset <- rename(dataset, c("Funding Circle"="ID", "X__1"="Loan name", "X__2"="Credit rating", "X__3"="Funding date",  "X__4"="Industry", "X__5"="Use of funds", "X__6"="State",  "X__7"="Years of operation", "X__8"="Principal", "X__9"="Interest Rate", "X__10"="Term", "X__11"="Payment Status", "X__12"="Outstanding balance", "X__13"="FICO score", "X__14"="Asset coverage ratio", "X__15"="Debt coverage ratio", "X__16"="Number of guarantors"))

#Create new column and add abbreviations for payment status
unique(dataset$`Payment Status`)
df <- dataset[grep("[[:digit:]]", dataset$`Payment Status`), ]
unique(df$`Payment Status`)
dataset$status = NA 
ind1 = dataset$`Payment Status` == 'paid_in_full' | dataset$`Payment Status` == 'paid_in_full*'
ind2 = dataset$`Payment Status` == 'prepaid' 
ind3 = dataset$`Payment Status` == 'defaulted' | dataset$`Payment Status` == 'default' 
ind4 = dataset$`Payment Status` == 'late_11_30' 
ind5 = dataset$`Payment Status` == 'late_60+' 
ind6 = dataset$`Payment Status` == 'current'
ind7 = dataset$`Payment Status` == 'repurchased'
ind8 = dataset$`Payment Status` == 'late_31_60'
ind9 = dataset$`Payment Status` == 'repaid'
ind10 = dataset$`Payment Status` == 'late' 

dataset$status[ind1] = 'pf'
dataset$status[ind2] = 'pp'
dataset$status[ind3] = 'd'
dataset$status[ind4] = 'l11-30'
dataset$status[ind5] = 'l60+'
dataset$status[ind6] = 'c'
dataset$status[ind7] = 'r'
dataset$status[ind8] = 'l31-60'
dataset$status[ind9] = 're'
dataset$status[ind10] = 'l'

unique(dataset$status)

#Concatenate payment status for a loan ID
data <- within(dataset, {
  Payment_Status_List <- ave(status, ID, FUN = function(x) paste(x, collapse = "-"))
})

test <- data[data$ID == 'fe44d2c4-2620-4ebb-b060-8821e6665622',]
