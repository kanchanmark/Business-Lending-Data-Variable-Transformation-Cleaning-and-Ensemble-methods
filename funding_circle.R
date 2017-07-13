#Load required packages
#install.packages("xlsx")
library(xlsx)
require(xlsx)
#install.packages("readxl")
library("readxl")
#install.packages("XLConnect")
library(XLConnect)

#Set directory
setwd("C:/Users/DELL/Downloads/UMCP/PiAnalytics/Raw Data/Raw Data/")

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
dataset <- rename(dataset, c("Funding Circle"="ID", "X__1"="Loan_name", "X__2"="Credit_rating", "X__3"="Funding_date",  "X__4"="Industry", "X__5"="Use_of_funds", "X__6"="State",  "X__7"="Years_of_operation", "X__8"="Principal", "X__9"="Interest_Rate", "X__10"="Term", "X__11"="Payment_Status", "X__12"="Outstanding_balance", "X__13"="FICO_score", "X__14"="Asset_coverage_ratio", "X__15"="Debt_coverage_ratio", "X__16"="Number_of_guarantors"))

#Create new column and add abbreviations for payment status
unique(dataset$`Payment_Status`)
df <- dataset[grep("[[:digit:]]", dataset$`Payment_Status`), ]
unique(df$`Payment_Status`)
dataset$status = NA 
ind1 = dataset$`Payment_Status` == 'paid_in_full' | dataset$`Payment_Status` == 'paid_in_full*'
ind2 = dataset$`Payment_Status` == 'prepaid' 
ind3 = dataset$`Payment_Status` == 'defaulted' | dataset$`Payment_Status` == 'default' 
ind4 = dataset$`Payment_Status` == 'late_11_30' 
ind5 = dataset$`Payment_Status` == 'late_60+' 
ind6 = dataset$`Payment_Status` == 'current'
ind7 = dataset$`Payment_Status` == 'repurchased'
ind8 = dataset$`Payment_Status` == 'late_31_60'
ind9 = dataset$`Payment_Status` == 'repaid'
ind10 = dataset$`Payment_Status` == 'late' 

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

data$Loan_Status <- ifelse(grepl("d", data$Payment_Status_List, ignore.case = T), "Default","Non-Default") 

#Undersample data
library(unbalanced)

unique(data$Loan_Status)
data$status_new <-ifelse(data$Loan_Status=="Default",1,0)

input <- data[,-c(1,2,5,6,12,18,19,20,21)]
output <- data$status_new
loan <- ubUnder(X = input, Y = output, perc = 25, method = "percPos")
loan <- cbind(loan$X, loan$Y)
loan$status <- loan$`loan$Y`
loan$`loan$Y` <- NULL
remove(input)
remove(output)
loan$`Funding date` <- NULL
loan$State <- NULL
loan$`Credit_rating` <- as.factor(loan$`Credit_rating`)
loan$Principal <- as.numeric(loan$Principal)
loan$`Interest_Rate` <- as.numeric(loan$`Interest_Rate`)
loan$`Outstanding_balance` <- as.numeric(loan$`Outstanding_balance`)
loan$`FICO_score` <- as.numeric(loan$`FICO_score`)
loan$`Asset_coverage_ratio` <- as.numeric(loan$`Asset_coverage_ratio`)
loan$`Debt_coverage_ratio` <- as.numeric(loan$`Debt_coverage_ratio`)
loan$status <- as.factor(loan$status)

#Replace NA values by mean values
na_columns <- colnames(loan)[ apply(loan, 2, anyNA) ]

for(i in 1:ncol(loan)){
  loan[is.na(loan[,i]), i] <- mean(loan[,i], na.rm = TRUE)
}

#Sampling data
set.seed(123457)

train<-sample(nrow(loan),0.7*nrow(loan))
data_train<-loan[train,]
data_val<-loan[-train,]


#Bagging
library(randomForest)
# We first do bagging (which is just RF with m = p)
set.seed(123457)
bag.LoanData=randomForest(status~.,data=data_train,mtry=11,importance=TRUE,na.action = na.exclude)
bag.LoanData
yhat.bag = predict(bag.LoanData,newdata=data_val)
yhat.bag
plot(yhat.bag, data_val$status)
abline(0,1)
yhat.test=data_val$status
yhat.test
(c = table(yhat.test,yhat.bag))
(acc = (c[1,1]+c[2,2])/sum(c))
importance(bag.LoanData)
varImpPlot(bag.LoanData)

#Random forest
set.seed(123457)
rf.LoanData=randomForest(status~.,data=data_train,mtry=2,importance=TRUE)
rf.LoanData
yhat.rf = predict(rf.LoanData,newdata=data_val)
yhat.rf
yhat.test = data_val$status
yhat.test
(c = table(yhat.test,yhat.rf))
(acc = (c[1,1]+c[2,2])/sum(c))
importance(rf.LoanData)
varImpPlot(rf.LoanData)

#Boosting
#install.packages("gbm")
library(gbm)
set.seed(123457)
data_train$Funding_date <- as.factor(data_train$Funding_date)
data_val$Funding_date <- as.factor(data_val$Funding_date)
boost.LoanData=gbm(status~.,data=data_train,distribution="gaussian",n.trees=5000,interaction.depth=4)
summary(boost.LoanData)
yhat.boost=predict(boost.LoanData,newdata=data_val,n.trees=5000,type="response")
yhat.boost
yhat.test=data_val$status
yhat.test
yhat.val = ifelse(yhat.boost > 1, 1, 0)
(c = table(yhat.test,yhat.val))
(acc = (c[1,1]+c[2,2])/sum(c))
