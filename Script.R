#install and load package arules
#install.packages("arules")
library(arules)
#install and load arulesViz
#install.packages("arulesViz")
library(arulesViz)
#install and load tidyverse
#install.packages("tidyverse")
library(tidyverse)
#install and load readxml
#install.packages("readxml")
library(readxl)
#install and load knitr
#install.packages("knitr")
library(knitr)
#load ggplot2 as it comes in tidyverse
library(ggplot2)
#install and load lubridate
#install.packages("lubridate")
library(lubridate)
#install and load plyr
#install.packages("plyr")
library(plyr)
library(dplyr)
library("openxlsx")

# Clear Workspace
rm(list = ls())
gc()
# Read xlsx file
#dataspread1 <- read.xlsx("d1.xlsx", colNames = TRUE, sheet = 1)
#dataspread2 <- read.xlsx("d2.xlsx", colNames = TRUE, sheet = 1)
#dataspread3 <- read.xlsx("d3.xlsx", colNames = TRUE, sheet = 1)
#dataspread4 <- read.xlsx("d4.xlsx", colNames = TRUE, sheet = 1)
#dataspread5 <- read.xlsx("d5.xlsx", colNames = TRUE, sheet = 1)
#dataspread6 <- read.xlsx("d6.xlsx", colNames = TRUE, sheet = 1)
# prodList <- read.xlsx("ProductList.xlsx", colNames = TRUE, sheet = 1)

# Combine Lists
#data <- unique(rbind(dataspread1,
          #dataspread2,
          #dataspread3,
          #dataspread4,
          #dataspread5,
          #dataspread6))

#if (!file.exists("dataSQL.rds")) {
    #dataCollect <- dbGetQuery(con, "SELECT DISTINCT DW1.dbo.Vw_Bz_Sale.FactorID, DW1.dbo.Vw_Bz_Sale.Date, DW1.dbo.Vw_Bz_Sale.RegionID, DW1.dbo.Vw_Bz_SupplierProduct.SubFamilyId,[DW1].[dbo].[Vw_Bz_SupplierProduct].SubFamilyName FROM DW1.dbo.Vw_Bz_Sale with (nolock) INNER JOIN [dbo].[Vw_Bz_SupplierProduct] with (nolock) ON DW1.dbo.Vw_Bz_Sale.ProductID = DW1.dbo.Vw_Bz_SupplierProduct.ProductId WHERE (DW1.dbo.Vw_Bz_SupplierProduct.SectionId !='1691.199') AND (DW1.dbo.Vw_Bz_Sale.RegionID = '1.199') AND ((DW1.dbo.Vw_Bz_Sale.Date BETWEEN '970701' AND '970703') OR (DW1.dbo.Vw_Bz_Sale.Date BETWEEN '970707' AND '970711') OR (DW1.dbo.Vw_Bz_Sale.Date BETWEEN '970714' AND '970718') OR (DW1.dbo.Vw_Bz_Sale.Date BETWEEN '970721' AND '970725') OR (DW1.dbo.Vw_Bz_Sale.Date BETWEEN '970728' AND '970730'))")
    #dataCollect$FactorID <- as.character(dataCollect$FactorID)
    #dataCollect$RegionID <- as.character(dataCollect$RegionID)
    #dataCollect$SubFamilyId <- as.character(dataCollect$SubFamilyId)
    #saveRDS(dataCollect, "dataSQL.rds")
#} else {
    dataCollect <- readRDS("dataSQL.rds")
#}

if (!file.exists("Sample1000000int.rds")) {
    sample1000000int <- sample.int(length(unique(dataCollect$FactorID)), 1000000)
    saveRDS(sample1000000int, file = "Sample1000000int.rds")
}
sample1000000int <- readRDS("Sample1000000int.rds")

# Factor ID
uniqueFactorId <- unique(dataCollect$FactorID)

# Get Sample from Factor ID
sampleFactorID <- uniqueFactorId[sample1000000int]

# Get All Sample Data
sampleData <- dataCollect[which(dataCollect$FactorID %in% sampleFactorID),]

# Clear Data
#rm(dataspread1,
          #dataspread2,
          #dataspread3,
          #dataspread4,
          #dataspread5,
          #dataspread6)
#gc()

#transactionData <- ddply(dataCollect, c("FactorID", "Date"),
                       #function(df1) paste(df1$SubFamilyName,
                       #collapse = ","))

#set column InvoiceNo of dataframe transactionData  
#transactionData$ReceiptId <- NULL
##set column Date of dataframe transactionData
#transactionData$Date <- NULL
#Rename column to items
#colnames(transactionData) <- c("items")
#Show Dataframe transactionData
#con <- file("market_basket_transactions.csv", encoding = "UTF-8")
#write.csv(transactionData, file = con, quote = FALSE, row.names = TRUE)
#tr <- read.transactions('market_basket_transactions.csv', format = 'basket', sep = ',')
#sampleData <- dataCollect
tmpFac <- sampleData[, c("FactorID", "SubFamilyName")]
rm(dataCollect)
gc()
tmpFac$FactorID <- as.factor(tmpFac$FactorID)
tmpFac$SubFamilyName <- as.factor(tmpFac$SubFamilyName)
tr <- as(split(tmpFac[, "SubFamilyName"], tmpFac[, "FactorID"]), "transactions")
rm(tmpFac)
gc()
# Create an item frequency plot for the top 20 items
if (!require("RColorBrewer")) {
    # install color package of R
    install.packages("RColorBrewer")
    #include library RColorBrewer
    library(RColorBrewer)
}
itemFrequencyPlot(tr, topN = 20, type = "absolute", col = brewer.pal(8, 'Pastel2'), main = "Absolute Item Frequency Plot")
itemFrequencyPlot(tr, topN = 20, type = "relative", col = brewer.pal(8, 'Pastel2'), main = "Relative Item Frequency Plot")
# Min Support as 0.001, confidence as 0.8.
association.rules <- apriori(tr, parameter = list(supp = 0.01, conf = 0.4, maxlen = 10))

inspect(association.rules[1:10])

shorter.association.rules <- apriori(tr, parameter = list(supp = 0.01, conf = 0.9, maxlen = 2))

subset.rules <- which(colSums(is.subset(association.rules, association.rules)) > 1) # get subset rules in vector

subset.association.rules <- association.rules[-subset.rules] # remove subset rules.


# Filter rules with confidence greater than 0.4 or 40%
subRules <- association.rules[quality(association.rules)$confidence > 0.4]
#Plot SubRules
plot(subRules)


finalRes <- cbind(as(subRules, "data.frame"), conviction = interestMeasure(subRules, "conviction", tr))
write.xlsx(finalRes, "rules.xlsx", encoding = "UTF-8")
top20subRules <- head(subRules, n = 100, by = "support")
plot(subRules, method = "two-key plot")

top10subRules <- head(subRules, n = 10, by = "confidence")

plot(top10subRules, method = "graph", engine = "htmlwidget")

saveAsGraph(head(subRules, n = 1000, by = "lift"), file = "rules.graphml")

# Filter top 20 rules with highest lift
subRules2 <- head(subRules, n = 20, by = "lift")
plot(subRules2, method = "paracoord")