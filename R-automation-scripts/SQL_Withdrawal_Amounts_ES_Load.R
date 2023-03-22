library(odbc)
library(openxlsx)

# Set connection details
con <- DBI::dbConnect(odbc::odbc(),
                      Driver = "ODBC Driver 17 for SQL Server",
                      Server = "Scorpio",
                      Database = "BIsmartWCSP",
                      Trusted_Connection = "yes")

# Read SQL query from file
sql_file <- "C:/Users/aleksandar.dimitrov/Desktop/R Tests/SQL Queries/Withdrawal_Amounts_ES_PIVOT.sql"
sql_lines <- readLines(sql_file)
sql <- paste(sql_lines, collapse = " ")

# Set parameters for query
current_date <- Sys.Date()
first_of_month <- ifelse(day(current_date) == 1, as.Date(paste0(year(current_date), "-", month(current_date) - 2, "-01")), as.Date(paste0(year(current_date), "-", month(current_date) - 1, "-01")))

# Execute query and retrieve results
query <- paste0("declare @CurrentDate date = getdate();
declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

Select
    CONVERT(date, fct.CDate) AS Date,
    SUM(CASE WHEN tc.Name = 'ECOMM' THEN AmountTransactionBilling * -1 ELSE 0 END) AS OnlineAmount,
    SUM(CASE WHEN tc.Name = 'POS' THEN AmountTransactionBilling * -1 ELSE 0 END) AS POSAmount,
    SUM(CASE WHEN tc.Name = 'ATM' THEN AmountTransactionBilling * -1 ELSE 0 END) AS ATMAmount,
    COUNT(CASE WHEN tc.Name = 'ECOMM' THEN AmountTransactionBilling ELSE NULL END) AS OnlineCount,
    COUNT(CASE WHEN tc.Name = 'POS' THEN AmountTransactionBilling ELSE NULL END) AS POSCount,
    COUNT(CASE WHEN tc.Name = 'ATM' THEN AmountTransactionBilling ELSE NULL END) AS ATMCount
from dwh.FactCardTransactions fct
JOIN dwh.DimCardsHist ch ON ch.CardSYSID = fct.CardSYSID
JOIN dwh.DimOfferHist oh ON oh.OfferSYSID = ch.OfferSYSID
JOIN dwh.DimProduct p ON p.ProductSK = oh.ProductSK
JOIN dwh.DimTransactionChannels tc ON tc.TransactionChannelSK = fct.TransactionChannelSK
WHERE CONVERT(DATE, fct.CDate, 4) >= @FirstOfMonth
    AND CONVERT(DATE, fct.CDate, 4) <= @CurrentDate
    AND tc.Name NOT IN ('BALANCE', 'PIN_CHANGE') 
    AND AmountTransactionBilling < 0 
    AND ResponseCodeSK = 1
    AND ch.Latest = 1
    AND oh.Latest = 1
GROUP BY CONVERT(date, fct.CDate)
ORDER BY CONVERT(date, fct.CDate)
")
result <- DBI::dbGetQuery(con, query)

# Transpose the result data frame
result_transposed <- as.data.frame(t(result))
names(result_transposed) <- result_transposed[1,]
result_transposed <- result_transposed[-1,]

# Create new workbook and worksheet
wb <- createWorkbook()
addWorksheet(wb, "Withdrawal_Amounts_ES")

# Write titles for the data in column A
writeData(wb, "Withdrawal_Amounts_ES", "Dates", startCol = 1, startRow = 2)
writeData(wb, "Withdrawal_Amounts_ES", "Online (sum) ", startCol = 1, startRow = 3)
writeData(wb, "Withdrawal_Amounts_ES", "POS (sum)", startCol = 1, startRow = 4)
writeData(wb, "Withdrawal_Amounts_ES", "ATM (sum)", startCol = 1, startRow = 5)
writeData(wb, "Withdrawal_Amounts_ES", "Online (transactions)", startCol = 1, startRow = 6)
writeData(wb, "Withdrawal_Amounts_ES", "POS (transactions)", startCol = 1, startRow = 7)
writeData(wb, "Withdrawal_Amounts_ES", "ATM (transactions)", startCol = 1, startRow = 8)






# Write the transposed data to the worksheet starting at column B
writeData(wb, "Withdrawal_Amounts_ES", result_transposed, startCol = 2, startRow = 2)

# Save workbook to file
saveWorkbook(wb, "KPI_Report_ES_03.xlsx", overwrite = TRUE) 

# Close database connection
DBI::dbDisconnect(con) 
