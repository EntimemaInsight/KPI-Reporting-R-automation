library(odbc)
library(openxlsx)


# Set connection details
con <- DBI::dbConnect(odbc::odbc(),
                      Driver = "ODBC Driver 17 for SQL Server",
                      Server = "Scorpio",
                      Database = "BIsmartWCSP",
                      Trusted_Connection = "yes")



# Read SQL query from file
sql_file <- "C:/Users/aleksandar.dimitrov/Desktop/R Tests/SQL Queries/ApprovedClients_ES.sql"
sql_lines <- readLines(sql_file)
sql <- paste(sql_lines, collapse = " ")

# Set parameters for query
current_date <- Sys.Date()
first_of_month <- ifelse(day(current_date) == 1, as.Date(paste0(year(current_date), "-", month(current_date) - 2, "-01")), as.Date(paste0(year(current_date), "-", month(current_date) - 1, "-01")))

# Execute query and retrieve results
query <- paste0("declare @CurrentDate date = getdate();
                 declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

                 select CONVERT(DATE, DateApproved, 4) AS Date,
                        SUM(cpsc.Limit) as Limit,
                        Count(cpsc.Limit) as Count
                 from dwh.DimOfferHist oh
                 JOIN dwh.DimProduct p ON p.ProductSK = oh.ProductSK
                 Join dwhsc.DimCreditProposalSC cpsc ON cpsc.OfferSYSID = oh.OfferSYSID
                 where CONVERT(DATE, DateApproved, 4) >= @FirstOfMonth
                       AND CONVERT(DATE, DateApproved, 4) <= @CurrentDate
                       AND (DateRefused < DateApproved)
                       AND (DateRejected < DateApproved)
                       and Latest = 1
                 group by p.Name, CONVERT(DATE, DateApproved, 4)")
result <- DBI::dbGetQuery(con, query)

# Transpose the result data frame
result_transposed <- t(result)

# Create new workbook and worksheet
wb <- createWorkbook()
addWorksheet(wb, "ApprovedClients_ES")
addWorksheet(wb, "ActivatedClients_ES")
addWorksheet(wb, "Withdrawal_Amounts_ES")

# Write titles for the data in column A
writeData(wb, "ApprovedClients_ES", "Dates", startCol = 1, startRow = 3)
writeData(wb, "ApprovedClients_ES", "Limits", startCol = 1, startRow = 4)
writeData(wb, "ApprovedClients_ES", "Counts", startCol = 1, startRow = 5)

# Write the transposed data to the worksheet starting at column B
writeData(wb, "ApprovedClients_ES", result_transposed, startCol = 2, startRow = 2)

# Add number format to the cells in the "Limits" and "Count" rows
limit_style <- createStyle(numFmt = "#,##0.00")
addStyle(wb, sheet = "ApprovedClients_ES", limit_style, rows = 4, cols = 2:ncol(result_transposed))
count_style <- createStyle(numFmt = "#,##0")
addStyle(wb, sheet = "ApprovedClients_ES", count_style, rows = 5, cols = 2:ncol(result_transposed))




# Save workbook to file
saveWorkbook(wb, "KPI_Report_ES_03.xlsx", overwrite = TRUE) 

# Close database connection
DBI::dbDisconnect(con) 




# Set connection details
con <- DBI::dbConnect(odbc::odbc(),
                      Driver = "ODBC Driver 17 for SQL Server",
                      Server = "Scorpio",
                      Database = "BIsmartWCSP",
                      Trusted_Connection = "yes")

# Read SQL query from file
sql_file <- "C:/Users/aleksandar.dimitrov/Desktop/R Tests/SQL Queries/ActivatedClients_ES.sql"
sql_lines <- readLines(sql_file)
sql <- paste(sql_lines, collapse = " ")

# Set parameters for query
current_date <- Sys.Date()
first_of_month <- ifelse(day(current_date) == 1, as.Date(paste0(year(current_date), "-", month(current_date) - 2, "-01")), as.Date(paste0(year(current_date), "-", month(current_date) - 1, "-01")))

# Execute query and retrieve results
query <- paste0("declare @CurrentDate date = getdate();
declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

SELECT 
    FORMAT(ActivationDate, 'dd.MM.yyyy') AS Date,
    COUNT(p.Name) AS ActivatedClients,
    SUM(ch.Limit) AS ActivatedLimit
FROM dwh.DimCardsHist ch
JOIN dwh.DimOfferHist oh ON oh.OfferSYSID = ch.OfferSYSID
JOIN dwh.DimProduct p ON oh.ProductSK = p.ProductSK
WHERE
    CONVERT(DATE, ActivationDate, 4) >= @FirstOfMonth AND
    CONVERT(DATE, ActivationDate, 4) <= @CurrentDate AND
    ch.Latest = 1 AND
    oh.Latest = 1
GROUP BY FORMAT(ActivationDate, 'dd.MM.yyyy')")
result <- DBI::dbGetQuery(con, query)

# Transpose the result data frame
result_transposed <- t(result)


# Write titles for the data in column A
writeData(wb, "ActivatedClients_ES", "Dates", startCol = 1, startRow = 3)
writeData(wb, "ActivatedClients_ES", "ActivatedClients", startCol = 1, startRow = 4)
writeData(wb, "ActivatedClients_ES", "ActivatedLimit", startCol = 1, startRow = 5)

# Write the transposed data to the worksheet starting at column B
writeData(wb, "ActivatedClients_ES", result_transposed, startCol = 2, startRow = 2)

# Add number format to the cells in the "ActivatedLimit" rows
limit_style <- createStyle(numFmt = "#,##0.00")
addStyle(wb, sheet = "ActivatedClients_ES", limit_style, rows = 4, cols = 2:ncol(result_transposed))
addStyle(wb, sheet = "ActivatedClients_ES", limit_style, rows = 5, cols = 2:ncol(result_transposed))

# Save workbook to file
saveWorkbook(wb, "KPI_Report_ES_03.xlsx", overwrite = TRUE) 

# Close database connection
DBI::dbDisconnect(con)




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

# Write titles for the data in column A
writeData(wb, "Withdrawal_Amounts_ES", "Dates", startCol = 1, startRow = 2)
writeData(wb, "Withdrawal_Amounts_ES", "OnlineAmount", startCol = 1, startRow = 3)
writeData(wb, "Withdrawal_Amounts_ES", "POSAmount", startCol = 1, startRow = 4)
writeData(wb, "Withdrawal_Amounts_ES", "ATMAmount", startCol = 1, startRow = 5)
writeData(wb, "Withdrawal_Amounts_ES", "OnlineCount", startCol = 1, startRow = 6)
writeData(wb, "Withdrawal_Amounts_ES", "POSCount", startCol = 1, startRow = 7)
writeData(wb, "Withdrawal_Amounts_ES", "ATMCount", startCol = 1, startRow = 8)






# Write the transposed data to the worksheet starting at column B
writeData(wb, "Withdrawal_Amounts_ES", result_transposed, startCol = 2, startRow = 2)

# Save workbook to file
saveWorkbook(wb, "KPI_Report_ES_03.xlsx", overwrite = TRUE) 

# Close database connection
DBI::dbDisconnect(con) 




