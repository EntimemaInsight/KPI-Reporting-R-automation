library(odbc)
library(openxlsx)

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

# Create new workbook and worksheet
wb <- createWorkbook()
addWorksheet(wb, "ActivatedClients_ES")

# Write titles for the data in column A
writeData(wb, "ActivatedClients_ES", "Dates", startCol = 1, startRow = 3)
writeData(wb, "ActivatedClients_ES", "Activated", startCol = 1, startRow = 4)
writeData(wb, "ActivatedClients_ES", "Activated limit", startCol = 1, startRow = 5)

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
