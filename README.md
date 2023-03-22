# KPI-Reporting-R-automation
This project contains R scripts that automate the transposition of SQL files to an Excel file with 3 sheets named KPI_Report_ES_03. Additionally, the project generates an additional file named ES_KPI_Daily_Summary, which summarizes the information from KPI_Report_ES_03 on a daily basis. The Daily KPI Summary file provides an overview of the performance metrics for the day and serves as a quick reference for stakeholders.

## Table of Contents
* Description
* File Structure
* Requirements
* Usage
* License

## Description
The R scripts in this project execute SQL queries to retrieve data from a database and transpose the result data frames to Excel worksheets. One of the SQL queries used in the project, Withdrawal_Amounts_ES_PIVOT.sql, includes a pivot operation at the SQL level. This query aggregates transaction data for a certain period of time and creates a new table with summarized information pivoted on specific columns. The pivot operation transforms the data from a long format to a wide format, making it easier to summarize and analyze the data. The resulting pivoted data is then transposed to an Excel worksheet and saved in the KPI_Report_ES_03.xlsx file. In the resulting Excel file, the Withdrawal_Amounts_ES worksheet contains the transaction data summarized by transaction channel (Online, POS, and ATM) and type (count and amount) for a certain period of time. 

## Project Structure 
The project is organized as follows:

KPI-Reporting-R-automation

I. MS SQL Queries:
   1. ApprovedClients_ES.sql
 ActivatedClients_ES.sql
 Withdrawal_Amounts_ES_PIVOT.sql

II. R Scripts/
 function1.R
 function2.R
function3.R
 Master_SQL_Load.R

III. Output
KPI_Report_ES_03.xlsx

IV. 

V. README.md

