declare @CurrentDate date = getdate();
declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

Select
    CONVERT(date, fct.CDate) AS Date,
    Case
	   When tc.Name = 'ECOMM'
            Then 'ONLINE'
        ELSE tc.Name
    END As TransactionChanel,
	SUM(AmountTransactionBilling * -1) AS Amount,
    COUNT(AmountTransactionBilling) AS CountTransaction 
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
GROUP BY CONVERT(date, fct.CDate), tc.Name
order by CONVERT(date, fct.CDate)

