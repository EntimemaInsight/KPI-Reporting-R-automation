declare @CurrentDate date = getdate();
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
