declare @CurrentDate date = getdate();
declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

SELECT 
FORMAT(ActivationDate, 'dd.MM.yyyy') AS Date
,COUNT(p.Name) AS ActivatedClients
,SUM(ch.Limit) AS ActivatedLimit
  FROM dwh.DimCardsHist ch
  JOIN dwh.DimOfferHist oh ON oh.OfferSYSID = ch.OfferSYSID
  JOIN dwh.DimProduct p ON oh.ProductSK = p.ProductSK
  WHERE
    	  	   CONVERT(DATE, ActivationDate, 4) >= @FirstOfMonth 
	       AND CONVERT(DATE, ActivationDate, 4) <= @CurrentDate
		   AND ch.Latest = 1
		   AND oh.Latest = 1
  GROUP BY FORMAT(ActivationDate, 'dd.MM.yyyy')