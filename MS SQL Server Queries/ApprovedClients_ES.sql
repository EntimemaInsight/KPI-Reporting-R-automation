declare @CurrentDate date = getdate();
declare @FirstOfMonth date = case when day(getdate())=1 then DATEADD(DAY,1,EOMONTH(@CurrentDate,-2)) else DATEADD(DAY,1,EOMONTH(@CurrentDate,-1)) end; 

select
CONVERT(DATE, DateApproved, 4) AS Date,
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
group by p.Name,
CONVERT(DATE, DateApproved, 4)










