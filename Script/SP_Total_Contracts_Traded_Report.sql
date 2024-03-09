/****** Script for SelectTopNRows command from SSMS  ******/
Create Procedure SP_Total_Contracts_Traded_Report
(
@DateFrom datetime = null,
@DateTo datetime =null
)
AS
SELECT  [ExpiryDate]  [File Date], [Contract], [ContractsTraded] [Contracts Traded],Count([ContractsTraded])[% Of Total Contracts]

FROM [DevTest].[dbo].[DailyMTM]
WHERE  [ExpiryDate] between @DateFrom and @DateTo  
GROUP BY [ExpiryDate], [Contract],[ContractsTraded]
HAVING Count([ContractsTraded]) > 0


EXEC SP_Total_Contracts_Traded_Report '2024-03-09 17:00:00.000','2024-02-08 17:00:00.000'
