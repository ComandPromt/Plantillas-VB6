DECLARE @CustomerID char(5)
SELECT @CustomerID = N'ALFKI'
SELECT * FROM [Northwind].[dbo].[GetCustomerContactView](@CustomerID)