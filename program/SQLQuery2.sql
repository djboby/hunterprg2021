CREATE TABLE [dbo].[Teszt] (
    [Id]             INT            NOT NULL IDENTITY,
    [Helyrajzi szám] NVARCHAR (255) NULL,
    [Műv#ág]         NVARCHAR (255) NULL,
    PRIMARY KEY CLUSTERED ([Id] ASC)
);
INSERT INTO [dbo].[Teszt] ([Helyrajzi szám], [Műv#ág]) 
SELECT TOP 1000 [Helyrajzi szám] ,[Műv#ág] 
FROM [dbo].[Munka1$]
 WHERE [Helyrajzi szám] Like N'BALASSAGYARMAT/K/1%'
 AND [Műv#ág] IS NOT NULL

