﻿CREATE TABLE [dbo].[Subscriptions] (
    [Id] INT NOT NULL IDENTITY, 
    [SmtpAddress]      NVARCHAR (255)  NULL,
    [Watermark]        NVARCHAR (2048) NULL,
    [PreviousWatermark]        NVARCHAR (2048) NULL,
    [LastRunTime]      DATETIME        NOT NULL,
    CONSTRAINT [PK_Subscriptions] PRIMARY KEY CLUSTERED ([Id])
);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_SubscriptionType] ON [dbo].[Subscriptions]
(
	[SmtpAddress] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

