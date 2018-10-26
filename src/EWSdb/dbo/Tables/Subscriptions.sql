CREATE TABLE [dbo].[Subscriptions] (
    [Id] INT NOT NULL IDENTITY, 
    [SubscriptionId]               NVARCHAR (128)  NOT NULL,
    [Watermark]        NVARCHAR (2048) NULL,
    [PreviousWatermark]        NVARCHAR (2048) NULL,
    [SmtpAddress]      NVARCHAR (255)  NULL,
    [LastRunTime]      DATETIME        NOT NULL,
    [SubscriptionType] INT             NOT NULL,
    [Terminated]       BIT             NOT NULL,
    CONSTRAINT [PK_Subscriptions] PRIMARY KEY CLUSTERED ([Id])
);
GO
CREATE UNIQUE NONCLUSTERED INDEX [IX_SubscriptionId] ON [dbo].[Subscriptions]([SubscriptionId] ASC);

