CREATE TABLE [dbo].[Subscriptions] (
    [SubscriptionId]               NVARCHAR (128)  NOT NULL,
    [Watermark]        NVARCHAR (2048) NULL,
    [PreviousWatermark]        NVARCHAR (2048) NULL,
    [SmtpAddress]      NVARCHAR (255)  NULL,
    [LastRunTime]      DATETIME        NOT NULL,
    [SubscriptionType] INT             NOT NULL,
    [Terminated]       BIT             NOT NULL,
    CONSTRAINT [PK_dbo.Subscriptions] PRIMARY KEY CLUSTERED ([SubscriptionId] ASC)
);

