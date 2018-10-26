CREATE TABLE [dbo].[RoomAppointment] (
    [Id]                   INT             IDENTITY (1, 1) NOT NULL,
    [OrganizerSmtpAddress] NVARCHAR (MAX)  NULL,
    [RoomId]               INT             NOT NULL,
    [StartUTC]             DATETIME        NOT NULL,
    [EndUTC]               DATETIME        NOT NULL,
    [Subject]              NVARCHAR (256)  NULL,
    [Location]             NVARCHAR (1024) NULL,
    [RecurrencePattern]    NVARCHAR (MAX)  NULL,
    [IsRecurringMeeting]   BIT             NOT NULL,
    [BookingId]  NVARCHAR (2048)  NULL,
    [BookingChangeKey]     NVARCHAR (2048)   NULL,
	[BookingReference]  NVARCHAR(50) NULL,
    [ExistsInExchange]     BIT             NOT NULL,
    [SyncedWithExchange] BIT NOT NULL DEFAULT 0, 
    [IsDeleted] BIT NOT NULL DEFAULT 0, 
    [DeletedLocally] BIT NOT NULL DEFAULT 0, 
    [ModifiedDate] DATETIME NULL, 
    CONSTRAINT [PK_RoomAppointment_Id] PRIMARY KEY CLUSTERED ([Id] ASC),
    CONSTRAINT [FK_RoomAppointment_RoomListRooms_RoomId] FOREIGN KEY ([RoomId]) REFERENCES [dbo].[RoomListRooms] ([Id]) ON DELETE CASCADE
);


GO
CREATE NONCLUSTERED INDEX [IX_RoomId]
    ON [dbo].[RoomAppointment]([RoomId] ASC);

