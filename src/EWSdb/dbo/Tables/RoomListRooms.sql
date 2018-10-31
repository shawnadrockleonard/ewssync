CREATE TABLE [dbo].[RoomListRooms] (
    [Id]           INT            IDENTITY (1, 1) NOT NULL,
    [Identity]     NVARCHAR (512) NULL,
    [RoomList]     NVARCHAR (155) NULL,
    [SmtpAddress]  NVARCHAR (155) NULL,
    [LastSyncDate] DATETIME       NULL,
    [KnownEvents]  INT            NULL,
    [SyncState] NVARCHAR(512) NULL, 
    [SyncTimestamp] DATETIME NULL, 
    CONSTRAINT [PK_dbo.RoomListRooms] PRIMARY KEY CLUSTERED ([Id] ASC)
);

