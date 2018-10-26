CREATE TABLE [dbo].[RoomListRoomsSyncState]
(
	[Id] INT NOT NULL IDENTITY, 
    [RoomId] INT NOT NULL, 
    [SyncState] NVARCHAR(512) NULL, 
    [SyncTimestamp] DATETIME NULL,
    CONSTRAINT [PK_RoomListRoomsSyncState_Id] PRIMARY KEY CLUSTERED ([Id] ASC),
    CONSTRAINT [FK_RoomListRoomsSyncState_RoomListRooms_RoomId] FOREIGN KEY ([RoomId]) REFERENCES [dbo].[RoomListRooms] ([Id]) ON DELETE CASCADE
)
