Attribute VB_Name = "modQueueObj"
' modQueueObj.mod
' Copyright (C) 2008 Eric Evans

Option Explicit

Public Enum PRIORITY
    SPECIAL_MESSAGE = 0
    CONSOLE_MESSAGE = 1
    CHANNEL_MODERATION_MESSAGE = 2
    COMMAND_RESPONSE_MESSAGE = 3
    MESSAGE_DEFAULT = 100
End Enum
