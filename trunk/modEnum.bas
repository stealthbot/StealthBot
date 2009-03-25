Attribute VB_Name = "modEnum"
Option Explicit

'modEnum - project StealthBot
'February 2004 - Stealth [stealth at stealthbot dot net]


'UDTS
Public Type udtChanList
    Current         As String
    flags           As Long
    Designated      As String
    staticDesignee  As String
End Type

Public Type udtOutFilters
    ofFind      As String
    ofReplace   As String
End Type

Public Type udtBanList
    Username        As String
    UsernameActual  As String
    cOperator       As String
End Type

Public Type udtAutoRespond
    Check As String
    Reply As String
End Type

Public Type udtCustomCommandData
    reqAccess   As Integer
    Query       As String * 20
    Action      As String * 500
End Type

Public Type udtDatabase
    Username    As String
    Access      As Integer
    flags       As String
    AddedBy     As String
    AddedOn     As Date
    ModifiedBy  As String
    ModifiedOn  As Date
    Type        As String
    Groups      As String
    BanMessage  As String
End Type

'Public Type udtQueue
'    Message     As String
'    Priority    As Byte
'End Type

Public Type udtGetAccessResponse
    Username    As String
    Access      As Integer
    flags       As String
    AddedBy     As String
    AddedOn     As Date
    ModifiedBy  As String
    ModifiedOn  As Date
    Type        As String
    Groups      As String
    BanMessage  As String
End Type

Public Type udtMail
    To      As String * 30
    From    As String * 30
    Message As String * 225
End Type
   
Public Type FLASHWINFO
    cbSize      As Long
    hWnd        As Long
    dwFlags     As Long
    uCount      As Long
    dwTimeout   As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Public Enum inetQueueModes
    inqReset = 0
    inqAdd = 1
    inqGet = 2
End Enum

Public Enum enuErrorSources
    BNET = 0
    BNLS = 1
    MCP = 2
End Enum

Public Enum enuProxyStatus
    psNotConnected = 0
    psConnecting = 1
    psLoggingIn = 2
    psOnline = 3
End Enum

Public Enum enuWebProfileTypes
    W3XP = 1
    WAR3 = 2
End Enum

Public Enum eCharacterTypes
    Unknown& = &H0
    Amazon& = &H1
    Sorceress& = &H2
    Necromancer& = &H3
    Paladin& = &H4
    Barbarian& = &H5
    Druid& = &H6
    Assassin& = &H7
End Enum

Public Enum enuSettingsPanels
    spConnectionConfig = 0
    spConnectionAdvanced = 1
    spInterfaceGeneral = 2
    spInterfaceFontsColors = 3
    '#4 IS SKIPPED
    spGenModeration = 5
    spGenGreets = 6
    spGenIdles = 7
    spGenMisc = 8
    spSplash = 9
End Enum

Public Enum enuDBActions
    AddEntry = 1
    RemEntry = 2
    ModEntry = 3
End Enum

Public Enum enuPL_ServerTypes
    stBNLS = 1
    stBNCS = 2
    stMCP = 3
End Enum

Public Enum enuPL_DirectionTypes
    CtoS = 1
    StoC = 2
End Enum

