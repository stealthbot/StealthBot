VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_VERSION         As Integer = 3

Private Const SECTION_CHANNELLIST   As String = "ChannelList"
Private Const SECTION_CHATUI        As String = "ChatUI"

Private m_ColorFilePath As String

' Channel list
Private m_ChannelLabelBack As Long
Private m_ChannelLabelText As Long
Private m_ChannelListBack As Long
Private m_ChannelListText As Long
Private m_ChannelListSelf As Long
Private m_ChannelListIdle As Long
Private m_ChannelListSquelched As Long
Private m_ChannelListOps As Long
Private m_ChannelListAdmin As Long

' Chat RTB
Private m_RTBBack As Long
Private m_SendBoxesBack As Long
Private m_SendBoxesText As Long
Private m_TalkBotUsername As Long
Private m_TalkUsernameNormal As Long
Private m_TalkUsernameOp As Long
Private m_TalkUsernameAdmin As Long
Private m_TalkNormalText As Long
Private m_Carats As Long
Private m_InformationText As Long
Private m_SuccessText As Long
Private m_ErrorMessageText As Long
Private m_TimeStamps As Long
Private m_ServerInfoText As Long
Private m_ConsoleText As Long
Private m_WhisperCarats As Long
Private m_WhisperUsernames As Long
Private m_WhisperText As Long
Private m_EmoteUsernames As Long
Private m_EmoteText As Long
Private m_JoinUsername As Long
Private m_JoinText As Long
Private m_JoinedChannelName As Long
Private m_JoinedChannelText As Long

Sub Class_Initialize()
    m_ColorFilePath = vbNullString
    Call LoadDefaults
End Sub

' Static colors
Public Property Get AliceBlue() As Long
    AliceBlue = 16775408
End Property

Public Property Get AntiqueWhite() As Long
    AntiqueWhite = 14150650
End Property

Public Property Get Aqua() As Long
    Aqua = 16776960
End Property

Public Property Get Aquamarine() As Long
    Aquamarine = 13959039
End Property

Public Property Get Azure() As Long
    Azure = 16777200
End Property

Public Property Get Beige() As Long
    Beige = 14480885
End Property

Public Property Get Bisque() As Long
    Bisque = 12903679
End Property

Public Property Get Black() As Long
    Black = 0
End Property

Public Property Get BlanchedAlmond() As Long
    BlanchedAlmond = 13495295
End Property

Public Property Get Blue() As Long
    Blue = 16711680
End Property

Public Property Get BlueViolet() As Long
    BlueViolet = 14822282
End Property

Public Property Get Brown() As Long
    Brown = 2763429
End Property

Public Property Get BurlyWood() As Long
    BurlyWood = 8894686
End Property

Public Property Get CadetBlue() As Long
    CadetBlue = 10526303
End Property

Public Property Get Chartreuse() As Long
    Chartreuse = 65407
End Property

Public Property Get Chocolate() As Long
    Chocolate = 1993170
End Property

Public Property Get Coral() As Long
    Coral = 5275647
End Property

Public Property Get CornflowerBlue() As Long
    CornflowerBlue = 15570276
End Property

Public Property Get Cornsilk() As Long
    Cornsilk = 14481663
End Property

Public Property Get Crimson() As Long
    Crimson = 3937500
End Property

Public Property Get Cyan() As Long
    Cyan = 16776960
End Property

Public Property Get DarkBlue() As Long
    DarkBlue = 9109504
End Property

Public Property Get DarkCyan() As Long
    DarkCyan = 9145088
End Property

Public Property Get DarkGoldenRod() As Long
    DarkGoldenRod = 755384
End Property

Public Property Get DarkGray() As Long
    DarkGray = 11119017
End Property

Public Property Get DarkGreen() As Long
    DarkGreen = 25600
End Property

Public Property Get DarkKhaki() As Long
    DarkKhaki = 7059389
End Property

Public Property Get DarkMagenta() As Long
    DarkMagenta = 9109643
End Property

Public Property Get DarkOliveGreen() As Long
    DarkOliveGreen = 3107669
End Property

Public Property Get Darkorange() As Long
    Darkorange = 36095
End Property

Public Property Get DarkOrchid() As Long
    DarkOrchid = 13382297
End Property

Public Property Get DarkRed() As Long
    DarkRed = 139
End Property

Public Property Get DarkSalmon() As Long
    DarkSalmon = 8034025
End Property

Public Property Get DarkSeaGreen() As Long
    DarkSeaGreen = 9419919
End Property

Public Property Get DarkSlateBlue() As Long
    DarkSlateBlue = 9125192
End Property

Public Property Get DarkSlateGray() As Long
    DarkSlateGray = 5197615
End Property

Public Property Get DarkTurquoise() As Long
    DarkTurquoise = 13749760
End Property

Public Property Get DarkViolet() As Long
    DarkViolet = 13828244
End Property

Public Property Get DeepPink() As Long
    DeepPink = 9639167
End Property

Public Property Get DeepSkyBlue() As Long
    DeepSkyBlue = 16760576
End Property

Public Property Get DimGray() As Long
    DimGray = 6908265
End Property

Public Property Get DodgerBlue() As Long
    DodgerBlue = 16748574
End Property

Public Property Get FireBrick() As Long
    FireBrick = 2237106
End Property

Public Property Get FloralWhite() As Long
    FloralWhite = 15792895
End Property

Public Property Get ForestGreen() As Long
    ForestGreen = 2263842
End Property

Public Property Get Fuchsia() As Long
    Fuchsia = 16711935
End Property

Public Property Get Gainsboro() As Long
    Gainsboro = 14474460
End Property

Public Property Get GhostWhite() As Long
    GhostWhite = 16775416
End Property

Public Property Get Gold() As Long
    Gold = 55295
End Property

Public Property Get GoldenRod() As Long
    GoldenRod = 2139610
End Property

Public Property Get Gray() As Long
    Gray = 8421504
End Property

Public Property Get Green() As Long
    Green = 32768
End Property

Public Property Get GreenYellow() As Long
    GreenYellow = 3145645
End Property

Public Property Get HoneyDew() As Long
    HoneyDew = 15794160
End Property

Public Property Get HotPink() As Long
    HotPink = 11823615
End Property

Public Property Get IndianRed() As Long
    IndianRed = 6053069
End Property

Public Property Get Indigo() As Long
    Indigo = 8519755
End Property

Public Property Get Ivory() As Long
    Ivory = 15794175
End Property

Public Property Get Khaki() As Long
    Khaki = 9234160
End Property

Public Property Get Lavender() As Long
    Lavender = 16443110
End Property

Public Property Get LavenderBlush() As Long
    LavenderBlush = 16118015
End Property

Public Property Get LawnGreen() As Long
    LawnGreen = 64636
End Property

Public Property Get LemonChiffon() As Long
    LemonChiffon = 13499135
End Property

Public Property Get LightBlue() As Long
    LightBlue = 15128749
End Property

Public Property Get LightCoral() As Long
    LightCoral = 8421616
End Property

Public Property Get LightCyan() As Long
    LightCyan = 16777184
End Property

Public Property Get LightGoldenRodYellow() As Long
    LightGoldenRodYellow = 13826810
End Property

Public Property Get LightGray() As Long
    LightGray = 13882323
End Property

Public Property Get LightGreen() As Long
    LightGreen = 9498256
End Property

Public Property Get LightPink() As Long
    LightPink = 12695295
End Property

Public Property Get LightSalmon() As Long
    LightSalmon = 8036607
End Property

Public Property Get LightSeaGreen() As Long
    LightSeaGreen = 11186720
End Property

Public Property Get LightSkyBlue() As Long
    LightSkyBlue = 16436871
End Property

Public Property Get LightSlateGray() As Long
    LightSlateGray = 10061943
End Property

Public Property Get LightSteelBlue() As Long
    LightSteelBlue = 14599344
End Property

Public Property Get LightYellow() As Long
    LightYellow = 14745599
End Property

Public Property Get Lime() As Long
    Lime = 65280
End Property

Public Property Get LimeGreen() As Long
    LimeGreen = 3329330
End Property

Public Property Get Linen() As Long
    Linen = 15134970
End Property

Public Property Get Magenta() As Long
    Magenta = 16711935
End Property

Public Property Get Maroon() As Long
    Maroon = 128
End Property

Public Property Get MediumAquaMarine() As Long
    MediumAquaMarine = 11193702
End Property

Public Property Get MediumBlue() As Long
    MediumBlue = 13434880
End Property

Public Property Get MediumOrchid() As Long
    MediumOrchid = 13850042
End Property

Public Property Get MediumPurple() As Long
    MediumPurple = 14184595
End Property

Public Property Get MediumSeaGreen() As Long
    MediumSeaGreen = 7451452
End Property

Public Property Get MediumSlateBlue() As Long
    MediumSlateBlue = 15624315
End Property

Public Property Get MediumSpringGreen() As Long
    MediumSpringGreen = 10156544
End Property

Public Property Get MediumTurquoise() As Long
    MediumTurquoise = 13422920
End Property

Public Property Get MediumVioletRed() As Long
    MediumVioletRed = 8721863
End Property

Public Property Get MidnightBlue() As Long
    MidnightBlue = 7346457
End Property

Public Property Get MintCream() As Long
    MintCream = 16449525
End Property

Public Property Get MistyRose() As Long
    MistyRose = 14804223
End Property

Public Property Get Moccasin() As Long
    Moccasin = 11920639
End Property

Public Property Get NavajoWhite() As Long
    NavajoWhite = 11394815
End Property

Public Property Get Navy() As Long
    Navy = 8388608
End Property

Public Property Get OldLace() As Long
    OldLace = 15136253
End Property

Public Property Get Olive() As Long
    Olive = 32896
End Property

Public Property Get OliveDrab() As Long
    OliveDrab = 2330219
End Property

Public Property Get Orange() As Long
    Orange = 42495
End Property

Public Property Get OrangeRed() As Long
    OrangeRed = 17919
End Property

Public Property Get Orchid() As Long
    Orchid = 14053594
End Property

Public Property Get PaleGoldenRod() As Long
    PaleGoldenRod = 11200750
End Property

Public Property Get PaleGreen() As Long
    PaleGreen = 10025880
End Property

Public Property Get PaleTurquoise() As Long
    PaleTurquoise = 15658671
End Property

Public Property Get PaleVioletRed() As Long
    PaleVioletRed = 9662680
End Property

Public Property Get PapayaWhip() As Long
    PapayaWhip = 14020607
End Property

Public Property Get PeachPuff() As Long
    PeachPuff = 12180223
End Property

Public Property Get Peru() As Long
    Peru = 4163021
End Property

Public Property Get Pink() As Long
    Pink = 13353215
End Property

Public Property Get Plum() As Long
    Plum = 14524637
End Property

Public Property Get PowderBlue() As Long
    PowderBlue = 15130800
End Property

Public Property Get Purple() As Long
    Purple = 8388736
End Property

Public Property Get Red() As Long
    Red = 255
End Property

Public Property Get RosyBrown() As Long
    RosyBrown = 9408444
End Property

Public Property Get RoyalBlue() As Long
    RoyalBlue = 14772545
End Property

Public Property Get SaddleBrown() As Long
    SaddleBrown = 1262987
End Property

Public Property Get Salmon() As Long
    Salmon = 7504122
End Property

Public Property Get SandyBrown() As Long
    SandyBrown = 6333684
End Property

Public Property Get SeaGreen() As Long
    SeaGreen = 5737262
End Property

Public Property Get SeaShell() As Long
    SeaShell = 15660543
End Property

Public Property Get Sienna() As Long
    Sienna = 2970272
End Property

Public Property Get Silver() As Long
    Silver = 12632256
End Property

Public Property Get SkyBlue() As Long
    SkyBlue = 15453831
End Property
Public Property Get SlateBlue() As Long
    SlateBlue = 13458026
End Property

Public Property Get SlateGray() As Long
    SlateGray = 9470064
End Property

Public Property Get Snow() As Long
    Snow = 16448255
End Property

Public Property Get SpringGreen() As Long
    SpringGreen = 8388352
End Property

Public Property Get SteelBlue() As Long
    SteelBlue = 11829830
End Property

Public Property Get Tan() As Long
    Tan = 9221330
End Property

Public Property Get Teal() As Long
    Teal = 8421376
End Property

Public Property Get Thistle() As Long
    Thistle = 14204888
End Property

Public Property Get Tomato() As Long
    Tomato = 4678655
End Property

Public Property Get Turquoise() As Long
    Turquoise = 13688896
End Property

Public Property Get Violet() As Long
    Violet = 15631086
End Property

Public Property Get Wheat() As Long
    Wheat = 11788021
End Property

Public Property Get White() As Long
    White = 16777215
End Property

Public Property Get WhiteSmoke() As Long
    WhiteSmoke = 16119285
End Property

Public Property Get Yellow() As Long
    Yellow = 65535
End Property

Public Property Get YellowGreen() As Long
    YellowGreen = 3329434
End Property


' Configurable colors

'----------------------------
'    SECTION: ChannelList
'----------------------------

Public Property Get ChannelLabelBack() As Long
    ChannelLabelBack = m_ChannelLabelBack
End Property

Public Property Let ChannelLabelBack(ByVal lValue As Long)
    m_ChannelLabelBack = lValue
End Property

Public Property Get ChannelLabelText() As Long
    ChannelLabelText = m_ChannelLabelText
End Property

Public Property Let ChannelLabelText(ByVal lValue As Long)
    m_ChannelLabelText = lValue
End Property


Public Property Get ChannelListBack() As Long
    ChannelListBack = m_ChannelListBack
End Property

Public Property Let ChannelListBack(ByVal lValue As Long)
    m_ChannelListBack = lValue
End Property

Public Property Get ChannelListText() As Long
    ChannelListText = m_ChannelListText
End Property

Public Property Let ChannelListText(ByVal lValue As Long)
    m_ChannelListText = lValue
End Property

Public Property Get ChannelListSelf() As Long
    ChannelListSelf = m_ChannelListSelf
End Property

Public Property Let ChannelListSelf(ByVal lValue As Long)
    m_ChannelListSelf = lValue
End Property

Public Property Get ChannelListIdle() As Long
    ChannelListIdle = m_ChannelListIdle
End Property

Public Property Let ChannelListIdle(ByVal lValue As Long)
    m_ChannelListIdle = lValue
End Property

Public Property Get ChannelListSquelched() As Long
    ChannelListSquelched = m_ChannelListSquelched
End Property

Public Property Let ChannelListSquelched(ByVal lValue As Long)
    m_ChannelListSquelched = lValue
End Property

Public Property Get ChannelListOps() As Long
    ChannelListOps = m_ChannelListOps
End Property

Public Property Let ChannelListOps(ByVal lValue As Long)
    m_ChannelListOps = lValue
End Property

Public Property Get ChannelListAdmin() As Long
    ChannelListAdmin = m_ChannelListAdmin
End Property

Public Property Let ChannelListAdmin(ByVal lValue As Long)
    m_ChannelListAdmin = lValue
End Property



'----------------------------
'    SECTION: ChatUI
'----------------------------

Public Property Get SendBoxesBack() As Long
    SendBoxesBack = m_SendBoxesBack
End Property

Public Property Let SendBoxesBack(ByVal lValue As Long)
    m_SendBoxesBack = lValue
End Property

Public Property Get SendBoxesText() As Long
    SendBoxesText = m_SendBoxesText
End Property

Public Property Let SendBoxesText(ByVal lValue As Long)
    m_SendBoxesText = lValue
End Property


Public Property Get RTBBack() As Long
    RTBBack = m_RTBBack
End Property

Public Property Let RTBBack(ByVal lValue As Long)
    m_RTBBack = lValue
End Property


Public Property Get TalkBotUsername() As Long
    TalkBotUsername = m_TalkBotUsername
End Property

Public Property Let TalkBotUsername(ByVal lValue As Long)
    m_TalkBotUsername = lValue
End Property

Public Property Get TalkUsernameNormal() As Long
    TalkUsernameNormal = m_TalkUsernameNormal
End Property

Public Property Let TalkUsernameNormal(ByVal lValue As Long)
    m_TalkUsernameNormal = lValue
End Property

Public Property Get TalkUsernameOp() As Long
    TalkUsernameOp = m_TalkUsernameOp
End Property

Public Property Let TalkUsernameOp(ByVal lValue As Long)
    m_TalkUsernameOp = lValue
End Property

Public Property Get TalkUsernameAdmin() As Long
    TalkUsernameAdmin = m_TalkUsernameAdmin
End Property

Public Property Let TalkUsernameAdmin(ByVal lValue As Long)
    m_TalkUsernameAdmin = lValue
End Property

Public Property Get TalkNormalText() As Long
    TalkNormalText = m_TalkNormalText
End Property

Public Property Let TalkNormalText(ByVal lValue As Long)
    m_TalkNormalText = lValue
End Property


Public Property Get Carats() As Long
    Carats = m_Carats
End Property

Public Property Let Carats(ByVal lValue As Long)
    m_Carats = lValue
End Property

Public Property Get InformationText() As Long
    InformationText = m_InformationText
End Property

Public Property Let InformationText(ByVal lValue As Long)
    m_InformationText = lValue
End Property

Public Property Get SuccessText() As Long
    SuccessText = m_SuccessText
End Property

Public Property Let SuccessText(ByVal lValue As Long)
    m_SuccessText = lValue
End Property

Public Property Get ErrorMessageText() As Long
    ErrorMessageText = m_ErrorMessageText
End Property

Public Property Let ErrorMessageText(ByVal lValue As Long)
    m_ErrorMessageText = lValue
End Property

Public Property Get TimeStamps() As Long
    TimeStamps = m_TimeStamps
End Property

Public Property Let TimeStamps(ByVal lValue As Long)
    m_TimeStamps = lValue
End Property

Public Property Get ServerInfoText() As Long
    ServerInfoText = m_ServerInfoText
End Property

Public Property Let ServerInfoText(ByVal lValue As Long)
    m_ServerInfoText = lValue
End Property

Public Property Get ConsoleText() As Long
    ConsoleText = m_ConsoleText
End Property

Public Property Let ConsoleText(ByVal lValue As Long)
    m_ConsoleText = lValue
End Property


Public Property Get WhisperCarats() As Long
    WhisperCarats = m_WhisperCarats
End Property

Public Property Let WhisperCarats(ByVal lValue As Long)
    m_WhisperCarats = lValue
End Property

Public Property Get WhisperUsernames() As Long
    WhisperUsernames = m_WhisperUsernames
End Property

Public Property Let WhisperUsernames(ByVal lValue As Long)
    m_WhisperUsernames = lValue
End Property

Public Property Get WhisperText() As Long
    WhisperText = m_WhisperText
End Property

Public Property Let WhisperText(ByVal lValue As Long)
    m_WhisperText = lValue
End Property


Public Property Get EmoteUsernames() As Long
    EmoteUsernames = m_EmoteUsernames
End Property

Public Property Let EmoteUsernames(ByVal lValue As Long)
    m_EmoteUsernames = lValue
End Property

Public Property Get EmoteText() As Long
    EmoteText = m_EmoteText
End Property

Public Property Let EmoteText(ByVal lValue As Long)
    m_EmoteText = lValue
End Property


Public Property Get JoinUsername() As Long
    JoinUsername = m_JoinUsername
End Property

Public Property Let JoinUsername(ByVal lValue As Long)
    m_JoinUsername = lValue
End Property

Public Property Get JoinText() As Long
    JoinText = m_JoinText
End Property

Public Property Let JoinText(ByVal lValue As Long)
    m_JoinText = lValue
End Property

Public Property Get JoinedChannelName() As Long
    JoinedChannelName = m_JoinedChannelName
End Property

Public Property Let JoinedChannelName(ByVal lValue As Long)
    m_JoinedChannelName = lValue
End Property

Public Property Get JoinedChannelText() As Long
    JoinedChannelText = m_JoinedChannelText
End Property

Public Property Let JoinedChannelText(ByVal lValue As Long)
    m_JoinedChannelText = lValue
End Property


'// shows a list of all the colors in the chat window
Public Sub List()
    frmChat.AddChat ChannelLabelBack, StringFormat("ChannelLabelBack{0}{0}({1})", vbTab, ToHex(ChannelLabelBack))
    frmChat.AddChat ChannelLabelText, StringFormat("ChannelLabelText{0}{0}({1})", vbTab, ToHex(ChannelLabelText))

    frmChat.AddChat ChannelListBack, StringFormat("ChannelListBack{0}{0}({1})", vbTab, ToHex(ChannelListBack))
    frmChat.AddChat ChannelListText, StringFormat("ChannelListText{0}{0}({1})", vbTab, ToHex(ChannelListText))
    frmChat.AddChat ChannelListSelf, StringFormat("ChannelListSelf{0}{0}({1})", vbTab, ToHex(ChannelListSelf))
    frmChat.AddChat ChannelListIdle, StringFormat("ChannelListIdle{0}{0}({1})", vbTab, ToHex(ChannelListIdle))
    frmChat.AddChat ChannelListSquelched, StringFormat("ChannelListSquelched{0}{0}({1})", vbTab, ToHex(ChannelListSquelched))
    frmChat.AddChat ChannelListOps, StringFormat("ChannelListOps{0}{0}({1})", vbTab, ToHex(ChannelListOps))
    frmChat.AddChat ChannelListAdmin, StringFormat("ChannelListAdmin{0}{0}({1})", vbTab, ToHex(ChannelListAdmin))

    frmChat.AddChat SendBoxesBack, StringFormat("SendBoxesBack{0}{0}({1})", vbTab, ToHex(SendBoxesBack))
    frmChat.AddChat SendBoxesText, StringFormat("SendBoxesText{0}{0}({1})", vbTab, ToHex(SendBoxesText))

    frmChat.AddChat RTBBack, StringFormat("RTBBack{0}{0}({1})", vbTab, ToHex(RTBBack))

    frmChat.AddChat TalkBotUsername, StringFormat("TalkBotUsername{0}{0}({1})", vbTab, ToHex(TalkBotUsername))
    frmChat.AddChat TalkUsernameNormal, StringFormat("TalkUsernameNormal{0}{0}({1})", vbTab, ToHex(TalkUsernameNormal))
    frmChat.AddChat TalkUsernameOp, StringFormat("TalkUsernameOp{0}{0}({1})", vbTab, ToHex(TalkUsernameOp))
    frmChat.AddChat TalkUsernameAdmin, StringFormat("TalkUsernameAdmin{0}{0}({1})", vbTab, ToHex(TalkUsernameAdmin))
    frmChat.AddChat TalkNormalText, StringFormat("TalkNormalText{0}{0}({1})", vbTab, ToHex(TalkNormalText))

    frmChat.AddChat Carats, StringFormat("Carats{0}{0}({1})", vbTab, ToHex(Carats))
    frmChat.AddChat InformationText, StringFormat("InformationText{0}{0}({1})", vbTab, ToHex(InformationText))
    frmChat.AddChat SuccessText, StringFormat("SuccessText{0}{0}({1})", vbTab, ToHex(SuccessText))
    frmChat.AddChat ErrorMessageText, StringFormat("ErrorMessageText{0}{0}({1})", vbTab, ToHex(ErrorMessageText))
    frmChat.AddChat TimeStamps, StringFormat("TimeStamps{0}{0}({1})", vbTab, ToHex(TimeStamps))
    frmChat.AddChat ServerInfoText, StringFormat("ServerInfoText{0}{0}({1})", vbTab, ToHex(ServerInfoText))
    frmChat.AddChat ConsoleText, StringFormat("ConsoleText{0}{0}({1})", vbTab, ToHex(ConsoleText))

    frmChat.AddChat WhisperCarats, StringFormat("WhisperCarats{0}{0}({1})", vbTab, ToHex(WhisperCarats))
    frmChat.AddChat WhisperUsernames, StringFormat("WhisperUsernames{0}{0}({1})", vbTab, ToHex(WhisperUsernames))
    frmChat.AddChat WhisperText, StringFormat("WhisperText{0}{0}({1})", vbTab, ToHex(WhisperText))

    frmChat.AddChat EmoteUsernames, StringFormat("EmoteUsernames{0}{0}({1})", vbTab, ToHex(EmoteUsernames))
    frmChat.AddChat EmoteText, StringFormat("EmoteText{0}{0}({1})", vbTab, ToHex(EmoteText))

    frmChat.AddChat JoinUsername, StringFormat("JoinUsername{0}{0}({1})", vbTab, ToHex(JoinUsername))
    frmChat.AddChat JoinText, StringFormat("JoinText{0}{0}({1})", vbTab, ToHex(JoinText))
    frmChat.AddChat JoinedChannelName, StringFormat("JoinedChannelName{0}{0}({1})", vbTab, ToHex(JoinedChannelName))
    frmChat.AddChat JoinedChannelText, StringFormat("JoinedChannelText{0}{0}({1})", vbTab, ToHex(JoinedChannelText))

End Sub


Public Property Get FilePath() As String
    FilePath = m_ColorFilePath
End Property

Public Function Load(ByVal sFilePath As String)
    m_ColorFilePath = sFilePath
    
    ' If the color file doesn't exist, load defaults.
    If Len(Dir(m_ColorFilePath)) = 0 Then Exit Function
    
    Load = False
    If LCase(Right(sFilePath, 5)) = ".sclf" Then
        Load = LoadSCLF
    ElseIf LCase(Right(sFilePath, 4)) = ".ini" Then
        Load = LoadINI
    End If
End Function

Public Sub Save(Optional ByVal sFilePath As String = vbNullString)
    If Len(sFilePath) > 0 Then
        m_ColorFilePath = sFilePath
    End If
    
    WriteEntryRaw "META", "Version", COLOR_VERSION
    
    WriteEntry SECTION_CHANNELLIST, "LabelBackground", m_ChannelLabelBack
    WriteEntry SECTION_CHANNELLIST, "LabelText", m_ChannelLabelText

    WriteEntry SECTION_CHANNELLIST, "Background", m_ChannelListBack
    WriteEntry SECTION_CHANNELLIST, "Usernames", m_ChannelListText
    WriteEntry SECTION_CHANNELLIST, "Self", m_ChannelListSelf
    WriteEntry SECTION_CHANNELLIST, "Idle", m_ChannelListIdle
    WriteEntry SECTION_CHANNELLIST, "Squelched", m_ChannelListSquelched
    WriteEntry SECTION_CHANNELLIST, "Operator", m_ChannelListOps
    WriteEntry SECTION_CHANNELLIST, "Admin", m_ChannelListAdmin

    WriteEntry SECTION_CHATUI, "SendBoxBackground", m_SendBoxesBack
    WriteEntry SECTION_CHATUI, "SendBoxText", m_SendBoxesText

    WriteEntry SECTION_CHATUI, "ChatBackground", m_RTBBack

    WriteEntry SECTION_CHATUI, "BotUsername", m_TalkBotUsername
    WriteEntry SECTION_CHATUI, "NormalUsername", m_TalkUsernameNormal
    WriteEntry SECTION_CHATUI, "OperatorUsername", m_TalkUsernameOp
    WriteEntry SECTION_CHATUI, "AdminUsername", m_TalkUsernameAdmin
    WriteEntry SECTION_CHATUI, "NormalText", m_TalkNormalText

    WriteEntry SECTION_CHATUI, "Carats", m_Carats
    WriteEntry SECTION_CHATUI, "InformationText", m_InformationText
    WriteEntry SECTION_CHATUI, "SuccessText", m_SuccessText
    WriteEntry SECTION_CHATUI, "ErrorMessageText", m_ErrorMessageText
    WriteEntry SECTION_CHATUI, "TimeStamps", m_TimeStamps
    WriteEntry SECTION_CHATUI, "ServerInfoText", m_ServerInfoText
    WriteEntry SECTION_CHATUI, "ConsoleText", m_ConsoleText

    WriteEntry SECTION_CHATUI, "WhisperCarats", m_WhisperCarats
    WriteEntry SECTION_CHATUI, "WhisperUsernames", m_WhisperUsernames
    WriteEntry SECTION_CHATUI, "WhisperText", m_WhisperText

    WriteEntry SECTION_CHATUI, "EmoteUsernames", m_EmoteUsernames
    WriteEntry SECTION_CHATUI, "EmoteText", m_EmoteText

    WriteEntry SECTION_CHATUI, "JoinUsername", m_JoinUsername
    WriteEntry SECTION_CHATUI, "JoinText", m_JoinText
    WriteEntry SECTION_CHATUI, "JoinedChannelName", m_JoinedChannelName
    WriteEntry SECTION_CHATUI, "JoinedChannelText", m_JoinedChannelText
End Sub

Private Sub LoadDefaults()
    m_ChannelLabelBack = &HCC3300
    m_ChannelLabelText = vbWhite
    
    m_ChannelListBack = vbBlack
    m_ChannelListText = &H99CC00
    m_ChannelListSelf = vbWhite
    m_ChannelListIdle = &HBBBBBB
    m_ChannelListSquelched = &H99
    m_ChannelListOps = &HDDDDDD
    m_ChannelListAdmin = &HCC9900
    
    m_SendBoxesBack = vbBlack
    m_SendBoxesText = vbWhite
    
    m_RTBBack = vbBlack
    
    m_TalkBotUsername = vbCyan
    m_TalkUsernameNormal = vbYellow
    m_TalkUsernameOp = vbWhite
    m_TalkUsernameAdmin = &HFF6961
    m_TalkNormalText = vbWhite
    
    m_Carats = &HCC9900
    m_InformationText = vbYellow
    m_SuccessText = vbGreen
    m_ErrorMessageText = vbRed
    m_TimeStamps = vbWhite
    m_ServerInfoText = &HCC9900
    m_ConsoleText = &H99CC00
    
    m_WhisperCarats = &H80FF&
    m_WhisperUsernames = vbYellow
    m_WhisperText = &H999999
    
    m_EmoteUsernames = vbWhite
    m_EmoteText = vbYellow
    
    m_JoinUsername = vbYellow
    m_JoinText = vbGreen
    m_JoinedChannelName = vbYellow
    m_JoinedChannelText = &H99CC00
End Sub

Private Function LoadSCLF()
    Dim iVersion As Integer
    Dim f, i As Integer
    
    iVersion = 0
    i = 1
    
    f = FreeFile
    Open m_ColorFilePath For Random As #f Len = 4
    
    Select Case LOF(f)
        Case 112
            iVersion = 1
        Case 128
            iVersion = 2
        Case Else
            LoadSCLF = False
            Exit Function
    End Select

    m_ChannelLabelBack = GetRecord(f, i)
    m_ChannelLabelText = GetRecord(f, i)
    m_ChannelListBack = GetRecord(f, i)
    m_ChannelListText = GetRecord(f, i)
    
    If iVersion = 2 Then
        m_ChannelListSelf = GetRecord(f, i)
        m_ChannelListIdle = GetRecord(f, i)
        m_ChannelListSquelched = GetRecord(f, i)
        m_ChannelListOps = GetRecord(f, i)
    End If
    
    m_RTBBack = GetRecord(f, i)
    m_SendBoxesBack = GetRecord(f, i)
    m_SendBoxesText = GetRecord(f, i)
    
    m_TalkBotUsername = GetRecord(f, i)
    m_TalkUsernameNormal = GetRecord(f, i)
    m_TalkUsernameOp = GetRecord(f, i)
    m_TalkNormalText = GetRecord(f, i)
    m_Carats = GetRecord(f, i)
    m_EmoteText = GetRecord(f, i)
    m_EmoteUsernames = GetRecord(f, i)
    m_InformationText = GetRecord(f, i)
    m_SuccessText = GetRecord(f, i)
    m_ErrorMessageText = GetRecord(f, i)
    m_TimeStamps = GetRecord(f, i)
    m_ServerInfoText = GetRecord(f, i)
    m_ConsoleText = GetRecord(f, i)
    m_JoinText = GetRecord(f, i)
    m_JoinUsername = GetRecord(f, i)
    m_JoinedChannelName = GetRecord(f, i)
    m_JoinedChannelText = GetRecord(f, i)
    m_WhisperCarats = GetRecord(f, i)
    m_WhisperText = GetRecord(f, i)
    m_WhisperUsernames = GetRecord(f, i)
    
    Close #f
    
    LoadSCLF = True
End Function

Private Function LoadINI()
    Dim lVersion As Long
    lVersion = ReadEntry("META", "Version")
    
    LoadINI = True
    Select Case lVersion
        Case 3: Call LoadINIVersion3
        Case Else
            LoadINI = False
    End Select
End Function

Private Sub LoadINIVersion3()
    m_ChannelLabelBack = ReadEntry(SECTION_CHANNELLIST, "LabelBackground", m_ChannelLabelBack)
    m_ChannelLabelText = ReadEntry(SECTION_CHANNELLIST, "LabelText", m_ChannelLabelText)

    m_ChannelListBack = ReadEntry(SECTION_CHANNELLIST, "Background", m_ChannelListBack)
    m_ChannelListText = ReadEntry(SECTION_CHANNELLIST, "Text", m_ChannelListText)
    m_ChannelListSelf = ReadEntry(SECTION_CHANNELLIST, "Self", m_ChannelListSelf)
    m_ChannelListIdle = ReadEntry(SECTION_CHANNELLIST, "Idle", m_ChannelListIdle)
    m_ChannelListSquelched = ReadEntry(SECTION_CHANNELLIST, "Squelched", m_ChannelListSquelched)
    m_ChannelListOps = ReadEntry(SECTION_CHANNELLIST, "Operator", m_ChannelListOps)
    m_ChannelListAdmin = ReadEntry(SECTION_CHANNELLIST, "Admin", m_ChannelListAdmin)

    m_SendBoxesBack = ReadEntry(SECTION_CHATUI, "SendBoxBackground", m_SendBoxesBack)
    m_SendBoxesText = ReadEntry(SECTION_CHATUI, "SendBoxText", m_SendBoxesText)

    m_RTBBack = ReadEntry(SECTION_CHATUI, "ChatBackground", m_RTBBack)

    m_TalkBotUsername = ReadEntry(SECTION_CHATUI, "BotUsername", m_TalkBotUsername)
    m_TalkUsernameNormal = ReadEntry(SECTION_CHATUI, "NormalUsername", m_TalkUsernameNormal)
    m_TalkUsernameOp = ReadEntry(SECTION_CHATUI, "OperatorUsername", m_TalkUsernameOp)
    m_TalkUsernameAdmin = ReadEntry(SECTION_CHATUI, "AdminUsername", m_TalkUsernameAdmin)
    m_TalkNormalText = ReadEntry(SECTION_CHATUI, "NormalText", m_TalkNormalText)

    m_Carats = ReadEntry(SECTION_CHATUI, "Carats", m_Carats)
    m_InformationText = ReadEntry(SECTION_CHATUI, "InformationText", m_InformationText)
    m_SuccessText = ReadEntry(SECTION_CHATUI, "SuccessText", m_SuccessText)
    m_ErrorMessageText = ReadEntry(SECTION_CHATUI, "ErrorMessageText", m_ErrorMessageText)
    m_TimeStamps = ReadEntry(SECTION_CHATUI, "TimeStamps", m_TimeStamps)
    m_ServerInfoText = ReadEntry(SECTION_CHATUI, "ServerInfoText", m_ServerInfoText)
    m_ConsoleText = ReadEntry(SECTION_CHATUI, "ConsoleText", m_ConsoleText)

    m_WhisperCarats = ReadEntry(SECTION_CHATUI, "WhisperCarats", m_WhisperCarats)
    m_WhisperUsernames = ReadEntry(SECTION_CHATUI, "WhisperUsernames", m_WhisperUsernames)
    m_WhisperText = ReadEntry(SECTION_CHATUI, "WhisperText", m_WhisperText)

    m_EmoteUsernames = ReadEntry(SECTION_CHATUI, "EmoteUsernames", m_EmoteUsernames)
    m_EmoteText = ReadEntry(SECTION_CHATUI, "EmoteText", m_EmoteText)

    m_JoinUsername = ReadEntry(SECTION_CHATUI, "JoinUsername", m_JoinUsername)
    m_JoinText = ReadEntry(SECTION_CHATUI, "JoinText", m_JoinText)
    m_JoinedChannelName = ReadEntry(SECTION_CHATUI, "JoinedChannelName", m_JoinedChannelName)
    m_JoinedChannelText = ReadEntry(SECTION_CHATUI, "JoinedChannelText", m_JoinedChannelText)
End Sub

' Converts long value to CSS hex string.
'  Can pass in color properties of this class.
'  Len(ToHex(n)) will always be 6.
Public Function ToHex(ByVal lValue As Long) As String
    Dim bR, bG, bB As Byte  ' MUST BE BYTE
    
    bR = lValue And &HFF
    bG = (lValue And &HFF00&) \ 256
    bB = lValue \ 65536
    
    ToHex = "#" & PadHex(bR) & PadHex(bG) & PadHex(bB)
End Function

' Converts a CSS hex string to a long value.
Public Function FromHex(ByVal sValue As String) As Long
    Dim sR As String * 4
    Dim sG As String * 4
    Dim sB As String * 4
    
    If Left$(sValue, 1) = "#" Then sValue = Mid(sValue, 2)
    
    sR = "&H" & Mid$(sValue, 1, 2)
    sG = "&H" & Mid$(sValue, 3, 2)
    sB = "&H" & Mid$(sValue, 5, 2)
    
    FromHex = RGB(Val(sR), Val(sG), Val(sB))
End Function

Private Function ReadEntry(ByVal sSection As String, ByVal sKey As String, Optional ByVal lDefaultValue As Long = -1) As Long
    Dim sBuffer As String
    Dim lLength As Long
    
    If Len(Dir(m_ColorFilePath)) = 0 Then
        ReadEntry = lDefaultValue
        Exit Function
    End If
    
    sBuffer = String(255, vbNull)
    
    lLength = GetPrivateProfileString(sSection, sKey, ToHex(lDefaultValue), sBuffer, 255, m_ColorFilePath)
    ReadEntry = FromHex(Left$(sBuffer, lLength))
End Function

Private Sub WriteEntry(ByVal sSection As String, ByVal sKey As String, ByVal lValue As Long)
    WriteEntryRaw sSection, sKey, ToHex(lValue)
End Sub

Private Sub WriteEntryRaw(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String)
    WritePrivateProfileString sSection, sKey, sValue, m_ColorFilePath
End Sub


' Used for getting records from an SCLF file.
Private Function GetRecord(ByVal f As Integer, ByRef i As Integer) As Long
    Dim lValue As Long
    Get #f, i, lValue
    GetRecord = lValue
    i = i + 1
End Function

Private Function PadHex(ByVal bByte As Byte) As String
    PadHex = Right("00" & Hex(bByte), 2)
End Function

