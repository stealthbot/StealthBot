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

Public Function Hex(ByVal hexcolor As String) As Long
    Dim r As String
    Dim g As String
    Dim b As String
    Dim os As Integer
    If Left$(hexcolor, 1) = "#" Then os = 1
    r = "&H" & Mid$(hexcolor, 1 + os, 2)
    g = "&H" & Mid$(hexcolor, 3 + os, 2)
    b = "&H" & Mid$(hexcolor, 5 + os, 2)
    Hex = RGB(CInt(r), CInt(g), CInt(b))
End Function

'// shows a list of all the colors in the chat window
Public Sub List()
    
    Dim IInfo   As InterfaceInfo
    Dim MInfo   As MemberInfo
    Dim Color   As Long
    Dim Name    As String
    
    Set IInfo = TLI.InterfaceInfoFromObject(Me)
    
    For Each MInfo In IInfo.Members
        If MInfo.InvokeKind = INVOKE_PROPERTYGET Then
            Name = MInfo.Name
            Color = CallByName(Me, MInfo.Name, VbGet)
            frmChat.AddChat Color, StringFormat("{0}{2}{2}({1})", Name, Color, vbTab)
        End If
        DoEvents
    Next MInfo
    
    Set MInfo = Nothing
    Set IInfo = Nothing
    
End Sub

'// properties do not use the internal dictionary
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

Public Property Get ChannelLabelBack() As Long
    ChannelLabelBack = FormColors.ChannelLabelBack
End Property
Public Property Get ChannelLabelText() As Long
    ChannelLabelText = FormColors.ChannelLabelText
End Property
Public Property Get ChannelListBack() As Long
    ChannelListBack = FormColors.ChannelListBack
End Property
Public Property Get ChannelListText() As Long
    ChannelListText = FormColors.ChannelListText
End Property
Public Property Get RTBBack() As Long
    RTBBack = FormColors.RTBBack
End Property
Public Property Get SendBoxesBack() As Long
    SendBoxesBack = FormColors.SendBoxesBack
End Property
Public Property Get SendBoxesText() As Long
    SendBoxesText = FormColors.SendBoxesText
End Property

Public Property Get Carots() As Long
    Carots = RTBColors.Carats
End Property
Public Property Get ConsoleText() As Long
    ConsoleText = RTBColors.ConsoleText
End Property
Public Property Get EmoteText() As Long
    EmoteText = RTBColors.EmoteText
End Property
Public Property Get EmoteUsernames() As Long
    EmoteUsernames = RTBColors.EmoteUsernames
End Property
Public Property Get ErrorMessageText() As Long
    ErrorMessageText = RTBColors.ErrorMessageText
End Property
Public Property Get InformationText() As Long
    InformationText = RTBColors.InformationText
End Property
Public Property Get JoinedChannelName() As Long
    JoinedChannelName = RTBColors.JoinedChannelName
End Property
Public Property Get JoinedChannelText() As Long
    JoinedChannelText = RTBColors.JoinedChannelText
End Property
Public Property Get JoinText() As Long
    JoinText = RTBColors.JoinText
End Property
Public Property Get JoinUsername() As Long
    JoinUsername = RTBColors.JoinUsername
End Property
Public Property Get ServerInfoText() As Long
    ServerInfoText = RTBColors.ServerInfoText
End Property
Public Property Get SuccessText() As Long
    SuccessText = RTBColors.SuccessText
End Property
Public Property Get TalkBotUsername() As Long
    TalkBotUsername = RTBColors.TalkBotUsername
End Property
Public Property Get TalkNormalText() As Long
    TalkNormalText = RTBColors.TalkNormalText
End Property
Public Property Get TalkUsernameNormal() As Long
    TalkUsernameNormal = RTBColors.TalkUsernameNormal
End Property
Public Property Get TalkUsernameOp() As Long
    TalkUsernameOp = RTBColors.TalkUsernameOp
End Property
Public Property Get TimeStamps() As Long
    TimeStamps = RTBColors.TimeStamps
End Property
Public Property Get WhisperCarots() As Long
    WhisperCarots = RTBColors.WhisperCarats
End Property
Public Property Get WhisperText() As Long
    WhisperText = RTBColors.WhisperText
End Property
Public Property Get WhisperUsernames() As Long
    WhisperUsernames = RTBColors.WhisperUsernames
End Property
