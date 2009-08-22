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


Dim namedColor As Scripting.Dictionary


Public Sub Class_Terminate()
    Set namedColor = Nothing
End Sub

Public Function Hex(hexcolor) As Long
    Dim r, g, b, os
    If Left(hexcolor, 1) = "#" Then os = 1
    r = "&H" & Mid(hexcolor, 1 + os, 2)
    g = "&H" & Mid(hexcolor, 3 + os, 2)
    b = "&H" & Mid(hexcolor, 5 + os, 2)
    Hex = RGB(CInt(r), CInt(g), CInt(b))
End Function

'// shows a list of all the colors in the chat window
Public Sub List()

    Dim prm(0 To 1) As Variant
    Dim sColor, a, i

    If namedColor Is Nothing Then
        Call CreateDictionary
    End If

    
    a = namedColor.Keys
    For i = 0 To namedColor.Count - 1
        frmChat.AddChat CSS(a(i)), StringFormat("{0}{2}{2}({1})", a(i), CSS(a(i)), vbTab)
    Next i

End Sub

'// returns the value of a color from the dictionary
Private Function CSS(colorName) As Long
    
    Call CreateDictionary
        
    If namedColor.Exists(colorName) Then
        CSS = Hex(namedColor.Item(colorName))
    Else
        CSS = vbWhite
    End If
    
End Function

Private Sub CreateDictionary()

    If Not (namedColor Is Nothing) Then
        '// already created
        Exit Sub
    End If

    '// create the dictionary
    Set namedColor = CreateObject("Scripting.Dictionary")
    namedColor.CompareMode = 1

    namedColor.Add "AliceBlue", "#F0F8FF"
    namedColor.Add "AntiqueWhite", "#FAEBD7"
    namedColor.Add "Aqua", "#00FFFF"
    namedColor.Add "Aquamarine", "#7FFFD4"
    namedColor.Add "Azure", "#F0FFFF"
    namedColor.Add "Beige", "#F5F5DC"
    namedColor.Add "Bisque", "#FFE4C4"
    namedColor.Add "Black", "#000000"
    namedColor.Add "BlanchedAlmond", "#FFEBCD"
    namedColor.Add "Blue", "#0000FF"
    namedColor.Add "BlueViolet", "#8A2BE2"
    namedColor.Add "Brown", "#A52A2A"
    namedColor.Add "BurlyWood", "#DEB887"
    namedColor.Add "CadetBlue", "#5F9EA0"
    namedColor.Add "Chartreuse", "#7FFF00"
    namedColor.Add "Chocolate", "#D2691E"
    namedColor.Add "Coral", "#FF7F50"
    namedColor.Add "CornflowerBlue", "#6495ED"
    namedColor.Add "Cornsilk", "#FFF8DC"
    namedColor.Add "Crimson", "#DC143C"
    namedColor.Add "Cyan", "#00FFFF"
    namedColor.Add "DarkBlue", "#00008B"
    namedColor.Add "DarkCyan", "#008B8B"
    namedColor.Add "DarkGoldenRod", "#B8860B"
    namedColor.Add "DarkGray", "#A9A9A9"
    namedColor.Add "DarkGreen", "#006400"
    namedColor.Add "DarkKhaki", "#BDB76B"
    namedColor.Add "DarkMagenta", "#8B008B"
    namedColor.Add "DarkOliveGreen", "#556B2F"
    namedColor.Add "Darkorange", "#FF8C00"
    namedColor.Add "DarkOrchid", "#9932CC"
    namedColor.Add "DarkRed", "#8B0000"
    namedColor.Add "DarkSalmon", "#E9967A"
    namedColor.Add "DarkSeaGreen", "#8FBC8F"
    namedColor.Add "DarkSlateBlue", "#483D8B"
    namedColor.Add "DarkSlateGray", "#2F4F4F"
    namedColor.Add "DarkTurquoise", "#00CED1"
    namedColor.Add "DarkViolet", "#9400D3"
    namedColor.Add "DeepPink", "#FF1493"
    namedColor.Add "DeepSkyBlue", "#00BFFF"
    namedColor.Add "DimGray", "#696969"
    namedColor.Add "DodgerBlue", "#1E90FF"
    namedColor.Add "FireBrick", "#B22222"
    namedColor.Add "FloralWhite", "#FFFAF0"
    namedColor.Add "ForestGreen", "#228B22"
    namedColor.Add "Fuchsia", "#FF00FF"
    namedColor.Add "Gainsboro", "#DCDCDC"
    namedColor.Add "GhostWhite", "#F8F8FF"
    namedColor.Add "Gold", "#FFD700"
    namedColor.Add "GoldenRod", "#DAA520"
    namedColor.Add "Gray", "#808080"
    namedColor.Add "Green", "#008000"
    namedColor.Add "GreenYellow", "#ADFF2F"
    namedColor.Add "HoneyDew", "#F0FFF0"
    namedColor.Add "HotPink", "#FF69B4"
    namedColor.Add "IndianRed ", "#CD5C5C"
    namedColor.Add "Indigo ", "#4B0082"
    namedColor.Add "Ivory", "#FFFFF0"
    namedColor.Add "Khaki", "#F0E68C"
    namedColor.Add "Lavender", "#E6E6FA"
    namedColor.Add "LavenderBlush", "#FFF0F5"
    namedColor.Add "LawnGreen", "#7CFC00"
    namedColor.Add "LemonChiffon", "#FFFACD"
    namedColor.Add "LightBlue", "#ADD8E6"
    namedColor.Add "LightCoral", "#F08080"
    namedColor.Add "LightCyan", "#E0FFFF"
    namedColor.Add "LightGoldenRodYellow", "#FAFAD2"
    namedColor.Add "LightGray", "#D3D3D3"
    namedColor.Add "LightGreen", "#90EE90"
    namedColor.Add "LightPink", "#FFB6C1"
    namedColor.Add "LightSalmon", "#FFA07A"
    namedColor.Add "LightSeaGreen", "#20B2AA"
    namedColor.Add "LightSkyBlue", "#87CEFA"
    namedColor.Add "LightSlateGray", "#778899"
    namedColor.Add "LightSteelBlue", "#B0C4DE"
    namedColor.Add "LightYellow", "#FFFFE0"
    namedColor.Add "Lime", "#00FF00"
    namedColor.Add "LimeGreen", "#32CD32"
    namedColor.Add "Linen", "#FAF0E6"
    namedColor.Add "Magenta", "#FF00FF"
    namedColor.Add "Maroon", "#800000"
    namedColor.Add "MediumAquaMarine", "#66CDAA"
    namedColor.Add "MediumBlue", "#0000CD"
    namedColor.Add "MediumOrchid", "#BA55D3"
    namedColor.Add "MediumPurple", "#9370D8"
    namedColor.Add "MediumSeaGreen", "#3CB371"
    namedColor.Add "MediumSlateBlue", "#7B68EE"
    namedColor.Add "MediumSpringGreen", "#00FA9A"
    namedColor.Add "MediumTurquoise", "#48D1CC"
    namedColor.Add "MediumVioletRed", "#C71585"
    namedColor.Add "MidnightBlue", "#191970"
    namedColor.Add "MintCream", "#F5FFFA"
    namedColor.Add "MistyRose", "#FFE4E1"
    namedColor.Add "Moccasin", "#FFE4B5"
    namedColor.Add "NavajoWhite", "#FFDEAD"
    namedColor.Add "Navy", "#000080"
    namedColor.Add "OldLace", "#FDF5E6"
    namedColor.Add "Olive", "#808000"
    namedColor.Add "OliveDrab", "#6B8E23"
    namedColor.Add "Orange", "#FFA500"
    namedColor.Add "OrangeRed", "#FF4500"
    namedColor.Add "Orchid", "#DA70D6"
    namedColor.Add "PaleGoldenRod", "#EEE8AA"
    namedColor.Add "PaleGreen", "#98FB98"
    namedColor.Add "PaleTurquoise", "#AFEEEE"
    namedColor.Add "PaleVioletRed", "#D87093"
    namedColor.Add "PapayaWhip", "#FFEFD5"
    namedColor.Add "PeachPuff", "#FFDAB9"
    namedColor.Add "Peru", "#CD853F"
    namedColor.Add "Pink", "#FFC0CB"
    namedColor.Add "Plum", "#DDA0DD"
    namedColor.Add "PowderBlue", "#B0E0E6"
    namedColor.Add "Purple", "#800080"
    namedColor.Add "Red", "#FF0000"
    namedColor.Add "RosyBrown", "#BC8F8F"
    namedColor.Add "RoyalBlue", "#4169E1"
    namedColor.Add "SaddleBrown", "#8B4513"
    namedColor.Add "Salmon", "#FA8072"
    namedColor.Add "SandyBrown", "#F4A460"
    namedColor.Add "SeaGreen", "#2E8B57"
    namedColor.Add "SeaShell", "#FFF5EE"
    namedColor.Add "Sienna", "#A0522D"
    namedColor.Add "Silver", "#C0C0C0"
    namedColor.Add "SkyBlue", "#87CEEB"
    namedColor.Add "SlateBlue", "#6A5ACD"
    namedColor.Add "SlateGray", "#708090"
    namedColor.Add "Snow", "#FFFAFA"
    namedColor.Add "SpringGreen", "#00FF7F"
    namedColor.Add "SteelBlue", "#4682B4"
    namedColor.Add "Tan", "#D2B48C"
    namedColor.Add "Teal", "#008080"
    namedColor.Add "Thistle", "#D8BFD8"
    namedColor.Add "Tomato", "#FF6347"
    namedColor.Add "Turquoise", "#40E0D0"
    namedColor.Add "Violet", "#EE82EE"
    namedColor.Add "Wheat", "#F5DEB3"
    namedColor.Add "White", "#FFFFFF"
    namedColor.Add "WhiteSmoke", "#F5F5F5"
    namedColor.Add "Yellow", "#FFFF00"
    namedColor.Add "YellowGreen", "#9ACD32"
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
