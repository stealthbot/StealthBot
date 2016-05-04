Option Strict Off
Option Explicit On
Friend Class clsColor
	
	Private Named As Scripting.Dictionary
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object Named may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Named = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	'UPGRADE_NOTE: Hex was upgraded to Hex_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Hex_Renamed(ByVal hexcolor As String) As Integer
		Dim r As New VB6.FixedLengthString(4)
		Dim g As New VB6.FixedLengthString(4)
		Dim B As New VB6.FixedLengthString(4)
		Dim os As Short
		If Left(hexcolor, 1) = "#" Then os = 1
		r.Value = "&H" & Mid(hexcolor, 1 + os, 2)
		g.Value = "&H" & Mid(hexcolor, 3 + os, 2)
		B.Value = "&H" & Mid(hexcolor, 5 + os, 2)
		Hex_Renamed = RGB(CShort(r.Value), CShort(g.Value), CShort(B.Value))
	End Function
	
	'// shows a list of all the colors in the chat window
	Public Sub List()
		
		Dim keys() As Object
		Dim i As Short
		Dim Co As Integer
		
		If Named Is Nothing Then PopulateList()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Named.keys(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		keys = Named.keys()
		For i = 0 To Named.Count - 1
			Co = CSS(keys(i))
			frmChat.AddChat(Co, StringFormat("{0}{2}{2}({1})", keys(i), Co, vbTab))
			System.Windows.Forms.Application.DoEvents()
		Next i
		
	End Sub
	
	'// returns the value of a color from the dictionary
	Private Function CSS(ByRef colorName As Object) As Integer
		
		If Named.Exists(colorName) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Named.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CSS = Hex_Renamed(Named.Item(colorName))
		Else
			CSS = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
		End If
		
	End Function
	
	Private Sub PopulateList()
		
		If Not (Named Is Nothing) Then
			'// already created
			Exit Sub
		End If
		
		'// create the dictionary
		Named = New Scripting.Dictionary
		
		With Named
			.CompareMode = 1
			
			.Add("AliceBlue", "#F0F8FF")
			.Add("AntiqueWhite", "#FAEBD7")
			.Add("Aqua", "#00FFFF")
			.Add("Aquamarine", "#7FFFD4")
			.Add("Azure", "#F0FFFF")
			.Add("Beige", "#F5F5DC")
			.Add("Bisque", "#FFE4C4")
			.Add("Black", "#000000")
			.Add("BlanchedAlmond", "#FFEBCD")
			.Add("Blue", "#0000FF")
			.Add("BlueViolet", "#8A2BE2")
			.Add("Brown", "#A52A2A")
			.Add("BurlyWood", "#DEB887")
			.Add("CadetBlue", "#5F9EA0")
			.Add("Chartreuse", "#7FFF00")
			.Add("Chocolate", "#D2691E")
			.Add("Coral", "#FF7F50")
			.Add("CornflowerBlue", "#6495ED")
			.Add("Cornsilk", "#FFF8DC")
			.Add("Crimson", "#DC143C")
			.Add("Cyan", "#00FFFF")
			.Add("DarkBlue", "#00008B")
			.Add("DarkCyan", "#008B8B")
			.Add("DarkGoldenRod", "#B8860B")
			.Add("DarkGray", "#A9A9A9")
			.Add("DarkGreen", "#006400")
			.Add("DarkKhaki", "#BDB76B")
			.Add("DarkMagenta", "#8B008B")
			.Add("DarkOliveGreen", "#556B2F")
			.Add("Darkorange", "#FF8C00")
			.Add("DarkOrchid", "#9932CC")
			.Add("DarkRed", "#8B0000")
			.Add("DarkSalmon", "#E9967A")
			.Add("DarkSeaGreen", "#8FBC8F")
			.Add("DarkSlateBlue", "#483D8B")
			.Add("DarkSlateGray", "#2F4F4F")
			.Add("DarkTurquoise", "#00CED1")
			.Add("DarkViolet", "#9400D3")
			.Add("DeepPink", "#FF1493")
			.Add("DeepSkyBlue", "#00BFFF")
			.Add("DimGray", "#696969")
			.Add("DodgerBlue", "#1E90FF")
			.Add("FireBrick", "#B22222")
			.Add("FloralWhite", "#FFFAF0")
			.Add("ForestGreen", "#228B22")
			.Add("Fuchsia", "#FF00FF")
			.Add("Gainsboro", "#DCDCDC")
			.Add("GhostWhite", "#F8F8FF")
			.Add("Gold", "#FFD700")
			.Add("GoldenRod", "#DAA520")
			.Add("Gray", "#808080")
			.Add("Green", "#008000")
			.Add("GreenYellow", "#ADFF2F")
			.Add("HoneyDew", "#F0FFF0")
			.Add("HotPink", "#FF69B4")
			.Add("IndianRed ", "#CD5C5C")
			.Add("Indigo ", "#4B0082")
			.Add("Ivory", "#FFFFF0")
			.Add("Khaki", "#F0E68C")
			.Add("Lavender", "#E6E6FA")
			.Add("LavenderBlush", "#FFF0F5")
			.Add("LawnGreen", "#7CFC00")
			.Add("LemonChiffon", "#FFFACD")
			.Add("LightBlue", "#ADD8E6")
			.Add("LightCoral", "#F08080")
			.Add("LightCyan", "#E0FFFF")
			.Add("LightGoldenRodYellow", "#FAFAD2")
			.Add("LightGray", "#D3D3D3")
			.Add("LightGreen", "#90EE90")
			.Add("LightPink", "#FFB6C1")
			.Add("LightSalmon", "#FFA07A")
			.Add("LightSeaGreen", "#20B2AA")
			.Add("LightSkyBlue", "#87CEFA")
			.Add("LightSlateGray", "#778899")
			.Add("LightSteelBlue", "#B0C4DE")
			.Add("LightYellow", "#FFFFE0")
			.Add("Lime", "#00FF00")
			.Add("LimeGreen", "#32CD32")
			.Add("Linen", "#FAF0E6")
			.Add("Magenta", "#FF00FF")
			.Add("Maroon", "#800000")
			.Add("MediumAquaMarine", "#66CDAA")
			.Add("MediumBlue", "#0000CD")
			.Add("MediumOrchid", "#BA55D3")
			.Add("MediumPurple", "#9370D8")
			.Add("MediumSeaGreen", "#3CB371")
			.Add("MediumSlateBlue", "#7B68EE")
			.Add("MediumSpringGreen", "#00FA9A")
			.Add("MediumTurquoise", "#48D1CC")
			.Add("MediumVioletRed", "#C71585")
			.Add("MidnightBlue", "#191970")
			.Add("MintCream", "#F5FFFA")
			.Add("MistyRose", "#FFE4E1")
			.Add("Moccasin", "#FFE4B5")
			.Add("NavajoWhite", "#FFDEAD")
			.Add("Navy", "#000080")
			.Add("OldLace", "#FDF5E6")
			.Add("Olive", "#808000")
			.Add("OliveDrab", "#6B8E23")
			.Add("Orange", "#FFA500")
			.Add("OrangeRed", "#FF4500")
			.Add("Orchid", "#DA70D6")
			.Add("PaleGoldenRod", "#EEE8AA")
			.Add("PaleGreen", "#98FB98")
			.Add("PaleTurquoise", "#AFEEEE")
			.Add("PaleVioletRed", "#D87093")
			.Add("PapayaWhip", "#FFEFD5")
			.Add("PeachPuff", "#FFDAB9")
			.Add("Peru", "#CD853F")
			.Add("Pink", "#FFC0CB")
			.Add("Plum", "#DDA0DD")
			.Add("PowderBlue", "#B0E0E6")
			.Add("Purple", "#800080")
			.Add("Red", "#FF0000")
			.Add("RosyBrown", "#BC8F8F")
			.Add("RoyalBlue", "#4169E1")
			.Add("SaddleBrown", "#8B4513")
			.Add("Salmon", "#FA8072")
			.Add("SandyBrown", "#F4A460")
			.Add("SeaGreen", "#2E8B57")
			.Add("SeaShell", "#FFF5EE")
			.Add("Sienna", "#A0522D")
			.Add("Silver", "#C0C0C0")
			.Add("SkyBlue", "#87CEEB")
			.Add("SlateBlue", "#6A5ACD")
			.Add("SlateGray", "#708090")
			.Add("Snow", "#FFFAFA")
			.Add("SpringGreen", "#00FF7F")
			.Add("SteelBlue", "#4682B4")
			.Add("Tan", "#D2B48C")
			.Add("Teal", "#008080")
			.Add("Thistle", "#D8BFD8")
			.Add("Tomato", "#FF6347")
			.Add("Turquoise", "#40E0D0")
			.Add("Violet", "#EE82EE")
			.Add("Wheat", "#F5DEB3")
			.Add("White", "#FFFFFF")
			.Add("WhiteSmoke", "#F5F5F5")
			.Add("Yellow", "#FFFF00")
			.Add("YellowGreen", "#9ACD32")
			
			' put internal colors now
		End With
	End Sub
	
	
	'// properties do not use the internal dictionary
	Public ReadOnly Property AliceBlue() As Integer
		Get
			AliceBlue = 16775408
		End Get
	End Property
	Public ReadOnly Property AntiqueWhite() As Integer
		Get
			AntiqueWhite = 14150650
		End Get
	End Property
	Public ReadOnly Property Aqua() As Integer
		Get
			Aqua = 16776960
		End Get
	End Property
	Public ReadOnly Property Aquamarine() As Integer
		Get
			Aquamarine = 13959039
		End Get
	End Property
	Public ReadOnly Property Azure() As Integer
		Get
			Azure = 16777200
		End Get
	End Property
	Public ReadOnly Property Beige() As Integer
		Get
			Beige = 14480885
		End Get
	End Property
	Public ReadOnly Property Bisque() As Integer
		Get
			Bisque = 12903679
		End Get
	End Property
	Public ReadOnly Property Black() As Integer
		Get
			Black = 0
		End Get
	End Property
	Public ReadOnly Property BlanchedAlmond() As Integer
		Get
			BlanchedAlmond = 13495295
		End Get
	End Property
	Public ReadOnly Property Blue() As Integer
		Get
			Blue = 16711680
		End Get
	End Property
	Public ReadOnly Property BlueViolet() As Integer
		Get
			BlueViolet = 14822282
		End Get
	End Property
	Public ReadOnly Property Brown() As Integer
		Get
			Brown = 2763429
		End Get
	End Property
	Public ReadOnly Property BurlyWood() As Integer
		Get
			BurlyWood = 8894686
		End Get
	End Property
	Public ReadOnly Property CadetBlue() As Integer
		Get
			CadetBlue = 10526303
		End Get
	End Property
	Public ReadOnly Property Chartreuse() As Integer
		Get
			Chartreuse = 65407
		End Get
	End Property
	Public ReadOnly Property Chocolate() As Integer
		Get
			Chocolate = 1993170
		End Get
	End Property
	Public ReadOnly Property Coral() As Integer
		Get
			Coral = 5275647
		End Get
	End Property
	Public ReadOnly Property CornflowerBlue() As Integer
		Get
			CornflowerBlue = 15570276
		End Get
	End Property
	Public ReadOnly Property Cornsilk() As Integer
		Get
			Cornsilk = 14481663
		End Get
	End Property
	Public ReadOnly Property Crimson() As Integer
		Get
			Crimson = 3937500
		End Get
	End Property
	Public ReadOnly Property Cyan() As Integer
		Get
			Cyan = 16776960
		End Get
	End Property
	Public ReadOnly Property DarkBlue() As Integer
		Get
			DarkBlue = 9109504
		End Get
	End Property
	Public ReadOnly Property DarkCyan() As Integer
		Get
			DarkCyan = 9145088
		End Get
	End Property
	Public ReadOnly Property DarkGoldenRod() As Integer
		Get
			DarkGoldenRod = 755384
		End Get
	End Property
	Public ReadOnly Property DarkGray() As Integer
		Get
			DarkGray = 11119017
		End Get
	End Property
	Public ReadOnly Property DarkGreen() As Integer
		Get
			DarkGreen = 25600
		End Get
	End Property
	Public ReadOnly Property DarkKhaki() As Integer
		Get
			DarkKhaki = 7059389
		End Get
	End Property
	Public ReadOnly Property DarkMagenta() As Integer
		Get
			DarkMagenta = 9109643
		End Get
	End Property
	Public ReadOnly Property DarkOliveGreen() As Integer
		Get
			DarkOliveGreen = 3107669
		End Get
	End Property
	Public ReadOnly Property Darkorange() As Integer
		Get
			Darkorange = 36095
		End Get
	End Property
	Public ReadOnly Property DarkOrchid() As Integer
		Get
			DarkOrchid = 13382297
		End Get
	End Property
	Public ReadOnly Property DarkRed() As Integer
		Get
			DarkRed = 139
		End Get
	End Property
	Public ReadOnly Property DarkSalmon() As Integer
		Get
			DarkSalmon = 8034025
		End Get
	End Property
	Public ReadOnly Property DarkSeaGreen() As Integer
		Get
			DarkSeaGreen = 9419919
		End Get
	End Property
	Public ReadOnly Property DarkSlateBlue() As Integer
		Get
			DarkSlateBlue = 9125192
		End Get
	End Property
	Public ReadOnly Property DarkSlateGray() As Integer
		Get
			DarkSlateGray = 5197615
		End Get
	End Property
	Public ReadOnly Property DarkTurquoise() As Integer
		Get
			DarkTurquoise = 13749760
		End Get
	End Property
	Public ReadOnly Property DarkViolet() As Integer
		Get
			DarkViolet = 13828244
		End Get
	End Property
	Public ReadOnly Property DeepPink() As Integer
		Get
			DeepPink = 9639167
		End Get
	End Property
	Public ReadOnly Property DeepSkyBlue() As Integer
		Get
			DeepSkyBlue = 16760576
		End Get
	End Property
	Public ReadOnly Property DimGray() As Integer
		Get
			DimGray = 6908265
		End Get
	End Property
	Public ReadOnly Property DodgerBlue() As Integer
		Get
			DodgerBlue = 16748574
		End Get
	End Property
	Public ReadOnly Property FireBrick() As Integer
		Get
			FireBrick = 2237106
		End Get
	End Property
	Public ReadOnly Property FloralWhite() As Integer
		Get
			FloralWhite = 15792895
		End Get
	End Property
	Public ReadOnly Property ForestGreen() As Integer
		Get
			ForestGreen = 2263842
		End Get
	End Property
	Public ReadOnly Property Fuchsia() As Integer
		Get
			Fuchsia = 16711935
		End Get
	End Property
	Public ReadOnly Property Gainsboro() As Integer
		Get
			Gainsboro = 14474460
		End Get
	End Property
	Public ReadOnly Property GhostWhite() As Integer
		Get
			GhostWhite = 16775416
		End Get
	End Property
	Public ReadOnly Property Gold() As Integer
		Get
			Gold = 55295
		End Get
	End Property
	Public ReadOnly Property GoldenRod() As Integer
		Get
			GoldenRod = 2139610
		End Get
	End Property
	Public ReadOnly Property Gray() As Integer
		Get
			Gray = 8421504
		End Get
	End Property
	Public ReadOnly Property Green() As Integer
		Get
			Green = 32768
		End Get
	End Property
	Public ReadOnly Property GreenYellow() As Integer
		Get
			GreenYellow = 3145645
		End Get
	End Property
	Public ReadOnly Property HoneyDew() As Integer
		Get
			HoneyDew = 15794160
		End Get
	End Property
	Public ReadOnly Property HotPink() As Integer
		Get
			HotPink = 11823615
		End Get
	End Property
	Public ReadOnly Property IndianRed() As Integer
		Get
			IndianRed = 6053069
		End Get
	End Property
	Public ReadOnly Property Indigo() As Integer
		Get
			Indigo = 8519755
		End Get
	End Property
	Public ReadOnly Property Ivory() As Integer
		Get
			Ivory = 15794175
		End Get
	End Property
	Public ReadOnly Property Khaki() As Integer
		Get
			Khaki = 9234160
		End Get
	End Property
	Public ReadOnly Property Lavender() As Integer
		Get
			Lavender = 16443110
		End Get
	End Property
	Public ReadOnly Property LavenderBlush() As Integer
		Get
			LavenderBlush = 16118015
		End Get
	End Property
	Public ReadOnly Property LawnGreen() As Integer
		Get
			LawnGreen = 64636
		End Get
	End Property
	Public ReadOnly Property LemonChiffon() As Integer
		Get
			LemonChiffon = 13499135
		End Get
	End Property
	Public ReadOnly Property LightBlue() As Integer
		Get
			LightBlue = 15128749
		End Get
	End Property
	Public ReadOnly Property LightCoral() As Integer
		Get
			LightCoral = 8421616
		End Get
	End Property
	Public ReadOnly Property LightCyan() As Integer
		Get
			LightCyan = 16777184
		End Get
	End Property
	Public ReadOnly Property LightGoldenRodYellow() As Integer
		Get
			LightGoldenRodYellow = 13826810
		End Get
	End Property
	Public ReadOnly Property LightGray() As Integer
		Get
			LightGray = 13882323
		End Get
	End Property
	Public ReadOnly Property LightGreen() As Integer
		Get
			LightGreen = 9498256
		End Get
	End Property
	Public ReadOnly Property LightPink() As Integer
		Get
			LightPink = 12695295
		End Get
	End Property
	Public ReadOnly Property LightSalmon() As Integer
		Get
			LightSalmon = 8036607
		End Get
	End Property
	Public ReadOnly Property LightSeaGreen() As Integer
		Get
			LightSeaGreen = 11186720
		End Get
	End Property
	Public ReadOnly Property LightSkyBlue() As Integer
		Get
			LightSkyBlue = 16436871
		End Get
	End Property
	Public ReadOnly Property LightSlateGray() As Integer
		Get
			LightSlateGray = 10061943
		End Get
	End Property
	Public ReadOnly Property LightSteelBlue() As Integer
		Get
			LightSteelBlue = 14599344
		End Get
	End Property
	Public ReadOnly Property LightYellow() As Integer
		Get
			LightYellow = 14745599
		End Get
	End Property
	Public ReadOnly Property Lime() As Integer
		Get
			Lime = 65280
		End Get
	End Property
	Public ReadOnly Property LimeGreen() As Integer
		Get
			LimeGreen = 3329330
		End Get
	End Property
	Public ReadOnly Property Linen() As Integer
		Get
			Linen = 15134970
		End Get
	End Property
	Public ReadOnly Property Magenta() As Integer
		Get
			Magenta = 16711935
		End Get
	End Property
	Public ReadOnly Property Maroon() As Integer
		Get
			Maroon = 128
		End Get
	End Property
	Public ReadOnly Property MediumAquaMarine() As Integer
		Get
			MediumAquaMarine = 11193702
		End Get
	End Property
	Public ReadOnly Property MediumBlue() As Integer
		Get
			MediumBlue = 13434880
		End Get
	End Property
	Public ReadOnly Property MediumOrchid() As Integer
		Get
			MediumOrchid = 13850042
		End Get
	End Property
	Public ReadOnly Property MediumPurple() As Integer
		Get
			MediumPurple = 14184595
		End Get
	End Property
	Public ReadOnly Property MediumSeaGreen() As Integer
		Get
			MediumSeaGreen = 7451452
		End Get
	End Property
	Public ReadOnly Property MediumSlateBlue() As Integer
		Get
			MediumSlateBlue = 15624315
		End Get
	End Property
	Public ReadOnly Property MediumSpringGreen() As Integer
		Get
			MediumSpringGreen = 10156544
		End Get
	End Property
	Public ReadOnly Property MediumTurquoise() As Integer
		Get
			MediumTurquoise = 13422920
		End Get
	End Property
	Public ReadOnly Property MediumVioletRed() As Integer
		Get
			MediumVioletRed = 8721863
		End Get
	End Property
	Public ReadOnly Property MidnightBlue() As Integer
		Get
			MidnightBlue = 7346457
		End Get
	End Property
	Public ReadOnly Property MintCream() As Integer
		Get
			MintCream = 16449525
		End Get
	End Property
	Public ReadOnly Property MistyRose() As Integer
		Get
			MistyRose = 14804223
		End Get
	End Property
	Public ReadOnly Property Moccasin() As Integer
		Get
			Moccasin = 11920639
		End Get
	End Property
	Public ReadOnly Property NavajoWhite() As Integer
		Get
			NavajoWhite = 11394815
		End Get
	End Property
	Public ReadOnly Property Navy() As Integer
		Get
			Navy = 8388608
		End Get
	End Property
	Public ReadOnly Property OldLace() As Integer
		Get
			OldLace = 15136253
		End Get
	End Property
	Public ReadOnly Property Olive() As Integer
		Get
			Olive = 32896
		End Get
	End Property
	Public ReadOnly Property OliveDrab() As Integer
		Get
			OliveDrab = 2330219
		End Get
	End Property
	Public ReadOnly Property Orange() As Integer
		Get
			Orange = 42495
		End Get
	End Property
	Public ReadOnly Property OrangeRed() As Integer
		Get
			OrangeRed = 17919
		End Get
	End Property
	Public ReadOnly Property Orchid() As Integer
		Get
			Orchid = 14053594
		End Get
	End Property
	Public ReadOnly Property PaleGoldenRod() As Integer
		Get
			PaleGoldenRod = 11200750
		End Get
	End Property
	Public ReadOnly Property PaleGreen() As Integer
		Get
			PaleGreen = 10025880
		End Get
	End Property
	Public ReadOnly Property PaleTurquoise() As Integer
		Get
			PaleTurquoise = 15658671
		End Get
	End Property
	Public ReadOnly Property PaleVioletRed() As Integer
		Get
			PaleVioletRed = 9662680
		End Get
	End Property
	Public ReadOnly Property PapayaWhip() As Integer
		Get
			PapayaWhip = 14020607
		End Get
	End Property
	Public ReadOnly Property PeachPuff() As Integer
		Get
			PeachPuff = 12180223
		End Get
	End Property
	Public ReadOnly Property Peru() As Integer
		Get
			Peru = 4163021
		End Get
	End Property
	Public ReadOnly Property Pink() As Integer
		Get
			Pink = 13353215
		End Get
	End Property
	Public ReadOnly Property Plum() As Integer
		Get
			Plum = 14524637
		End Get
	End Property
	Public ReadOnly Property PowderBlue() As Integer
		Get
			PowderBlue = 15130800
		End Get
	End Property
	Public ReadOnly Property Purple() As Integer
		Get
			Purple = 8388736
		End Get
	End Property
	Public ReadOnly Property Red() As Integer
		Get
			Red = 255
		End Get
	End Property
	Public ReadOnly Property RosyBrown() As Integer
		Get
			RosyBrown = 9408444
		End Get
	End Property
	Public ReadOnly Property RoyalBlue() As Integer
		Get
			RoyalBlue = 14772545
		End Get
	End Property
	Public ReadOnly Property SaddleBrown() As Integer
		Get
			SaddleBrown = 1262987
		End Get
	End Property
	Public ReadOnly Property Salmon() As Integer
		Get
			Salmon = 7504122
		End Get
	End Property
	Public ReadOnly Property SandyBrown() As Integer
		Get
			SandyBrown = 6333684
		End Get
	End Property
	Public ReadOnly Property SeaGreen() As Integer
		Get
			SeaGreen = 5737262
		End Get
	End Property
	Public ReadOnly Property SeaShell() As Integer
		Get
			SeaShell = 15660543
		End Get
	End Property
	Public ReadOnly Property Sienna() As Integer
		Get
			Sienna = 2970272
		End Get
	End Property
	Public ReadOnly Property Silver() As Integer
		Get
			Silver = 12632256
		End Get
	End Property
	Public ReadOnly Property SkyBlue() As Integer
		Get
			SkyBlue = 15453831
		End Get
	End Property
	Public ReadOnly Property SlateBlue() As Integer
		Get
			SlateBlue = 13458026
		End Get
	End Property
	Public ReadOnly Property SlateGray() As Integer
		Get
			SlateGray = 9470064
		End Get
	End Property
	Public ReadOnly Property Snow() As Integer
		Get
			Snow = 16448255
		End Get
	End Property
	Public ReadOnly Property SpringGreen() As Integer
		Get
			SpringGreen = 8388352
		End Get
	End Property
	Public ReadOnly Property SteelBlue() As Integer
		Get
			SteelBlue = 11829830
		End Get
	End Property
	Public ReadOnly Property Tan() As Integer
		Get
			Tan = 9221330
		End Get
	End Property
	Public ReadOnly Property Teal() As Integer
		Get
			Teal = 8421376
		End Get
	End Property
	Public ReadOnly Property Thistle() As Integer
		Get
			Thistle = 14204888
		End Get
	End Property
	Public ReadOnly Property Tomato() As Integer
		Get
			Tomato = 4678655
		End Get
	End Property
	Public ReadOnly Property Turquoise() As Integer
		Get
			Turquoise = 13688896
		End Get
	End Property
	Public ReadOnly Property Violet() As Integer
		Get
			Violet = 15631086
		End Get
	End Property
	Public ReadOnly Property Wheat() As Integer
		Get
			Wheat = 11788021
		End Get
	End Property
	Public ReadOnly Property White() As Integer
		Get
			White = 16777215
		End Get
	End Property
	Public ReadOnly Property WhiteSmoke() As Integer
		Get
			WhiteSmoke = 16119285
		End Get
	End Property
	Public ReadOnly Property Yellow() As Integer
		Get
			Yellow = 65535
		End Get
	End Property
	Public ReadOnly Property YellowGreen() As Integer
		Get
			YellowGreen = 3329434
		End Get
	End Property
	
	Public ReadOnly Property ChannelLabelBack() As Integer
		Get
			ChannelLabelBack = FormColors.ChannelLabelBack
		End Get
	End Property
	Public ReadOnly Property ChannelLabelText() As Integer
		Get
			ChannelLabelText = FormColors.ChannelLabelText
		End Get
	End Property
	Public ReadOnly Property ChannelListBack() As Integer
		Get
			ChannelListBack = FormColors.ChannelListBack
		End Get
	End Property
	Public ReadOnly Property ChannelListNormal() As Integer
		Get
			ChannelListNormal = FormColors.ChannelListText
		End Get
	End Property
	Public ReadOnly Property ChannelListSelf() As Integer
		Get
			ChannelListSelf = FormColors.ChannelListSelf
		End Get
	End Property
	Public ReadOnly Property ChannelListIdle() As Integer
		Get
			ChannelListIdle = FormColors.ChannelListIdle
		End Get
	End Property
	Public ReadOnly Property ChannelListSquelched() As Integer
		Get
			ChannelListSquelched = FormColors.ChannelListSquelched
		End Get
	End Property
	Public ReadOnly Property ChannelListOps() As Integer
		Get
			ChannelListOps = FormColors.ChannelListOps
		End Get
	End Property
	Public ReadOnly Property RTBBack() As Integer
		Get
			RTBBack = FormColors.RTBBack
		End Get
	End Property
	Public ReadOnly Property SendBoxesBack() As Integer
		Get
			SendBoxesBack = FormColors.SendBoxesBack
		End Get
	End Property
	Public ReadOnly Property SendBoxesText() As Integer
		Get
			SendBoxesText = FormColors.SendBoxesText
		End Get
	End Property
	
	Public ReadOnly Property Carats() As Integer
		Get
			Carats = RTBColors.Carats
		End Get
	End Property
	Public ReadOnly Property ConsoleText() As Integer
		Get
			ConsoleText = RTBColors.ConsoleText
		End Get
	End Property
	Public ReadOnly Property EmoteText() As Integer
		Get
			EmoteText = RTBColors.EmoteText
		End Get
	End Property
	Public ReadOnly Property EmoteUsernames() As Integer
		Get
			EmoteUsernames = RTBColors.EmoteUsernames
		End Get
	End Property
	Public ReadOnly Property ErrorMessageText() As Integer
		Get
			ErrorMessageText = RTBColors.ErrorMessageText
		End Get
	End Property
	Public ReadOnly Property InformationText() As Integer
		Get
			InformationText = RTBColors.InformationText
		End Get
	End Property
	Public ReadOnly Property JoinedChannelName() As Integer
		Get
			JoinedChannelName = RTBColors.JoinedChannelName
		End Get
	End Property
	Public ReadOnly Property JoinedChannelText() As Integer
		Get
			JoinedChannelText = RTBColors.JoinedChannelText
		End Get
	End Property
	Public ReadOnly Property JoinText() As Integer
		Get
			JoinText = RTBColors.JoinText
		End Get
	End Property
	Public ReadOnly Property JoinUsername() As Integer
		Get
			JoinUsername = RTBColors.JoinUsername
		End Get
	End Property
	Public ReadOnly Property ServerInfoText() As Integer
		Get
			ServerInfoText = RTBColors.ServerInfoText
		End Get
	End Property
	Public ReadOnly Property SuccessText() As Integer
		Get
			SuccessText = RTBColors.SuccessText
		End Get
	End Property
	Public ReadOnly Property TalkBotUsername() As Integer
		Get
			TalkBotUsername = RTBColors.TalkBotUsername
		End Get
	End Property
	Public ReadOnly Property TalkNormalText() As Integer
		Get
			TalkNormalText = RTBColors.TalkNormalText
		End Get
	End Property
	Public ReadOnly Property TalkUsernameNormal() As Integer
		Get
			TalkUsernameNormal = RTBColors.TalkUsernameNormal
		End Get
	End Property
	Public ReadOnly Property TalkUsernameOp() As Integer
		Get
			TalkUsernameOp = RTBColors.TalkUsernameOp
		End Get
	End Property
	Public ReadOnly Property TimeStamps() As Integer
		Get
			TimeStamps = RTBColors.TimeStamps
		End Get
	End Property
	Public ReadOnly Property WhisperCarats() As Integer
		Get
			WhisperCarats = RTBColors.WhisperCarats
		End Get
	End Property
	Public ReadOnly Property WhisperText() As Integer
		Get
			WhisperText = RTBColors.WhisperText
		End Get
	End Property
	Public ReadOnly Property WhisperUsernames() As Integer
		Get
			WhisperUsernames = RTBColors.WhisperUsernames
		End Get
	End Property
End Class