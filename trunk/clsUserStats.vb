Option Strict Off
Option Explicit On
Friend Class clsUserStats
	' clsUserStats.cls
	' Copyright (C) 2008 Eric Evans
	
	
	' variables
	Private m_stat_string As String
	Private m_Game As String
	Private m_icon As String
	Private m_spawn As Boolean
	Private m_clan As String
	Private m_level As Integer
	Private m_character_name As String
	Private m_character_class_id As Short
	Private m_character_flags As Integer
	Private m_acts_completed As Integer
	Private m_character_ladder As Boolean
	Private m_wins As Integer
	Private m_ladder_rating As Integer
	Private m_high_rating As Integer
	Private m_ladder_rank As Integer
	Private m_strength As Integer
	Private m_dexterity As Integer
	Private m_vitality As Integer
	Private m_gold As Integer
	Private m_magic As Integer
	Private m_dots As Integer
	Private m_expansion As Boolean
	Private m_hardcore As Boolean
	Private m_realm As String
	Private m_are_stats_valid As Boolean
	
	
	Public Property Game() As String
		Get
			Game = m_Game
		End Get
		Set(ByVal Value As String)
			m_Game = Value
		End Set
	End Property
	
	
	Public Property IsValid() As Boolean
		Get
			IsValid = m_are_stats_valid
		End Get
		Set(ByVal Value As Boolean)
			m_are_stats_valid = Value
		End Set
	End Property
	
	Public ReadOnly Property IsWCG() As Boolean
		Get
			Select Case (StrReverse(Icon))
				Case "WCRF" : IsWCG = True
				Case "WCPL" : IsWCG = True
				Case "WCGO" : IsWCG = True
				Case "WCSI" : IsWCG = True
				Case "WCBR" : IsWCG = True
				Case "WCPG" : IsWCG = True
			End Select
		End Get
	End Property
	
	
	Public Property Icon() As String
		Get
			Icon = m_icon
		End Get
		Set(ByVal Value As String)
			m_icon = Value
		End Set
	End Property
	
	Public ReadOnly Property IconTier() As String
		Get
			If ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP)) Then
				Select Case (Mid(Icon, 2, 1))
					Case "H" : IconTier = "Human"
					Case "N" : IconTier = "Night Elf"
					Case "U" : IconTier = "Undead"
					Case "O" : IconTier = "Orc"
					Case "R" : IconTier = "Random"
					Case "D" : IconTier = "Tournament"
					Case Else : IconTier = Mid(Icon, 2, 1)
				End Select
			End If
		End Get
	End Property
	
	Public ReadOnly Property IconName() As String
		Get
			If ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP)) Then
				Select Case (StrReverse(Icon))
					Case "WCRF" : IconName = "WCG Referee"
					Case "WCPL" : IconName = "WCG Player"
					Case "WCGO" : IconName = "WCG Gold Medalist"
					Case "WCSI" : IconName = "WCG Silver Medalist"
					Case "WCBR" : IconName = "WCG Bronze Medalist"
					Case "WCPG" : IconName = "WCG Professional Gamer"
				End Select
				
				If (Len(IconName) > 0) Then
					Exit Property
				End If
				
				Select Case (Mid(Icon, 2, 1))
					Case "H" ' Human
						If (Game = PRODUCT_WAR3) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "footman"
								Case 3 : IconName = "knight"
								Case 4 : IconName = "Archmage"
								Case 5 : IconName = "Medivh"
							End Select
						Else
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "rifleman"
								Case 3 : IconName = "sorceress"
								Case 4 : IconName = "spellbreaker"
								Case 5 : IconName = "Blood Mage"
								Case 6 : IconName = "Jaina Proudmore"
							End Select
						End If
						
					Case "N" ' Night Elf
						If (Game = PRODUCT_WAR3) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "archer"
								Case 3 : IconName = "druid of the claw"
								Case 4 : IconName = "Priestess of the Moon"
								Case 5 : IconName = "Furion"
							End Select
						Else
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "huntress"
								Case 3 : IconName = "druid of the talon"
								Case 4 : IconName = "dryad"
								Case 5 : IconName = "Keeper of the Grove"
								Case 6 : IconName = "Maiev"
							End Select
						End If
						
					Case "U" ' Undead
						If (Game = PRODUCT_WAR3) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "ghoul"
								Case 3 : IconName = "abomination"
								Case 4 : IconName = "Lich"
								Case 5 : IconName = "Tichondrius"
							End Select
						Else
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "crypt fiend"
								Case 3 : IconName = "banshee"
								Case 4 : IconName = "destroyer"
								Case 5 : IconName = "Crypt Lord"
								Case 6 : IconName = "Sylvanas"
							End Select
						End If
						
					Case "O" ' Orc
						If (Game = PRODUCT_WAR3) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "grunt"
								Case 3 : IconName = "tauren"
								Case 4 : IconName = "Far Seer"
								Case 5 : IconName = "Thrall"
							End Select
						Else
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "headhunter"
								Case 3 : IconName = "shaman"
								Case 4 : IconName = "Spirit Walker"
								Case 5 : IconName = "Shadow Hunter"
								Case 6 : IconName = "Rexxar"
							End Select
						End If
						
					Case "R" ' Random
						If (Game = PRODUCT_WAR3) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "dragon whelp"
								Case 3 : IconName = "blue dragon"
								Case 4 : IconName = "red dragon"
								Case 5 : IconName = "Deathwing"
							End Select
						Else
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "myrmidon"
								Case 3 : IconName = "siren"
								Case 4 : IconName = "dragon turtle"
								Case 5 : IconName = "sea witch"
								Case 6 : IconName = "Illidan"
							End Select
						End If
						
					Case "D" ' Tournament
						If (Game = PRODUCT_W3XP) Then
							Select Case (Val(Mid(Icon, 1, 1)))
								Case 1 : IconName = "peon"
								Case 2 : IconName = "Felguard"
								Case 3 : IconName = "infernal"
								Case 4 : IconName = "doomguard"
								Case 5 : IconName = "pit lord"
								Case 6 : IconName = "Archimonde"
							End Select
						End If
				End Select
				
				If (IconName = vbNullString) Then
					IconName = "unknown"
				End If
			End If
		End Get
	End Property
	
	
	Public Property Statstring() As String
		Get
			Statstring = m_stat_string
		End Get
		Set(ByVal Value As String)
			If (Len(Value) < 4) Then
				Exit Property
			End If
			Game = StrReverse(Left(Value, 4))
			If (Len(Value) >= 4) Then
				Value = Mid(Value, 5)
				
				If (Left(Value, 1) = " ") Then
					Value = Mid(Value, 2)
				End If
				
				m_stat_string = Value
				
				Call Parse()
			End If
		End Set
	End Property
	
	
	Public Property IsSpawn() As Boolean
		Get
			IsSpawn = m_spawn
		End Get
		Set(ByVal Value As Boolean)
			m_spawn = Value
		End Set
	End Property
	
	
	Public Property Clan() As String
		Get
			Clan = m_clan
		End Get
		Set(ByVal Value As String)
			m_clan = Value
		End Set
	End Property
	
	
	Public Property Wins() As Integer
		Get
			Wins = m_wins
		End Get
		Set(ByVal Value As Integer)
			m_wins = Value
		End Set
	End Property
	
	
	Public Property LadderRating() As Integer
		Get
			LadderRating = m_ladder_rating
		End Get
		Set(ByVal Value As Integer)
			m_ladder_rating = Value
		End Set
	End Property
	
	
	Public Property HighLadderRating() As Integer
		Get
			HighLadderRating = m_high_rating
		End Get
		Set(ByVal Value As Integer)
			m_high_rating = Value
		End Set
	End Property
	
	
	Public Property LadderRank() As Integer
		Get
			LadderRank = m_ladder_rank
		End Get
		Set(ByVal Value As Integer)
			m_ladder_rank = Value
		End Set
	End Property
	
	
	Public Property Level() As Integer
		Get
			Level = m_level
		End Get
		Set(ByVal Value As Integer)
			m_level = Value
		End Set
	End Property
	
	
	Public Property CharacterName() As String
		Get
			CharacterName = m_character_name
		End Get
		Set(ByVal Value As String)
			m_character_name = Value
		End Set
	End Property
	
	
	Public Property CharacterClassID() As Short
		Get
			CharacterClassID = m_character_class_id
		End Get
		Set(ByVal Value As Short)
			m_character_class_id = Value
		End Set
	End Property
	
	Public ReadOnly Property CharacterClass() As String
		Get
			On Error GoTo ERROR_HANDLER
			
			Dim DIClasses(3) As String
			Dim DIIClasses(7) As String
			If ((Game = PRODUCT_DSHR) Or (Game = PRODUCT_DRTL)) Then
				
				DIClasses(0) = "Warrior"
				DIClasses(1) = "Rogue"
				DIClasses(2) = "Sorceror"
				
				If (CharacterClassID < UBound(DIClasses)) Then
					CharacterClass = DIClasses(CharacterClassID)
				End If
				
			ElseIf ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then 
				
				
				DIIClasses(0) = "Amazon"
				DIIClasses(1) = "Sorceress"
				DIIClasses(2) = "Necromancer"
				DIIClasses(3) = "Paladin"
				DIIClasses(4) = "Barbarian"
				DIIClasses(5) = "Druid"
				DIIClasses(6) = "Assassin"
				
				If (CharacterClassID - 1 < UBound(DIIClasses)) Then
					CharacterClass = DIIClasses(CharacterClassID - 1)
				End If
			End If
			
			Exit Property
			
ERROR_HANDLER: 
			IsValid = False
			Exit Property
		End Get
	End Property
	
	
	Public Property CharacterFlags() As Integer
		Get
			CharacterFlags = m_character_flags
		End Get
		Set(ByVal Value As Integer)
			m_character_flags = Value
		End Set
	End Property
	
	Public ReadOnly Property IsHardcoreCharacter() As Boolean
		Get
			IsHardcoreCharacter = ((CharacterFlags And &H4) = &H4)
		End Get
	End Property
	
	Public ReadOnly Property IsFemaleCharacter() As Boolean
		Get
			IsFemaleCharacter = ((m_character_class_id = 1) Or (m_character_class_id = 2) Or (m_character_class_id = 7))
		End Get
	End Property
	
	Public ReadOnly Property IsExpansionCharacter() As Boolean
		Get
			If (StrReverse(Game) = PRODUCT_D2XP) Then
				IsExpansionCharacter = True
			Else
				IsExpansionCharacter = ((CharacterFlags And &H20) = &H20)
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsCharacterDead() As Boolean
		Get
			IsCharacterDead = ((IsHardcoreCharacter) And ((CharacterFlags And &H8) = &H8))
		End Get
	End Property
	
	
	Public Property IsLadderCharacter() As Boolean
		Get
			IsLadderCharacter = m_character_ladder
		End Get
		Set(ByVal Value As Boolean)
			m_character_ladder = Value
		End Set
	End Property
	
	
	Public Property ActsCompleted() As Short
		Get
			ActsCompleted = m_acts_completed
		End Get
		Set(ByVal Value As Short)
			m_acts_completed = Value
		End Set
	End Property
	
	Public ReadOnly Property CurrentAct() As Short
		Get
			If ((Game() = PRODUCT_D2DV) Or (Game() = PRODUCT_D2XP)) Then
				If (IsExpansionCharacter() = False) Then
					CurrentAct = ((ActsCompleted() Mod 4) + 1)
				Else
					CurrentAct = ((ActsCompleted() Mod 5) + 1)
				End If
			End If
		End Get
	End Property
	
	
	Public ReadOnly Property CurrentActName() As String
		Get
			Select Case CurrentAct()
				Case 1 : CurrentActName = "Act I"
				Case 2 : CurrentActName = "Act II"
				Case 3 : CurrentActName = "Act III"
				Case 4
					If IsExpansionCharacter() Then
						CurrentActName = "Act IV/V" ' either IV or V
					Else
						CurrentActName = "Act IV"
					End If
				Case 5 : CurrentActName = "Act V" ' this is never seen, but would be proper value for Act V
			End Select
		End Get
	End Property
	
	Public ReadOnly Property CurrentActAndDifficulty() As String
		Get
			If CurrentDifficultyID() >= 3 Then
				CurrentActAndDifficulty = CurrentDifficulty()
			Else
				CurrentActAndDifficulty = CurrentActName() & " of " & CurrentDifficulty()
			End If
		End Get
	End Property
	
	Public ReadOnly Property CurrentDifficulty() As String
		Get
			If ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
				Select Case (CurrentDifficultyID)
					Case 0 : CurrentDifficulty = "Normal"
					Case 1 : CurrentDifficulty = "Nightmare"
					Case 2 : CurrentDifficulty = "Hell"
					Case 3 : CurrentDifficulty = "All Acts"
				End Select
			End If
		End Get
	End Property
	
	Public ReadOnly Property CurrentDifficultyID() As Short
		Get
			Dim difficulty As Short
			
			If ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
				If (IsExpansionCharacter = False) Then
					difficulty = Fix(ActsCompleted / 4)
				Else
					difficulty = Fix(ActsCompleted / 5)
				End If
			End If
			CurrentDifficultyID = difficulty
		End Get
	End Property
	
	Public ReadOnly Property CharacterTitle() As String
		Get
			On Error GoTo ERROR_HANDLER
			
			' thanks c0ol for multi-dimensional array idea
			Dim Classic(2, 3, 2) As String
			Dim Expansion(2, 3, 2) As String
			
			' softcore
			Classic(0, 0, 0) = "Sir"
			Classic(0, 0, 1) = "Dame"
			Classic(0, 1, 0) = "Lord"
			Classic(0, 1, 1) = "Lady"
			Classic(0, 2, 0) = "Baron"
			Classic(0, 2, 1) = "Baroness"
			
			' hardcore
			Classic(1, 0, 0) = "Count"
			Classic(1, 0, 1) = "Countess"
			Classic(1, 1, 0) = "Duke"
			Classic(1, 1, 1) = "Duchess"
			Classic(1, 2, 0) = "King"
			Classic(1, 2, 1) = "Queen"
			
			' softcore
			Expansion(0, 0, 0) = "Slayer"
			Expansion(0, 1, 0) = "Champion"
			Expansion(0, 2, 0) = "Patriarch"
			Expansion(0, 2, 1) = "Matriarch"
			
			' hardcore
			Expansion(1, 0, 0) = "Destroyer"
			Expansion(1, 1, 0) = "Conquerer"
			Expansion(1, 2, 0) = "Guardian"
			
			If ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
				If (CurrentDifficultyID >= 1) Then
					If (IsExpansionCharacter = False) Then
						CharacterTitle = Classic(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, IIf(IsFemaleCharacter, 1, 0))
						
						' run again with is female as false
						If (CharacterTitle = vbNullString) Then
							CharacterTitle = Classic(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, 0)
						End If
					Else
						CharacterTitle = Expansion(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, IIf(IsFemaleCharacter, 1, 0))
						
						' run again with is female as false
						If (CharacterTitle = vbNullString) Then
							CharacterTitle = Expansion(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, 0)
						End If
					End If
				End If
			End If
			
			Exit Property
			
ERROR_HANDLER: 
			CharacterTitle = vbNullString
			
			Exit Property
		End Get
	End Property
	
	Public ReadOnly Property CharacterTitleAndName() As String
		Get
			CharacterTitleAndName = CharacterTitle()
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LenB(CharacterTitleAndName) > 0) Then CharacterTitleAndName = CharacterTitleAndName & " "
			CharacterTitleAndName = CharacterTitleAndName & CharacterName()
		End Get
	End Property
	
	
	Public Property Dots() As Integer
		Get
			Dots = m_dots
		End Get
		Set(ByVal Value As Integer)
			m_dots = Value
		End Set
	End Property
	
	
	Public Property Strength() As Integer
		Get
			Strength = m_strength
		End Get
		Set(ByVal Value As Integer)
			m_strength = Value
		End Set
	End Property
	
	
	Public Property Magic() As Integer
		Get
			Magic = m_magic
		End Get
		Set(ByVal Value As Integer)
			m_magic = Value
		End Set
	End Property
	
	
	Public Property Gold() As Integer
		Get
			Gold = m_gold
		End Get
		Set(ByVal Value As Integer)
			m_gold = Value
		End Set
	End Property
	
	
	Public Property Dexterity() As Integer
		Get
			Dexterity = m_dexterity
		End Get
		Set(ByVal Value As Integer)
			m_dexterity = Value
		End Set
	End Property
	
	
	Public Property Vitality() As Integer
		Get
			Vitality = m_vitality
		End Get
		Set(ByVal Value As Integer)
			m_vitality = Value
		End Set
	End Property
	
	
	Public Property Realm() As String
		Get
			Realm = m_realm
		End Get
		Set(ByVal Value As String)
			m_realm = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: ToString was upgraded to ToString_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public ReadOnly Property ToString_Renamed() As String
		Get
			Dim buf As String
			
			buf = GetProductInfo(Game).FullName
			
			Select Case (Game)
				Case PRODUCT_SSHR, PRODUCT_STAR, PRODUCT_JSTR, PRODUCT_SEXP
					buf = buf & StarCraft_ToString()
					
				Case PRODUCT_DSHR, PRODUCT_DRTL
					buf = buf & Diablo_ToString()
					
				Case PRODUCT_D2DV, PRODUCT_D2XP
					buf = buf & DiabloII_ToString()
					
				Case PRODUCT_W2BN
					buf = buf & WarCraftII_ToString()
					
				Case PRODUCT_WAR3, PRODUCT_W3XP
					buf = buf & WarCraftIII_ToString()
			End Select
			
			'buf = buf & "."
			
			ToString_Renamed = buf
		End Get
	End Property
	
	Public ReadOnly Property IconCode() As Short
		Get
			Dim intWins As Integer
			Dim strIcon As String
			
			If (BotVars.ShowStatsIcons) Then
				strIcon = StrReverse(Icon())
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If (IsValid And LenB(Statstring) > 0) Then
					Select Case strIcon
						Case vbNullString, Game ' icon is current product or icon is not set
						Case Else ' if icon field is present and custom
							' display icon
							IconCode = GetIconImageListPosition(strIcon)
							Exit Property
					End Select
					
					Select Case Game
						Case PRODUCT_DRTL, PRODUCT_DSHR
							' display icon based on D1 class and number of dots
							IconCode = (ICON_START_D1 + (CharacterClassID * 4) + Dots)
							
							' if spawn flag, use DSHR warrior
							If IsSpawn Then IconCode = IC_DIAB_SPAWN
							
						Case PRODUCT_D2DV, PRODUCT_D2XP
							' display icon based on D2 class
							IconCode = (ICON_START_D2 + CharacterClassID - 1)
							
						Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_SSHR, PRODUCT_JSTR, PRODUCT_W2BN
							' display icon based on wins
							intWins = Wins
							If (intWins > 10) Then intWins = 10
							
							' choose starting point
							If Game = PRODUCT_W2BN Then
								IconCode = CShort(ICON_START_W2 + intWins)
							Else
								IconCode = CShort(ICON_START_SC + intWins)
							End If
							
							' if spawn flag, use spawn icon
							If IsSpawn Then
								Select Case Game
									Case PRODUCT_W2BN : IconCode = IC_W2BN_SPAWN
									Case PRODUCT_JSTR : IconCode = IC_JSTR_SPAWN
									Case Else : IconCode = IC_STAR_SPAWN
								End Select
							End If
							
					End Select
				End If
			End If
			
			If ((IconCode = 0) Or (IconCode = ICUNKNOWN)) Then
				IconCode = GetIconImageListPosition(Game())
			End If
		End Get
	End Property
	
	Private Function GetIconImageListPosition(ByVal Icon As String) As Short
		Dim IconCode As Short
		Dim IMLSeq As Short
		Dim PerTier As Short
		
		IconCode = ICUNKNOWN
		
		Select Case Icon
			Case PRODUCT_CHAT : IconCode = ICCHAT
			Case PRODUCT_SSHR : IconCode = ICSCSW
			Case PRODUCT_STAR : IconCode = ICSTAR
			Case PRODUCT_JSTR : IconCode = ICJSTR
			Case PRODUCT_SEXP : IconCode = ICSEXP
			Case PRODUCT_DSHR : IconCode = ICDIABLOSW
			Case PRODUCT_DRTL : IconCode = ICDIABLO
			Case PRODUCT_W2BN : IconCode = ICW2BN
			Case PRODUCT_D2DV : IconCode = ICD2DV
			Case PRODUCT_D2XP : IconCode = ICD2XP
			Case PRODUCT_WAR3 : IconCode = ICWAR3
			Case PRODUCT_W3XP : IconCode = ICWAR3X
				
			Case "WCRF" : IconCode = IC_WCRF
			Case "WCPL" : IconCode = IC_WCPL
			Case "WCGO" : IconCode = IC_WCGO
			Case "WCSI" : IconCode = IC_WCSI
			Case "WCBR" : IconCode = IC_WCBR
			Case "WCPG" : IconCode = IC_WCPG
				
			Case Else : IconCode = ICUNKNOWN
		End Select
		
		If IconCode <> ICUNKNOWN Then
			' special named icon
			GetIconImageListPosition = IconCode
		ElseIf (StrComp(Left(Icon, 2), "W3", CompareMethod.Binary) = 0) Then 
			' W3 stats icon: choose starting point and per-tier
			If (Game = PRODUCT_WAR3) Then
				IMLSeq = ICON_START_WAR3
				PerTier = 5
			Else
				IMLSeq = ICON_START_W3XP
				PerTier = 6
			End If
			
			' choose race
			Select Case Mid(Icon, 3, 1)
				Case "H"
					IconCode = IMLSeq + 1
				Case "N"
					IconCode = IMLSeq + 1 + (PerTier * 1)
				Case "U"
					IconCode = IMLSeq + 1 + (PerTier * 2)
				Case "O"
					IconCode = IMLSeq + 1 + (PerTier * 3)
				Case "R"
					IconCode = IMLSeq + 1 + (PerTier * 4)
				Case "T", "D"
					IconCode = IMLSeq + 1 + (PerTier * 5)
				Case Else 'starting point is "unknown icon" icon
					IconCode = IMLSeq
			End Select
			
			' if race known, add tier to it
			' otherwise, defaults to IMLPos 1, which is W3 unknown icon (but an icon is present)
			If StrictIsNumeric(Mid(Icon, 4, 1)) And IconCode > IMLSeq Then
				IconCode = IconCode + (CShort(Mid(Icon, 4, 1)) - 1)
			End If
		Else
			' unknown icon - use "unknown icon" icon
			IconCode = ICON_START_WAR3
		End If
		
		GetIconImageListPosition = IconCode
	End Function
	
	Private Sub Parse()
		On Error GoTo ERROR_HANDLER
		
		' empty stats are invalid, unless a specific product says otherwise
		IsValid = False
		
		Select Case (Game)
			Case PRODUCT_SSHR, PRODUCT_STAR, PRODUCT_JSTR, PRODUCT_SEXP
				Call ParseStarCraft()
				
			Case PRODUCT_DSHR, PRODUCT_DRTL
				Call ParseDiablo()
				
			Case PRODUCT_D2DV, PRODUCT_D2XP
				Call ParseDiabloII()
				
			Case PRODUCT_W2BN
				Call ParseWarCraftII()
				
			Case PRODUCT_WAR3, PRODUCT_W3XP
				Call ParseWarCraftIII()
		End Select
		
		Exit Sub
		
ERROR_HANDLER: 
		IsValid = False
		
		Exit Sub
	End Sub
	
	Private Sub ParseStarCraft()
		Dim Values() As String
		Dim lngIsSpawn As Integer
		
		If (Statstring = vbNullString) Then
			Exit Sub
		End If
		
		IsValid = True
		
		ReDim Preserve Values(0)
		Values = Split(Statstring, Space(1))
		
		If (UBound(Values) >= 4) Then
			LadderRating = Val(Values(0))
			LadderRank = Val(Values(1))
			Wins = Val(Values(2))
			lngIsSpawn = Val(Values(3))
			IsSpawn = CBool(lngIsSpawn)
			
			If (UBound(Values) >= 8) Then
				Icon = CStr(Values(8))
			End If
		Else
			IsValid = False
		End If
		
		If lngIsSpawn < 0 Or lngIsSpawn > 1 Then
			IsValid = False
		End If
	End Sub
	
	Private Sub ParseDiablo()
		Dim Values() As String
		Dim lngIsSpawn As Integer
		
		If (Statstring = vbNullString) Then
			Exit Sub
		End If
		
		IsValid = True
		
		ReDim Preserve Values(0)
		Values = Split(Statstring, Space(1))
		
		If (UBound(Values) <> 8) Then
			IsValid = False
			
			Exit Sub
		End If
		
		Level = Val(Values(0))
		CharacterClassID = Val(Values(1))
		Dots = Val(Values(2))
		Strength = Val(Values(3))
		Magic = Val(Values(4))
		Dexterity = Val(Values(5))
		Vitality = Val(Values(6))
		Gold = Val(Values(7))
		lngIsSpawn = Val(Values(8))
		IsSpawn = CBool(lngIsSpawn)
		
		If CharacterClassID < 0 Or CharacterClassID > 2 Or Dots < 0 Or Dots > 3 Or lngIsSpawn < 0 Or lngIsSpawn > 1 Then
			IsValid = False
		End If
	End Sub
	
	Private Sub ParseDiabloII()
		Dim Values() As String
		Dim charData() As Short
		
		' empty D2 stats is valid ("Open character")
		IsValid = True
		
		If (Statstring = vbNullString) Then
			Exit Sub
		End If
		
		ReDim Preserve Values(0)
		Values = Split(Statstring, ",", 3)
		
		If (UBound(Values) >= 1) Then
			Realm = Values(0)
			CharacterName = Values(1)
			
			If (UBound(Values) >= 2) Then
				MakeArr(Values(2), charData)
				
				If (UBound(charData) >= 13) Then
					CharacterClassID = charData(13)
					If (UBound(charData) >= 27) Then
						Level = charData(25)
						CharacterFlags = charData(26)
						ActsCompleted = (CShort(charData(27) Xor &H80) / 2)
						
						If (UBound(charData) >= 30) Then
							IsLadderCharacter = (charData(30) <> &HFF)
						End If
					End If
				End If
			End If
		End If
		
		If CharacterClassID > 7 Or CharacterClassID < 0 Or CurrentDifficultyID > 3 Or CurrentDifficultyID < 0 Then
			IsValid = False
		End If
	End Sub
	
	Private Sub ParseWarCraftII()
		Call ParseStarCraft()
	End Sub
	
	Private Sub ParseWarCraftIII()
		Dim Values() As String
		
		If (Statstring = vbNullString) Then
			Exit Sub
		End If
		
		IsValid = True
		
		ReDim Preserve Values(0)
		Values = Split(Statstring, Space(1))
		
		If (UBound(Values) = 0) Then
			Level = Val(Values(0))
			Exit Sub
		End If
		
		Level = Val(Values(1))
		Icon = Values(0)
		
		If (UBound(Values) > 1) Then
			Clan = StrReverse(Values(2))
		End If
	End Sub
	
	Private Function StarCraft_ToString() As String
		Dim buf As String
		
		If (Statstring = vbNullString) Then
			buf = " (No stats available)"
		ElseIf (IsValid = False) Then 
			buf = " (Unrecognized stats)"
		Else
			buf = " (" & Wins() & " wins" & IIf(LadderRating(), ", with a rating of " & LadderRating() & " on the ladder", "") & ")"
			
			If (IsSpawn()) Then
				buf = buf & " (spawn)"
			End If
		End If
		
		StarCraft_ToString = buf
	End Function
	
	Private Function Diablo_ToString() As String
		Dim buf As String
		
		If (Statstring = vbNullString) Then
			buf = " (No stats available)"
		ElseIf (IsValid = False) Then 
			buf = " (Unrecognized stats)"
		Else
			buf = " (Level " & Level() & " " & CharacterClass() & " with " & Dots() & " dots, " & Strength() & " strength, " & Magic() & " magic, " & Dexterity() & " dexterity, " & Vitality() & " vitality, and " & Gold() & " gold)"
			
			If (IsSpawn()) Then
				buf = buf & " (spawn)"
			End If
		End If
		
		Diablo_ToString = buf
	End Function
	
	Private Function DiabloII_ToString() As String
		Dim buf As String
		
		If (Statstring = vbNullString) Then
			buf = " (Open character)"
		ElseIf (IsValid = False) Then 
			buf = " (Unrecognized stats)"
		Else
			buf = " ("
			
			buf = buf & CharacterTitleAndName() & ", a " & IIf(IsCharacterDead(), "dead ", vbNullString) & IIf(IsHardcoreCharacter(), "hardcore ", vbNullString) & "level " & Level() & " " & IIf(IsLadderCharacter(), "ladder ", vbNullString) & CharacterClass()
			
			buf = buf & " on Realm " & Realm() & ")"
		End If
		
		DiabloII_ToString = buf
	End Function
	
	Private Function WarCraftII_ToString() As String
		WarCraftII_ToString = StarCraft_ToString()
	End Function
	
	Private Function WarCraftIII_ToString() As String
		Dim buf As String
		
		If (Statstring = vbNullString) Then
			buf = " (No stats available)"
		ElseIf (IsValid = False) Then 
			buf = " (Unrecognized stats)"
		Else
			buf = " (Level " & Level()
			If (IsWCG) Then
				buf = buf & ", " & IconName() & " icon"
			ElseIf (Icon() <> vbNullString) Then 
				buf = buf & ", icon tier " & IconTier() & ", " & IconName() & " icon"
			End If
			
			If (Clan() <> vbNullString) Then
				buf = buf & ", in Clan " & Clan()
			End If
			
			buf = buf & ")"
		End If
		
		WarCraftIII_ToString = buf
	End Function
	
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MakeArr(ByVal str_Renamed As String, ByRef arr() As Short)
		Dim i As Short
		
		ReDim arr(0)
		
		For i = 1 To Len(str_Renamed)
			If (i > 1) Then
				ReDim Preserve arr(i - 1)
			End If
			
			arr(i - 1) = Asc(Mid(str_Renamed, i, 1))
		Next i
	End Sub
	
	Private Function MakeRomanNum(ByVal Num As Short) As String
		Select Case (Num)
			Case 1 : MakeRomanNum = "I"
			Case 2 : MakeRomanNum = "II"
			Case 3 : MakeRomanNum = "III"
			Case 4 : MakeRomanNum = "IV"
			Case 5 : MakeRomanNum = "V"
		End Select
	End Function
End Class