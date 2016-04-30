VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsUserStats.cls
' Copyright (C) 2008 Eric Evans

Option Explicit

' variables
Private m_stat_string         As String
Private m_Game                As String
Private m_icon                As String
Private m_spawn               As Boolean
Private m_clan                As String
Private m_level               As Long
Private m_character_name      As String
Private m_character_class_id  As Integer
Private m_character_flags     As Long
Private m_acts_completed      As Long
Private m_character_ladder    As Boolean
Private m_wins                As Long
Private m_ladder_rating       As Long
Private m_high_rating         As Long
Private m_ladder_rank         As Long
Private m_strength            As Long
Private m_dexterity           As Long
Private m_vitality            As Long
Private m_gold                As Long
Private m_magic               As Long
Private m_dots                As Long
Private m_expansion           As Boolean
Private m_hardcore            As Boolean
Private m_realm               As String
Private m_are_stats_valid     As Boolean

Public Property Let Game(ByVal str As String)
    m_Game = str
End Property

Public Property Get Game() As String
    Game = m_Game
End Property

Public Property Let IsValid(ByVal bln As Boolean)
    m_are_stats_valid = bln
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_are_stats_valid
End Property

Public Property Get IsWCG() As Boolean
    Select Case (StrReverse$(Icon))
        Case "WCRF": IsWCG = True
        Case "WCPL": IsWCG = True
        Case "WCGO": IsWCG = True
        Case "WCSI": IsWCG = True
        Case "WCBR": IsWCG = True
        Case "WCPG": IsWCG = True
    End Select
End Property

Public Property Let Icon(ByVal str As String)
    m_icon = str
End Property

Public Property Get Icon() As String
    Icon = m_icon
End Property

Public Property Get IconTier() As String
    If ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP)) Then
        Select Case (Mid$(Icon, 2, 1))
            Case "H": IconTier = "Human"
            Case "N": IconTier = "Night Elf"
            Case "U": IconTier = "Undead"
            Case "O": IconTier = "Orc"
            Case "R": IconTier = "Random"
            Case "D": IconTier = "Tournament"
            Case Else: IconTier = Mid$(Icon, 2, 1)
        End Select
    End If
End Property

Public Property Get IconName() As String
    If ((Game = PRODUCT_WAR3) Or (Game = PRODUCT_W3XP)) Then
        Select Case (StrReverse$(Icon))
            Case "WCRF": IconName = "WCG Referee"
            Case "WCPL": IconName = "WCG Player"
            Case "WCGO": IconName = "WCG Gold Medalist"
            Case "WCSI": IconName = "WCG Silver Medalist"
            Case "WCBR": IconName = "WCG Bronze Medalist"
            Case "WCPG": IconName = "WCG Professional Gamer"
        End Select
        
        If (Len(IconName) > 0) Then
            Exit Property
        End If
    
        Select Case (Mid$(Icon, 2, 1))
            Case "H" ' Human
                If (Game = PRODUCT_WAR3) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "footman"
                        Case 3: IconName = "knight"
                        Case 4: IconName = "Archmage"
                        Case 5: IconName = "Medivh"
                    End Select
                Else
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "rifleman"
                        Case 3: IconName = "sorceress"
                        Case 4: IconName = "spellbreaker"
                        Case 5: IconName = "Blood Mage"
                        Case 6: IconName = "Jaina Proudmore"
                    End Select
                End If
            
            Case "N" ' Night Elf
                If (Game = PRODUCT_WAR3) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "archer"
                        Case 3: IconName = "druid of the claw"
                        Case 4: IconName = "Priestess of the Moon"
                        Case 5: IconName = "Furion"
                    End Select
                Else
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "huntress"
                        Case 3: IconName = "druid of the talon"
                        Case 4: IconName = "dryad"
                        Case 5: IconName = "Keeper of the Grove"
                        Case 6: IconName = "Maiev"
                    End Select
                End If
                
            Case "U" ' Undead
                If (Game = PRODUCT_WAR3) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "ghoul"
                        Case 3: IconName = "abomination"
                        Case 4: IconName = "Lich"
                        Case 5: IconName = "Tichondrius"
                    End Select
                Else
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "crypt fiend"
                        Case 3: IconName = "banshee"
                        Case 4: IconName = "destroyer"
                        Case 5: IconName = "Crypt Lord"
                        Case 6: IconName = "Sylvanas"
                    End Select
                End If
            
            Case "O" ' Orc
                If (Game = PRODUCT_WAR3) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "grunt"
                        Case 3: IconName = "tauren"
                        Case 4: IconName = "Far Seer"
                        Case 5: IconName = "Thrall"
                    End Select
                Else
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "headhunter"
                        Case 3: IconName = "shaman"
                        Case 4: IconName = "Spirit Walker"
                        Case 5: IconName = "Shadow Hunter"
                        Case 6: IconName = "Rexxar"
                    End Select
                End If

            Case "R" ' Random
                If (Game = PRODUCT_WAR3) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "dragon whelp"
                        Case 3: IconName = "blue dragon"
                        Case 4: IconName = "red dragon"
                        Case 5: IconName = "Deathwing"
                    End Select
                Else
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "myrmidon"
                        Case 3: IconName = "siren"
                        Case 4: IconName = "dragon turtle"
                        Case 5: IconName = "sea witch"
                        Case 6: IconName = "Illidan"
                    End Select
                End If

            Case "D" ' Tournament
                If (Game = PRODUCT_W3XP) Then
                    Select Case (Val(Mid(Icon, 1, 1)))
                        Case 1: IconName = "peon"
                        Case 2: IconName = "Felguard"
                        Case 3: IconName = "infernal"
                        Case 4: IconName = "doomguard"
                        Case 5: IconName = "pit lord"
                        Case 6: IconName = "Archimonde"
                    End Select
                End If
        End Select
        
        If (IconName = vbNullString) Then
            IconName = "unknown"
        End If
    End If
End Property

Public Property Let Statstring(ByVal str As String)
    If (Len(str) < 4) Then
        Exit Property
    End If
    Game = StrReverse$(Left$(str, 4))
    If (Len(str) >= 4) Then
        str = Mid$(str, 5)
        
        If (Left$(str, 1) = " ") Then
            str = Mid$(str, 2)
        End If
        
        m_stat_string = str
    
        Call Parse
    End If
End Property

Public Property Get Statstring() As String
    Statstring = m_stat_string
End Property

Public Property Get IsSpawn() As Boolean
    IsSpawn = m_spawn
End Property

Public Property Let IsSpawn(ByVal bln As Boolean)
    m_spawn = bln
End Property

Public Property Get Clan() As String
    Clan = m_clan
End Property

Public Property Let Clan(ByVal str As String)
    m_clan = str
End Property

Public Property Get Wins() As Long
    Wins = m_wins
End Property

Public Property Let Wins(ByVal lng As Long)
    m_wins = lng
End Property

Public Property Get LadderRating() As Long
    LadderRating = m_ladder_rating
End Property

Public Property Let LadderRating(ByVal lng As Long)
    m_ladder_rating = lng
End Property

Public Property Get HighLadderRating() As Long
    HighLadderRating = m_high_rating
End Property

Public Property Let HighLadderRating(ByVal lng As Long)
    m_high_rating = lng
End Property

Public Property Get LadderRank() As Long
    LadderRank = m_ladder_rank
End Property

Public Property Let LadderRank(ByVal lng As Long)
    m_ladder_rank = lng
End Property

Public Property Get Level() As Long
    Level = m_level
End Property

Public Property Let Level(ByVal lng As Long)
    m_level = lng
End Property

Public Property Get CharacterName() As String
    CharacterName = m_character_name
End Property

Public Property Let CharacterName(ByVal str As String)
    m_character_name = str
End Property

Public Property Get CharacterClassID() As Integer
    CharacterClassID = m_character_class_id
End Property

Public Property Let CharacterClassID(ByVal i As Integer)
    m_character_class_id = i
End Property

Public Property Get CharacterClass() As String
    On Error GoTo ERROR_HANDLER
    
    If ((Game = PRODUCT_DSHR) Or (Game = PRODUCT_DRTL)) Then
        Dim DIClasses(0 To 3) As String
        
        DIClasses(0) = "Warrior"
        DIClasses(1) = "Rogue"
        DIClasses(2) = "Sorceror"
        
        If (CharacterClassID < UBound(DIClasses)) Then
            CharacterClass = DIClasses(CharacterClassID)
        End If
        
    ElseIf ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
    
        Dim DIIClasses(0 To 7) As String
        
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
End Property

Public Property Let CharacterFlags(ByVal lng As Long)
    m_character_flags = lng
End Property

Public Property Get CharacterFlags() As Long
    CharacterFlags = m_character_flags
End Property

Public Property Get IsHardcoreCharacter() As Boolean
    IsHardcoreCharacter = ((CharacterFlags And &H4) = &H4)
End Property

Public Property Get IsFemaleCharacter() As Boolean
    IsFemaleCharacter = _
        ((m_character_class_id = 1) Or (m_character_class_id = 2) Or _
            (m_character_class_id = 7))
End Property

Public Property Get IsExpansionCharacter() As Boolean
    If (StrReverse$(Game) = PRODUCT_D2XP) Then
        IsExpansionCharacter = True
    Else
        IsExpansionCharacter = ((CharacterFlags And &H20) = &H20)
    End If
End Property

Public Property Get IsCharacterDead() As Boolean
    IsCharacterDead = ((IsHardcoreCharacter) And ((CharacterFlags And &H8) = &H8))
End Property

Public Property Let IsLadderCharacter(ByVal bln As Boolean)
    m_character_ladder = bln
End Property

Public Property Get IsLadderCharacter() As Boolean
    IsLadderCharacter = m_character_ladder
End Property

Public Property Get ActsCompleted() As Integer
    ActsCompleted = m_acts_completed
End Property

Public Property Let ActsCompleted(ByVal i As Integer)
    m_acts_completed = i
End Property

Public Property Get CurrentAct() As Integer
    If ((Game() = PRODUCT_D2DV) Or (Game() = PRODUCT_D2XP)) Then
        If (IsExpansionCharacter() = False) Then
            CurrentAct = ((ActsCompleted() Mod 4) + 1)
        Else
            CurrentAct = ((ActsCompleted() Mod 5) + 1)
        End If
    End If
End Property


Public Property Get CurrentActName() As String
    Select Case CurrentAct()
        Case 1: CurrentActName = "Act I"
        Case 2: CurrentActName = "Act II"
        Case 3: CurrentActName = "Act III"
        Case 4:
            If IsExpansionCharacter() Then
                CurrentActName = "Act IV/V" ' either IV or V
            Else
                CurrentActName = "Act IV"
            End If
        Case 5: CurrentActName = "Act V" ' this is never seen, but would be proper value for Act V
    End Select
End Property

Public Property Get CurrentActAndDifficulty() As String
    If CurrentDifficultyID() >= 3 Then
        CurrentActAndDifficulty = CurrentDifficulty()
    Else
        CurrentActAndDifficulty = CurrentActName() & " of " & CurrentDifficulty()
    End If
End Property

Public Property Get CurrentDifficulty() As String
    If ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
        Select Case (CurrentDifficultyID)
            Case 0: CurrentDifficulty = "Normal"
            Case 1: CurrentDifficulty = "Nightmare"
            Case 2: CurrentDifficulty = "Hell"
            Case 3: CurrentDifficulty = "All Acts"
        End Select
    End If
End Property

Public Property Get CurrentDifficultyID() As Integer
    Dim difficulty As Integer
    
    If ((Game = PRODUCT_D2DV) Or (Game = PRODUCT_D2XP)) Then
        If (IsExpansionCharacter = False) Then
            difficulty = Fix(ActsCompleted / 4)
        Else
            difficulty = Fix(ActsCompleted / 5)
        End If
    End If
    CurrentDifficultyID = difficulty
End Property

Public Property Get CharacterTitle() As String
    On Error GoTo ERROR_HANDLER

    ' thanks c0ol for multi-dimensional array idea
    Dim Classic(2, 3, 2)   As String
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
                CharacterTitle = _
                    Classic(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, IIf(IsFemaleCharacter, 1, 0))
            
                ' run again with is female as false
                If (CharacterTitle = vbNullString) Then
                    CharacterTitle = Classic(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, 0)
                End If
            Else
                CharacterTitle = _
                    Expansion(IIf(IsHardcoreCharacter, 1, 0), CurrentDifficultyID - 1, IIf(IsFemaleCharacter, 1, 0))
            
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
End Property

Public Property Get CharacterTitleAndName() As String
    CharacterTitleAndName = CharacterTitle()
    If (LenB(CharacterTitleAndName) > 0) Then CharacterTitleAndName = CharacterTitleAndName & " "
    CharacterTitleAndName = CharacterTitleAndName & CharacterName()
End Property

Public Property Get Dots() As Long
    Dots = m_dots
End Property

Public Property Let Dots(ByVal lng As Long)
    m_dots = lng
End Property

Public Property Get Strength() As Long
    Strength = m_strength
End Property

Public Property Let Strength(ByVal lng As Long)
    m_strength = lng
End Property

Public Property Get Magic() As Long
    Magic = m_magic
End Property

Public Property Let Magic(ByVal lng As Long)
    m_magic = lng
End Property

Public Property Get Gold() As Long
    Gold = m_gold
End Property

Public Property Let Gold(ByVal lng As Long)
    m_gold = lng
End Property

Public Property Get Dexterity() As Long
    Dexterity = m_dexterity
End Property

Public Property Let Dexterity(ByVal lng As Long)
    m_dexterity = lng
End Property

Public Property Get Vitality() As Long
    Vitality = m_vitality
End Property

Public Property Let Vitality(ByVal lng As Long)
    m_vitality = lng
End Property

Public Property Get Realm() As String
    Realm = m_realm
End Property

Public Property Let Realm(ByVal str As String)
    m_realm = str
End Property

Public Property Get ToString() As String
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
    
    ToString = buf
End Property

Public Property Get IconCode() As Integer
    Dim intWins As Long
    Dim strIcon As String
    
    If (BotVars.ShowStatsIcons) Then
        strIcon = StrReverse(Icon())
        If (IsValid And LenB(Statstring) > 0) Then
            Select Case strIcon
                Case vbNullString, Game: ' icon is current product or icon is not set
                Case Else: ' if icon field is present and custom
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
                    If (intWins > 10&) Then intWins = 10&
                    
                    ' choose starting point
                    If Game = PRODUCT_W2BN Then
                        IconCode = CInt(ICON_START_W2 + intWins)
                    Else
                        IconCode = CInt(ICON_START_SC + intWins)
                    End If
                    
                    ' if spawn flag, use spawn icon
                    If IsSpawn Then
                        Select Case Game
                            Case PRODUCT_W2BN: IconCode = IC_W2BN_SPAWN
                            Case PRODUCT_JSTR: IconCode = IC_JSTR_SPAWN
                            Case Else:         IconCode = IC_STAR_SPAWN
                        End Select
                    End If
                
            End Select
        End If
    End If
    
    If ((IconCode = 0) Or (IconCode = ICUNKNOWN)) Then
        IconCode = GetIconImageListPosition(Game())
    End If
End Property

Private Function GetIconImageListPosition(ByVal Icon As String) As Integer
    Dim IconCode As Integer
    Dim IMLSeq As Integer
    Dim PerTier As Integer
    
    IconCode = ICUNKNOWN
    
    Select Case Icon
        Case PRODUCT_CHAT: IconCode = ICCHAT
        Case PRODUCT_SSHR: IconCode = ICSCSW
        Case PRODUCT_STAR: IconCode = ICSTAR
        Case PRODUCT_JSTR: IconCode = ICJSTR
        Case PRODUCT_SEXP: IconCode = ICSEXP
        Case PRODUCT_DSHR: IconCode = ICDIABLOSW
        Case PRODUCT_DRTL: IconCode = ICDIABLO
        Case PRODUCT_W2BN: IconCode = ICW2BN
        Case PRODUCT_D2DV: IconCode = ICD2DV
        Case PRODUCT_D2XP: IconCode = ICD2XP
        Case PRODUCT_WAR3: IconCode = ICWAR3
        Case PRODUCT_W3XP: IconCode = ICWAR3X
        
        Case "WCRF": IconCode = IC_WCRF
        Case "WCPL": IconCode = IC_WCPL
        Case "WCGO": IconCode = IC_WCGO
        Case "WCSI": IconCode = IC_WCSI
        Case "WCBR": IconCode = IC_WCBR
        Case "WCPG": IconCode = IC_WCPG
        
        Case Else: IconCode = ICUNKNOWN
    End Select
    
    If IconCode <> ICUNKNOWN Then
        ' special named icon
        GetIconImageListPosition = IconCode
    ElseIf (StrComp(Left$(Icon, 2), "W3", vbBinaryCompare) = 0) Then
        ' W3 stats icon: choose starting point and per-tier
        If (Game = PRODUCT_WAR3) Then
            IMLSeq = ICON_START_WAR3
            PerTier = 5
        Else
            IMLSeq = ICON_START_W3XP
            PerTier = 6
        End If
        
        ' choose race
        Select Case Mid$(Icon, 3, 1)
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
        If StrictIsNumeric(Mid$(Icon, 4, 1)) And IconCode > IMLSeq Then
            IconCode = IconCode + (CInt(Mid$(Icon, 4, 1)) - 1)
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
            Call ParseStarCraft
            
        Case PRODUCT_DSHR, PRODUCT_DRTL
            Call ParseDiablo
            
        Case PRODUCT_D2DV, PRODUCT_D2XP
            Call ParseDiabloII

        Case PRODUCT_W2BN
            Call ParseWarCraftII
        
        Case PRODUCT_WAR3, PRODUCT_W3XP
            Call ParseWarCraftIII
    End Select
    
    Exit Sub
    
ERROR_HANDLER:
    IsValid = False
    
    Exit Sub
End Sub

Private Sub ParseStarCraft()
    Dim Values() As String
    Dim lngIsSpawn As Long
    
    If (Statstring = vbNullString) Then
        Exit Sub
    End If
    
    IsValid = True
    
    ReDim Preserve Values(0)
    Values() = Split(Statstring, Space$(1))
    
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
    Dim lngIsSpawn As Long
    
    If (Statstring = vbNullString) Then
        Exit Sub
    End If
    
    IsValid = True
    
    ReDim Preserve Values(0)
    Values() = Split(Statstring, Space$(1))
    
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
    Dim Values()   As String
    Dim charData() As Integer
    
    ' empty D2 stats is valid ("Open character")
    IsValid = True
    
    If (Statstring = vbNullString) Then
        Exit Sub
    End If
    
    ReDim Preserve Values(0)
    Values() = Split(Statstring, ",", 3)
  
    If (UBound(Values) >= 1) Then
        Realm = Values(0)
        CharacterName = Values(1)
        
        If (UBound(Values) >= 2) Then
            MakeArr Values(2), charData
            
            If (UBound(charData) >= 13) Then
                CharacterClassID = charData(13)
                If (UBound(charData) >= 27) Then
                    Level = charData(25)
                    CharacterFlags = charData(26)
                    ActsCompleted = ((charData(27) Xor &H80) / 2)
    
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
    Call ParseStarCraft
End Sub

Private Sub ParseWarCraftIII()
    Dim Values() As String
    
    If (Statstring = vbNullString) Then
        Exit Sub
    End If
    
    IsValid = True
    
    ReDim Preserve Values(0)
    Values() = Split(Statstring, Space$(1))
    
    If (UBound(Values) = 0) Then
        Level = Val(Values(0))
        Exit Sub
    End If
    
    Level = Val(Values(1))
    Icon = Values(0)
    
    If (UBound(Values) > 1) Then
        Clan = StrReverse$(Values(2))
    End If
End Sub

Private Function StarCraft_ToString() As String
    Dim buf As String
    
    If (Statstring = vbNullString) Then
        buf = " (No stats available)"
    ElseIf (IsValid = False) Then
        buf = " (Unrecognized stats)"
    Else
        buf = _
            " (" & Wins() & " wins" & IIf(LadderRating(), ", with a rating of " & LadderRating() & _
                " on the ladder", "") & ")"
        
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
        buf = _
            " (Level " & Level() & " " & CharacterClass() & " with " & Dots() & " dots, " & Strength() & _
                " strength, " & Magic() & " magic, " & Dexterity() & " dexterity, " & Vitality() & _
                    " vitality, and " & Gold() & " gold)"
        
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

        buf = _
            buf & CharacterTitleAndName() & ", a " & IIf(IsCharacterDead(), "dead ", vbNullString) & _
                IIf(IsHardcoreCharacter(), "hardcore ", vbNullString) & "level " & Level() & " " & _
                    IIf(IsLadderCharacter(), "ladder ", vbNullString) & CharacterClass()
        
        buf = _
            buf & " on Realm " & Realm() & ")"
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

Private Sub MakeArr(ByVal str As String, ByRef arr() As Integer)
    Dim i As Integer
    
    ReDim arr(0)
    
    For i = 1 To Len(str)
        If (i > 1) Then
            ReDim Preserve arr(0 To i - 1)
        End If
        
        arr(i - 1) = Asc(Mid(str, i, 1))
    Next i
End Sub

Private Function MakeRomanNum(ByVal Num As Integer) As String
    Select Case (Num)
        Case 1: MakeRomanNum = "I"
        Case 2: MakeRomanNum = "II"
        Case 3: MakeRomanNum = "III"
        Case 4: MakeRomanNum = "IV"
        Case 5: MakeRomanNum = "V"
    End Select
End Function
